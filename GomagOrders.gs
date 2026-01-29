/**
 * @file GomagOrders.gs
 * @description Exports detailed Gomag order data to "GomagOrders" sheet.
 */

// ==========================================
// 1. CONFIGURATION
// ==========================================

const GOMAG_EXPORT_SHEET_NAME = "GomagOrders";

// Internal config object (populated from Sheet)
let GOMAG_EXPORT_CONFIG = {
  shopUrl: '', 
  apiKey: '', 
  timeframeDays: 30
};

// ==========================================
// 2. MAIN FUNCTION
// ==========================================

function runGomagOrderExport() {
  // 1. Load Configuration
  try {
    const config = fetchConfigFromSheet();
    if (config.gomagDomain) GOMAG_EXPORT_CONFIG.shopUrl = config.gomagDomain;
    if (config.gomagAPIkey) GOMAG_EXPORT_CONFIG.apiKey = config.gomagAPIkey;
    
    // Validation
    if (!GOMAG_EXPORT_CONFIG.shopUrl || !GOMAG_EXPORT_CONFIG.apiKey) {
       throw new Error("Missing Gomag credentials in Config sheet.");
    }
  } catch (e) {
    Logger.log(`[GomagExport] Config Error: ${e.message}`);
    const ui = SpreadsheetApp.getUi();
    ui.alert(`Config Error: ${e.message}`);
    return;
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 2. Calculate Timeframe
  const endDate = new Date();
  const startDate = new Date();
  startDate.setDate(endDate.getDate() - GOMAG_EXPORT_CONFIG.timeframeDays);
  startDate.setHours(0, 0, 0, 0); 
  
  Logger.log(`[GomagExport] Starting export from ${startDate.toISOString()}...`);

  // 3. Fetch Data
  try {
    const data = fetchGomagOrdersForExport_(startDate);
    
    // 4. Process Data
    const processedRows = data.orders.map(processGomagOrderForExport_);

    Logger.log(`[GomagExport] Processed ${processedRows.length} orders.`);

    // 5. Write to Sheet
    writeGomagExportToSheet_(spreadsheet, processedRows);

  } catch (e) {
    Logger.log(`[GomagExport] ERROR: ${e.message}`);
    const ui = SpreadsheetApp.getUi();
    ui.alert(`Error: ${e.message}`);
  }
}

// ==========================================
// 3. DATA PROCESSING
// ==========================================

function processGomagOrderForExport_(order) {
  // 1. Date
  const dateStr = order.date;
  const date = dateStr ? new Date(dateStr) : "";
  
  // 2. ID tranzactie
  const transactionId = order.id || "";
  
  // 3. Value (Total Price)
  let value = 0;
  if (order.total) {
    let priceRaw = order.total;
     if (typeof priceRaw === 'string') {
          priceRaw = priceRaw.replace(/[^0-9.-]+/g,"");
      }
      value = parseFloat(priceRaw);
      if (isNaN(value)) value = 0;
  }
  
  // 4. Transport (Shipping)
  let transport = 0;
  if (order.shipping_value) {
     let shipRaw = order.shipping_value;
     if (typeof shipRaw === 'string') {
          shipRaw = shipRaw.replace(/[^0-9.-]+/g,"");
      }
      transport = parseFloat(shipRaw);
      if (isNaN(transport)) transport = 0;
  }

  // 5. Venituri items (Sum of price * quantity)
  let itemsRevenue = 0;
  let productIds = [];
  
  // Gomag returns 'products' as an array for items
  // Sometimes it can be an object with numeric keys, careful handling needed similar to fetch logic if API varies,
  // but typically 'products' inside an order object is an array in the standard V1 response.
  let items = order.products || [];
  
  // Normalize if it's an object acting as array
  if (!Array.isArray(items) && typeof items === 'object') {
     items = Object.values(items);
  }

  if (Array.isArray(items)) {
    items.forEach(item => {
      let price = 0;
      if (item.price) { // Assuming 'price' is the unit price field
          let pRaw = item.price;
          if (typeof pRaw === 'string') pRaw = pRaw.replace(/[^0-9.-]+/g,"");
          price = parseFloat(pRaw) || 0;
      }

      let quantity = 0;
      if (item.quantity) {
          quantity = parseInt(item.quantity) || 0;
      }
      
      itemsRevenue += (price * quantity);
      
      if (item.id) productIds.push(item.id);
      else if (item.code) productIds.push(item.code);
    });
  }
  
  // 6. ID products (Joined string)
  const productsIdStr = productIds.join(", ");
  
  // 7. Status
  const status = order.status || "";

  // Return row order: Date | ID tranzactie | Value | Transport | Venituri items | ID products | Status
  return [
    date,
    transactionId,
    value,
    transport,
    itemsRevenue,
    productsIdStr,
    status
  ];
}

// ==========================================
// 4. API & SHEET HELPERS
// ==========================================

function writeGomagExportToSheet_(spreadsheet, rows) {
  let sheet = spreadsheet.getSheetByName(GOMAG_EXPORT_SHEET_NAME);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(GOMAG_EXPORT_SHEET_NAME);
  } else {
    sheet.clear(); 
  }

  const headers = ["Date", "ID tranzactie", "Value", "Transport", "Venituri items", "ID products", "Status"];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#d9ead3");

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    // Format numeric columns
    sheet.getRange(2, 3, rows.length, 3).setNumberFormat("#,##0.00"); // Value, Transport, Venituri items
  }
  
  sheet.autoResizeColumns(1, headers.length);
}

// ==========================================
// 5. CORE ENGINE (PAGINATION)
// ==========================================

function fetchGomagOrdersForExport_(startDate) {
  const allOrders = [];
  let page = 1;
  let keepFetching = true;

  while (keepFetching) {
    // We need more details than a simple list, but the '/read/json' endpoint usually returns full objects or a summary.
    // Gomag API usually returns full order objects in the list.
    const url = `https://api.gomag.ro/api/v1/order/read/json?limit=50&page=${page}`;
    
    Logger.log(`[GomagExport] Fetching page ${page}...`);
    
    const response = fetchGomagUrlExport_(url);

    if (!response) {
      Logger.log("[GomagExport] Null response. Stopping.");
      break;
    }

    let ordersBatch = [];

    // Robust parsing logic copied from GomagOrderTotals.gs, slightly simplified for clarity but keeping core safety
    if (Array.isArray(response)) {
      ordersBatch = response;
    } else if (response.orders) {
      const val = response.orders;
      ordersBatch = Array.isArray(val) ? val : Object.values(val);
    } else if (response.data) {
       if (Array.isArray(response.data)) {
         ordersBatch = response.data;
       } else if (response.data.orders) {
         const val = response.data.orders;
         ordersBatch = Array.isArray(val) ? val : Object.values(val);
       }
    } 
    // Fallback for object keyed by ID
    else if (typeof response === 'object') {
       const vals = Object.values(response);
       if (vals.length > 0 && vals[0] && typeof vals[0] === 'object' && vals[0].date) {
         ordersBatch = vals;
       }
    }

    if (!ordersBatch || ordersBatch.length === 0) {
      break;
    }

    let ordersInBatchAdded = 0;
    for (const order of ordersBatch) {
      const orderDateStr = order.date;
      if (!orderDateStr) continue; 

      const orderDate = new Date(orderDateStr);
      
      if (orderDate < startDate) {
        keepFetching = false;
        break; 
      }
      
      allOrders.push(order);
      ordersInBatchAdded++;
    }

    // Edge case where batch has orders but all are older than start date (and loop broke above)
    // or checks to ensure we don't loop endlessly if date isn't sorted perfectly (Gomag usually sorts desc)
    
    if (keepFetching) {
      page++;
      // Safety break to prevent infinite loops if something is wrong
      if (page > 100) keepFetching = false; 
    }
  }

  return { orders: allOrders };
}

function fetchGomagUrlExport_(url) {
  const options = {
    method: 'get',
    muteHttpExceptions: true,
    headers: {
      'Apikey': GOMAG_EXPORT_CONFIG.apiKey,
      'ApiShop': GOMAG_EXPORT_CONFIG.shopUrl,
      'Content-Type': 'application/json'
    }
  };

  for (let i = 0; i < 3; i++) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      const contentText = response.getContentText();
      
      if (code === 200) {
        return JSON.parse(contentText);
      } else if (code === 429) {
        Utilities.sleep(5000);
        continue;
      } else {
        Logger.log(`[Exports API ERROR] Code: ${code}. Response: ${contentText}`);
        return null; 
      }
    } catch (e) {
      if (i === 2) {
        Logger.log(`[Exports FETCH ERROR] ${e.message}`);
        return null;
      }
      Utilities.sleep(2000);
    }
  }
  return null;
}
