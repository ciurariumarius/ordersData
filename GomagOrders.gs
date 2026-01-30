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
    
    if (config.Days) {
      const days = parseInt(config.Days);
      if (!isNaN(days) && days > 0) {
        GOMAG_EXPORT_CONFIG.timeframeDays = days;
      }
    }
    
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
  const dateStr = order.date || order.created_at || "";
  const date = dateStr ? new Date(dateStr) : "";
  
  // 2. ID tranzactie
  const transactionId = order.id || order.increment_id || "";
  
  // 3. Value (Total Price)
  let value = 0;
  // Check 'total', 'grand_total'
  let priceRaw = order.total || order.grand_total; 
  if (priceRaw) {
     if (typeof priceRaw === 'string') {
          priceRaw = priceRaw.replace(/[^0-9.-]+/g,"");
      }
      value = parseFloat(priceRaw);
      if (isNaN(value)) value = 0;
  }
  
  // 4. Transport (Shipping)
  let transport = 0;
  // Check 'shipping_value', 'shipping_amount', 'shipping_total' at root
  let shipRaw = order.shipping_value || order.shipping_amount || order.shipping_total;
  
  // If not at root, check inside 'shipping' object
  if (!shipRaw && order.shipping && typeof order.shipping === 'object') {
     // Logging for debugging shipping structure
     // if (!DEBUG_LOGGED) { // utilizing existing flag if possible, or just log first non-null
     //   Logger.log("[DEBUG SHIPPING OBJ]: " + JSON.stringify(order.shipping));
     // }
      Logger.log("[DEBUG SHIPPING OBJ]: " + JSON.stringify(order.shipping)); // Always log for now (or throttle)
     
     shipRaw = order.shipping.value || order.shipping.amount || order.shipping.cost || order.shipping.total || order.shipping.price;
  }
  
  if (shipRaw) {
     if (typeof shipRaw === 'string') {
          shipRaw = shipRaw.replace(/[^0-9.-]+/g,"");
      }
      transport = parseFloat(shipRaw);
      if (isNaN(transport)) transport = 0;
  }

  // 5. Venituri items & 6. ID products
  let itemsRevenue = 0;
  let productIds = [];
  
  // Try to find the item list in various common fields
  let items = order.products || order.items || order.line_items || order.lines || (order.details ? order.details.products : null);
  
  // Normalize object-as-list
  if (items && !Array.isArray(items) && typeof items === 'object') {
     items = Object.values(items);
  }

  if (Array.isArray(items) && items.length > 0) {
    items.forEach(item => {
      // DEBUG: Log item structure for the first valid item found
      // if (productIds.length === 0 && DEBUG_LOGGED) { 
      //    Logger.log("[DEBUG ITEM] Item Keys: " + Object.keys(item).join(", "));
      // }

      let price = 0;
      // Check 'price', 'unit_price', 'value'
      let pRaw = item.price || item.unit_price || item.value;
      if (pRaw) {
          if (typeof pRaw === 'string') pRaw = pRaw.replace(/[^0-9.-]+/g,"");
          price = parseFloat(pRaw) || 0;
      }

      let quantity = 0;
      // Check 'quantity', 'qty', 'count'
      let qRaw = item.quantity || item.qty || item.count || item.pieces;
      if (qRaw) {
          quantity = parseInt(qRaw) || 0;
      }
      
      itemsRevenue += (price * quantity);
      
      // Check ID fields: 'id', 'product_id', 'code', 'sku'
      if (item.id) productIds.push(item.id);
      else if (item.product_id) productIds.push(item.product_id);
      else if (item.code) productIds.push(item.code);
      else if (item.sku) productIds.push(item.sku);
    });
  } else {
    // Log warning only if it's not a status-only update or similar partial object
    // Logger.log(`[DEBUG] No items found for order ${transactionId}`);
  }
  
  // 6. ID products (Joined string)
  const productsIdStr = productIds.join(", ");
  
  // 7. Status
  const status = order.status || order.state || "";

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
