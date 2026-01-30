/**
 * @file ShopifyOrders.gs
 * @description Exports detailed Shopify order data to "ShopifyOrders" sheet.
 */

// ==========================================
// 1. CONFIGURATION
// ==========================================

const SHOPIFY_EXPORT_SHEET_NAME = "ShopifyOrders";

// Internal config object (populated from Sheet)
let SHOPIFY_EXPORT_CONFIG = {
  domain: '', 
  accessToken: '', 
  timeframeDays: 30, 
  apiVersion: '2024-04' 
};

// ==========================================
// 2. MAIN FUNCTION
// ==========================================

function runShopifyOrderExport() {
  // 1. Load Configuration
  try {
    const config = fetchConfigFromSheet();
    if (config.shopifyDomain) SHOPIFY_EXPORT_CONFIG.domain = config.shopifyDomain;
    if (config.shopifyAPIkey) SHOPIFY_EXPORT_CONFIG.accessToken = config.shopifyAPIkey;
    
    if (config.Days) {
      const days = parseInt(config.Days);
      if (!isNaN(days) && days > 0) {
        SHOPIFY_EXPORT_CONFIG.timeframeDays = days;
      }
    }
    
    // Validation
    if (!SHOPIFY_EXPORT_CONFIG.domain || !SHOPIFY_EXPORT_CONFIG.accessToken) {
       throw new Error("Missing Shopify credentials in Config sheet.");
    }
  } catch (e) {
    Logger.log(`[ShopifyExport] Config Error: ${e.message}`);
    const ui = SpreadsheetApp.getUi();
    ui.alert(`Config Error: ${e.message}`);
    return;
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 2. Calculate Timeframe
  const endDate = new Date();
  const startDate = new Date();
  startDate.setDate(endDate.getDate() - SHOPIFY_EXPORT_CONFIG.timeframeDays);
  
  Logger.log(`[ShopifyExport] Starting export from ${startDate.toISOString()}...`);

  // 3. Fetch Data
  try {
    const orders = fetchShopifyOrdersForExport_(startDate);
    
    // 4. Process Data
    const processedRows = orders.map(processOrderForExport_);

    Logger.log(`[ShopifyExport] Processed ${processedRows.length} orders.`);

    // 5. Write to Sheet
    writeShopifyExportToSheet_(spreadsheet, processedRows);

  } catch (e) {
    Logger.log(`[ShopifyExport] ERROR: ${e.message}`);
    const ui = SpreadsheetApp.getUi();
    ui.alert(`Error: ${e.message}`);
  }
}

// ==========================================
// 3. DATA PROCESSING
// ==========================================

function processOrderForExport_(order) {
  // 1. Date
  const date = new Date(order.created_at);
  
  // 2. ID tranzactie (Name e.g. #1001)
  const transactionId = order.name;
  
  // 3. Value (Total Price)
  const value = parseFloat(order.total_price) || 0;
  
  // 4. Transport (Shipping)
  let transport = 0;
  if (order.total_shipping_price_set && 
      order.total_shipping_price_set.shop_money && 
      order.total_shipping_price_set.shop_money.amount) {
    transport = parseFloat(order.total_shipping_price_set.shop_money.amount);
  }

  // 5. Venituri items (Sum of price * quantity)
  let itemsRevenue = 0;
  let productIds = [];
  
  if (order.line_items && Array.isArray(order.line_items)) {
    order.line_items.forEach(item => {
      const price = parseFloat(item.price) || 0;
      const quantity = parseInt(item.quantity) || 0;
      itemsRevenue += (price * quantity);
      
      if (item.product_id) {
        productIds.push(item.product_id);
      }
    });
  }
  
  // 6. ID products (Joined string)
  const productsIdStr = productIds.join(", ");
  
  // 7. Status
  const status = order.financial_status;

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

function fetchShopifyOrdersForExport_(startDate) {
  const endpoint = `https://${SHOPIFY_EXPORT_CONFIG.domain}/admin/api/${SHOPIFY_EXPORT_CONFIG.apiVersion}/orders.json` +
    `?status=any` + 
    `&created_at_min=${startDate.toISOString()}` +
    `&fields=id,name,created_at,total_price,total_shipping_price_set,line_items,financial_status` + 
    `&limit=250`; 

  return fetchShopifyDataWithPaginationExport_(endpoint, SHOPIFY_EXPORT_CONFIG.accessToken);
}

function writeShopifyExportToSheet_(spreadsheet, rows) {
  let sheet = spreadsheet.getSheetByName(SHOPIFY_EXPORT_SHEET_NAME);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHOPIFY_EXPORT_SHEET_NAME);
  } else {
    sheet.clear(); 
  }

  const headers = ["Date", "ID tranzactie", "Value", "Transport", "Venituri items", "ID products", "Status"];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#d9ead3");

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    // Format numeric columns if needed
    sheet.getRange(2, 3, rows.length, 3).setNumberFormat("#,##0.00"); // Value, Transport, Venituri items
  }
  
  sheet.autoResizeColumns(1, headers.length);
}

// ==========================================
// 5. CORE ENGINE (PAGINATION)
// ==========================================

function fetchShopifyDataWithPaginationExport_(initialUrl, accessToken) {
  const allItems = [];
  let nextUrl = initialUrl;

  const options = {
    method: 'get',
    muteHttpExceptions: true,
    headers: { 'X-Shopify-Access-Token': accessToken }
  };

  while (nextUrl) {
    let response;
    for (let i = 0; i < 3; i++) { 
      try {
        response = UrlFetchApp.fetch(nextUrl, options);
        const responseCode = response.getResponseCode();
        
        if (responseCode === 429) { 
          const retryAfter = response.getHeaders()['Retry-After'] || 2;
          Logger.log(`[Export] Rate limit hit. Sleeping for ${retryAfter}s...`);
          Utilities.sleep(parseInt(retryAfter) * 1000);
          continue; 
        }
        
        if (responseCode >= 200 && responseCode < 300) {
          const content = JSON.parse(response.getContentText());
          const dataKey = Object.keys(content)[0]; 
          if (content[dataKey]) {
            allItems.push(...content[dataKey]);
          }

          const linkHeader = response.getHeaders()['Link'];
          const links = linkHeader ? linkHeader.split(',') : [];
          const nextLink = links.find(link => link.includes('rel="next"'));
          nextUrl = nextLink ? nextLink.match(/<([^>]+)>/)[1] : null;
          break; 
          
        } else {
          throw new Error(`API Error: ${responseCode} - ${response.getContentText()}`);
        }
      } catch (e) {
        if (i === 2) throw e; 
        Utilities.sleep(1000); 
      }
    }
  }
  return allItems;
}
