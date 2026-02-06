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
  const { startDate, endDate } = calculateDateRange(SHOPIFY_EXPORT_CONFIG.timeframeDays);
  
  Logger.log(`[ShopifyExport] Starting export from ${startDate.toISOString()} to ${endDate.toISOString()}...`);

  // 3. Fetch Data
  try {
    const orders = fetchShopifyOrdersForExport_(startDate, endDate);
    
    // 4. Process Data
    const processedRows = orders.map(processOrderForExport_);

    Logger.log(`[ShopifyExport] Processed ${processedRows.length} orders.`);

    // 5. Write to Sheet
    const headers = ["Date", "Transaction ID", "Value", "Shipping", "Items Revenue", "Product IDs", "Status"];
    writeParamsToSheet(spreadsheet, SHOPIFY_EXPORT_SHEET_NAME, headers, processedRows);
    
    // Format numeric columns specific to this report (cols 3, 4, 5)
    const sheet = spreadsheet.getSheetByName(SHOPIFY_EXPORT_SHEET_NAME);
    if (sheet && processedRows.length > 0) {
      sheet.getRange(2, 3, processedRows.length, 3).setNumberFormat("#,##0.00");
    }

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
  const transactionId = order.name ? order.name.replace('#', '') : "";
  
  // 3. Value (Total Price)
  const value = parseFloat(order.total_price) || 0;
  
  // 4. Transport (Shipping)
  let transport = 0;
  // Priority 1: total_shipping_price_set (Modern, includes tax/discounts properly)
  if (order.total_shipping_price_set && 
      order.total_shipping_price_set.shop_money && 
      order.total_shipping_price_set.shop_money.amount) {
    transport = parseFloat(order.total_shipping_price_set.shop_money.amount);
  } 
  // Priority 2: total_shipping_price (Legacy field)
  else if (order.total_shipping_price) {
    transport = parseFloat(order.total_shipping_price);
  }
  // Priority 3: Sum of shipping_lines (Manual fallback)
  else if (order.shipping_lines && Array.isArray(order.shipping_lines)) {
    order.shipping_lines.forEach(line => {
      // Use discounted_price if available, else price
      let linePrice = line.discounted_price || line.price;
      transport += parseFloat(linePrice) || 0;
    });
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
// 4. API HELPERS
// ==========================================

// ==========================================
// 4. API HELPERS
// ==========================================

function fetchShopifyOrdersForExport_(startDate, endDate) {
  const endpoint = `https://${SHOPIFY_EXPORT_CONFIG.domain}/admin/api/${SHOPIFY_EXPORT_CONFIG.apiVersion}/orders.json` +
    `?status=any` + 
    `&created_at_min=${startDate.toISOString()}` +
    `&created_at_max=${endDate.toISOString()}` +
    `&fields=id,name,created_at,total_price,total_shipping_price_set,line_items,financial_status` + 
    `&limit=250`; 

  return fetchShopifyDataWithPaginationExport_(endpoint, SHOPIFY_EXPORT_CONFIG.accessToken);
}

function fetchShopifyDataWithPaginationExport_(initialUrl, accessToken) {
  const allItems = [];
  let nextUrl = initialUrl;

  const options = {
    method: 'get',
    muteHttpExceptions: true,
    headers: { 'X-Shopify-Access-Token': accessToken }
  };

  while (nextUrl) {
    // Request full response (content + headers) for pagination
    const response = fetchUrlWithRetry(nextUrl, options, true);

    if (response && response.content) {
       const content = response.content;
       const dataKey = Object.keys(content)[0];
       if (content[dataKey]) {
         allItems.push(...content[dataKey]);
       }

       // Parse 'Link' header for next page
       const linkHeader = response.headers['Link'];
       const links = linkHeader ? linkHeader.split(',') : [];
       const nextLink = links.find(link => link.includes('rel="next"'));
       nextUrl = nextLink ? nextLink.match(/<([^>]+)>/)[1] : null;

    } else {
       Logger.log("[ShopifyExport] Stop fetching due to error or null response.");
       break; 
    }
  }
  return allItems;
}
