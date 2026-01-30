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
// 4. API HELPERS
// ==========================================

function fetchShopifyOrdersForExport_(startDate) {
  const endpoint = `https://${SHOPIFY_EXPORT_CONFIG.domain}/admin/api/${SHOPIFY_EXPORT_CONFIG.apiVersion}/orders.json` +
    `?status=any` + 
    `&created_at_min=${startDate.toISOString()}` +
    `&fields=id,name,created_at,total_price,total_shipping_price_set,line_items,financial_status` + 
    `&limit=250`; 

  const allItems = [];
  let nextUrl = endpoint;

  const options = {
    method: 'get',
    muteHttpExceptions: true,
    headers: { 'X-Shopify-Access-Token': SHOPIFY_EXPORT_CONFIG.accessToken }
  };

  while (nextUrl) {
    // Use Shared Utility
    const content = fetchUrlWithRetry(nextUrl, options);
    
    if (content) {
      const dataKey = Object.keys(content)[0]; 
      if (content[dataKey]) {
        allItems.push(...content[dataKey]);
      }

      // Pagination is slightly tricky to abstract purely generically without callback, so we keep logic here but use helper for fetch
      // We need to re-fetch the headers to get the 'Link' -> UrlFetchApp.fetch returns a HttpResponse, 
      // but our helper returns JSON. This is a trade-off. 
      // ACTUALLY, the helper parsing JSON swallows the Headers. 
      // To properly paginate Shopify, we need headers.
      // So I will REVERT using `fetchUrlWithRetry` inside the loop FOR SHOPIFY if it needs headers, 
      // OR update `fetchUrlWithRetry` to return both object and headers?
      // Simpler: Keep the loop logic here but use the Retry pattern logic found in Utilities? 
      // NO, better to modify Utilities to support full response return if needed, or just specific to Shopify pagination.
      // 
      // FOR NOW: Stick to local loop for Shopify because of the 'Link' header requirement which JSON parsing hides.
      // I will implement a local retry loop using the standard pattern but cleaned up.
      
      // WAIT: I can just use UrlFetchApp directly with a helper for the retry wrapper?
      // Let's optimize: I'll use a local retry block here similar to before but cleaner.
      // The shared utility `fetchUrlWithRetry` returns JSON content, losing headers. 
      // So for Shopify pagination, I'll keep the specialized loop but ensure it matches the robustness.
    }
    
    // RE-IMPLEMENTING PAGINATION LOOP EFFICIENTLY
    // Note: The previous implementation was already good. I'll just clean it up.
    // Actually, I can use a slightly modified helper or just keep it local.
    // Given the constraints, I'll keep the robust local loop in `fetchShopifyDataWithPaginationExport_` (renamed/inlined)
    // but simplified.
    break; // Placeholder break to avoid infinite loop in this thought process block.
  }
  
  // REAL IMPLEMENTATION BELOW
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
    let response;
    let successful = false;
    
    for (let i = 0; i < 3; i++) { 
      try {
        response = UrlFetchApp.fetch(nextUrl, options);
        const responseCode = response.getResponseCode();
        
        if (responseCode === 429) { 
          const retryAfter = response.getHeaders()['Retry-After'] || 2;
          Utilities.sleep(parseInt(retryAfter) * 1000);
          continue; 
        }
        
        if (responseCode >= 200 && responseCode < 300) {
          successful = true;
          break;
        } else {
           Logger.log(`[Shopify API Error] ${responseCode}`);
           break; // Don't retry non-transient errors blindly
        }
      } catch (e) {
        Utilities.sleep(1000); 
      }
    }

    if (successful && response) {
       const content = JSON.parse(response.getContentText());
       const dataKey = Object.keys(content)[0];
       if (content[dataKey]) {
         allItems.push(...content[dataKey]);
       }

       const linkHeader = response.getHeaders()['Link'];
       const links = linkHeader ? linkHeader.split(',') : [];
       const nextLink = links.find(link => link.includes('rel="next"'));
       nextUrl = nextLink ? nextLink.match(/<([^>]+)>/)[1] : null;
    } else {
       break; // Stop if failed
    }
  }
  return allItems;
}
