/**
 * @file WooCommerceOrders.gs
 * @description Exports detailed WooCommerce order data to "WooCommerceOrders" sheet.
 */

// ==========================================
// 1. CONFIGURATION
// ==========================================

const WOO_EXPORT_SHEET_NAME = "WooCommerceOrders";

// Internal config object (populated from Sheet)
let WOO_EXPORT_CONFIG = {
  url: '', 
  consumerKey: '', 
  consumerSecret: '',
  timeframeDays: 30
};

// ==========================================
// 2. MAIN FUNCTION
// ==========================================

function runWooCommerceOrderExport() {
  // 1. Load Configuration
  try {
    const config = fetchConfigFromSheet();
    if (config.wooUrl) WOO_EXPORT_CONFIG.url = config.wooUrl.replace(/\/$/, ""); // Remove trailing slash
    if (config.wooConsumerKey) WOO_EXPORT_CONFIG.consumerKey = config.wooConsumerKey;
    if (config.wooConsumerSecret) WOO_EXPORT_CONFIG.consumerSecret = config.wooConsumerSecret;
    
    if (config.Days) {
      const days = parseInt(config.Days);
      if (!isNaN(days) && days > 0) {
        WOO_EXPORT_CONFIG.timeframeDays = days;
      }
    }
    
    // Validation
    if (!WOO_EXPORT_CONFIG.url || !WOO_EXPORT_CONFIG.consumerKey || !WOO_EXPORT_CONFIG.consumerSecret) {
       throw new Error("Missing WooCommerce credentials in Config sheet (wooUrl, wooConsumerKey, wooConsumerSecret).");
    }
    
    // Ensure URL has protocol
    if (!WOO_EXPORT_CONFIG.url.startsWith("http")) {
      WOO_EXPORT_CONFIG.url = "https://" + WOO_EXPORT_CONFIG.url;
    }

  } catch (e) {
    Logger.log(`[WooExport] Config Error: ${e.message}`);
    const ui = SpreadsheetApp.getUi();
    ui.alert(`Config Error: ${e.message}`);
    return;
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 2. Calculate Timeframe
  const { startDate, endDate } = calculateDateRange(WOO_EXPORT_CONFIG.timeframeDays);
  
  Logger.log(`[WooExport] Starting export from ${startDate.toISOString()} to ${endDate.toISOString()}...`);

  // 3. Fetch Data
  try {
    const orders = fetchWooOrdersForExport_(startDate, endDate);
    
    // 4. Process Data
    const processedRows = orders.map(processWooOrderForExport_);

    Logger.log(`[WooExport] Processed ${processedRows.length} orders.`);

    // 5. Write to Sheet
    const headers = ["Date", "Transaction ID", "Value", "Shipping", "Items Revenue", "Product IDs", "Status"];
    writeParamsToSheet(spreadsheet, WOO_EXPORT_SHEET_NAME, headers, processedRows);
    
    // Format numeric columns specific to this report (cols 3, 4, 5)
    const sheet = spreadsheet.getSheetByName(WOO_EXPORT_SHEET_NAME);
    if (sheet && processedRows.length > 0) {
      sheet.getRange(2, 3, processedRows.length, 3).setNumberFormat("#,##0.00");
    }

  } catch (e) {
    Logger.log(`[WooExport] ERROR: ${e.message}`);
    const ui = SpreadsheetApp.getUi();
    ui.alert(`Error: ${e.message}`);
  }
}

// ==========================================
// 3. DATA PROCESSING
// ==========================================

function processWooOrderForExport_(order) {
  // 1. Date
  const dateStr = order.date_created || order.date_created_gmt;
  const date = dateStr ? new Date(dateStr) : "";
  
  // 2. ID tranzactie
  const transactionId = order.id ? String(order.id) : "";
  
  // 3. Value (Total Price)
  const value = parseFloat(order.total) || 0;
  
  // 4. Transport (Shipping)
  const transport = parseFloat(order.shipping_total) || 0;

  // 5. Venituri items (Sum of price * quantity)
  // Note: WooCommerce 'line_items' usually have 'total' (row total excluding tax usually, or including depending on settings)
  // but 'total' + 'total_tax' matches what user paid.
  // Logic: User previously asked for (price * quantity).
  // In Woo: line_item.price is usually per unit.
  
  let itemsRevenue = 0;
  let productIds = [];
  
  if (order.line_items && Array.isArray(order.line_items)) {
    order.line_items.forEach(item => {
      // item.total is the line total (quantity * price), usually excluding tax
      // To be safe and consistent with logic "price * quantity":
      const price = parseFloat(item.price) || 0;
      const quantity = parseInt(item.quantity) || 0;
      
      // We'll use the calculated one to match previous logic
      itemsRevenue += (price * quantity);
      
      if (item.product_id) {
        productIds.push(item.product_id);
      }
    });
  }
  
  // 6. ID products (Joined string)
  const productsIdStr = productIds.join(", ");
  
  // 7. Status
  const status = order.status;

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

function fetchWooOrdersForExport_(startDate, endDate) {
  const allOrders = [];
  let page = 1;
  let keepFetching = true;
  
  // WooCommerce REST API Basic Auth
  const authHeader = "Basic " + Utilities.base64Encode(WOO_EXPORT_CONFIG.consumerKey + ":" + WOO_EXPORT_CONFIG.consumerSecret);
  
  const options = {
    method: 'get',
    muteHttpExceptions: true,
    headers: { 'Authorization': authHeader }
  };

  while (keepFetching) {
    // API v3 is standard for recent Woo versions
    // Order 'desc' is default
    const url = `${WOO_EXPORT_CONFIG.url}/wp-json/wc/v3/orders?per_page=50&page=${page}&after=${startDate.toISOString()}&before=${endDate.toISOString()}`;
    
    Logger.log(`[WooExport] Fetching page ${page}...`);
    
    // Use Shared Utility
    const response = fetchUrlWithRetry(url, options, true); // Get full response for headers if needed (Woo puts paging in header x-wp-totalpages)
    
    if (!response || !response.content) {
      Logger.log("[WooExport] Null response. Stopping.");
      break;
    }
    
    const ordersBatch = response.content;
    
    if (!Array.isArray(ordersBatch) || ordersBatch.length === 0) {
      break;
    }
    
    allOrders.push(...ordersBatch);
    
    // Check pagination
    // WooCommerce sends 'X-WP-TotalPages' header
    if (response.headers) {
       const totalPages = parseInt(response.headers['X-WP-TotalPages'] || response.headers['x-wp-totalpages']);
       if (!isNaN(totalPages) && page >= totalPages) {
         keepFetching = false;
       } else {
         page++;
       }
    } else {
       // If no header, just try next page until empty
       if (ordersBatch.length < 50) {
         keepFetching = false;
       } else {
         page++;
       }
    }
    
    // Safety break
    if (page > 100) keepFetching = false;
  }
  
  return allOrders;
}
