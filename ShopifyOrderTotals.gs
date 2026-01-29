/**
 * @file ShopifyOrderSum.gs
 * @description Script Shopify redenumit pentru a evita conflictele cu Gomag.
 */

// ==========================================
// 1. CONFIGURATION (SHOPIFY SPECIFIC)
// ==========================================

// Am redenumit CONFIG in SHOPIFY_CONFIG
// ==========================================
// 1. CONFIGURATION (SHOPIFY SPECIFIC)
// ==========================================

// Initialise defaults - values will be overwritten by fetchConfigFromSheet()
let SHOPIFY_CONFIG = {
  // Your myshopify domain
  domain: '', 
  
  // Your Admin API Access Token
  accessToken: '', 
  
  // How many days back to look?
  timeframeDays: 30, 
  
  // API Version
  apiVersion: '2024-04' 
};

// Am redenumit variabila pentru numele foii
const SHOPIFY_SHEET_NAME = "ShopifyOrdersTotal"; 

// ==========================================
// 2. MAIN FUNCTION
// ==========================================

function runOrderTotalCalculator() {
  // 0. Load Configuration
  try {
    const config = fetchConfigFromSheet();
    if (config.shopifyDomain) SHOPIFY_CONFIG.domain = config.shopifyDomain;
    if (config.shopifyAPIkey) SHOPIFY_CONFIG.accessToken = config.shopifyAPIkey;
    
    // Validare
    if (!SHOPIFY_CONFIG.domain || !SHOPIFY_CONFIG.accessToken) {
       throw new Error("Missing Shopify credentials in Config sheet (shopifyDomain or shopifyAPIkey).");
    }
  } catch (e) {
    Logger.log(`[Shopify] Config Error: ${e.message}`);
    return;
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Calculate Timeframe
  const endDate = new Date();
  const startDate = new Date();
  // Folosim SHOPIFY_CONFIG
  startDate.setDate(endDate.getDate() - SHOPIFY_CONFIG.timeframeDays);
  
  Logger.log(`[Shopify] Starting run. Fetching orders from the last ${SHOPIFY_CONFIG.timeframeDays} days...`);

  // 2. Fetch Data
  try {
    const orders = fetchShopifyOrders_(startDate);
    
    // 3. Process Data
    let totalValue = 0;
    let orderCount = 0;

    orders.forEach(order => {
      if (order.financial_status === 'voided') return;

      const price = parseFloat(order.total_price);
      if (!isNaN(price)) {
        totalValue += price;
        orderCount++;
      }
    });

    Logger.log(`[Shopify] Processed ${orderCount} orders. Total Value: ${totalValue}`);

    // 4. Write to Sheet
    writeShopifyResultToSheet_(spreadsheet, startDate, totalValue, orderCount);

  } catch (e) {
    Logger.log(`[Shopify] ERROR: ${e.message}`);
    // Browser.msgBox eliminat pentru a nu bloca execuția automată, poți decomenta dacă vrei
  }
}

// ==========================================
// 3. HELPER FUNCTIONS
// ==========================================

function fetchShopifyOrders_(startDate) {
  // Folosim SHOPIFY_CONFIG
  const endpoint = `https://${SHOPIFY_CONFIG.domain}/admin/api/${SHOPIFY_CONFIG.apiVersion}/orders.json` +
    `?status=any` + 
    `&created_at_min=${startDate.toISOString()}` +
    `&fields=id,total_price,financial_status,created_at` + 
    `&limit=250`; 

  return fetchShopifyDataWithPagination_(endpoint, SHOPIFY_CONFIG.accessToken);
}

function writeShopifyResultToSheet_(spreadsheet, startDate, totalValue, orderCount) {
  // Folosim SHOPIFY_SHEET_NAME
  let sheet = spreadsheet.getSheetByName(SHOPIFY_SHEET_NAME);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHOPIFY_SHEET_NAME);
  } else {
    sheet.clear(); 
  }

  const headers = ["Calculation Date", "Timeframe Start", "Days Lookback", "Order Count", "Total Value"];
  sheet.getRange("A1:E1").setValues([headers]).setFontWeight("bold").setBackground("#f3f3f3");

  const rowData = [
    new Date(), 
    startDate,  
    SHOPIFY_CONFIG.timeframeDays,
    orderCount,
    totalValue
  ];

  sheet.getRange("A2:E2").setValues([rowData]);
  sheet.getRange("E2").setNumberFormat("#,##0.00"); 
  sheet.autoResizeColumns(1, 5);
}

// ==========================================
// 4. CORE ENGINE
// ==========================================

function fetchShopifyDataWithPagination_(initialUrl, accessToken) {
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
          Logger.log(`Rate limit hit. Sleeping for ${retryAfter}s...`);
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