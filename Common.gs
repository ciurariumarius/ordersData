/**
 * @file Config.gs
 * @description Helper script to read configuration values from the "Config" sheet.
 */

// ==========================================
// CONFIGURATION HELPER
// ==========================================

const CONFIG_SHEET_NAME = "Config";

/**
 * Reads configuration key-value pairs from the "Config" sheet.
 * Expected structure: Column A = Key, Column B = Value.
 * @returns {Object} object containing config values
 */
function fetchConfigFromSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);

  if (!sheet) {
    throw new Error(`CRITICAL: Sheet "${CONFIG_SHEET_NAME}" not found. Please create it and add credentials.`);
  }

  // Read all data from the sheet (assuming keys in col A, values in col B)
  // Using getDataRange to avoid reading empty rows if possible, or just read A:B
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    throw new Error(`CRITICAL: Sheet "${CONFIG_SHEET_NAME}" is empty.`);
  }

  const range = sheet.getRange(1, 1, lastRow, 2); // Read A1:B[lastRow]
  const data = range.getValues();

  const config = {};

  data.forEach(row => {
    const key = row[0];
    const value = row[1];
    if (key && typeof key === 'string') {
        // Trim whitespace from keys/values to be safe
        config[key.trim()] = value ? value.toString().trim() : "";
    }
  });

  // Basic validation
  // Note: We check for keys, but scripts should handle missing specific keys gracefully if only one platform is used.
  // Actually, forcing all might annoy users who only use one. 
  // Better approach: Log warnings for missing keys but don't crash Config unless empty.
  // Updating list to include Woo keys as "known keys" we might want to warn about if some but not all are present?
  // For now, let's keep it simple: log warning if standard keys are missing.
  const requiredKeys = ['shopifyDomain', 'shopifyAPIkey', 'gomagDomain', 'gomagAPIkey', 'wooUrl', 'wooConsumerKey', 'wooConsumerSecret'];
  const missingKeys = requiredKeys.filter(k => !config[k]);

  if (missingKeys.length > 0) {
    Logger.log(`[Config] WARNING: Missing keys in config sheet: ${missingKeys.join(', ')}`);
  }

  return config;
}
// ==========================================
// SHARED UTILITIES
// ==========================================

/**
 * Calculates the start and end dates based on a lookback window of N days,
 * excluding the current day (up to Today 00:00:00).
 * 
 * Example:
 * Today = Feb 06. Days = 1.
 * endDate = Feb 06 00:00.
 * startDate = Feb 05 00:00.
 * Range: [Feb 05, Feb 06) -> covers strictly Feb 05.
 * 
 * @param {number} days - Number of days to look back.
 * @returns {Object} { startDate: Date, endDate: Date }
 */
function calculateDateRange(days) {
  const endDate = new Date();
  endDate.setHours(0, 0, 0, 0); // Midnight today starts the exclusion zone
  
  const startDate = new Date(endDate);
  startDate.setDate(endDate.getDate() - days); // Go back N days from midnight
  
  return { startDate, endDate };
}

/**
 * Fetches a URL with retry logic for 429 errors and transient failures.
 * @param {string} url - The URL to fetch.
 * @param {Object} options - UrlFetchApp options.
 * @param {boolean} returnFullResponse - If true, returns {content: ..., headers: ...}. Default false.
 * @returns {Object|null} Parsed JSON response (or wrapper) or null on failure/error.
 */
function fetchUrlWithRetry(url, options, returnFullResponse = false) {
  for (let i = 0; i < 3; i++) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      const contentText = response.getContentText();
      
      if (code >= 200 && code < 300) {
        const jsonContent = JSON.parse(contentText);
        if (returnFullResponse) {
          return {
            content: jsonContent,
            headers: response.getHeaders()
          };
        }
        return jsonContent;
      } else if (code === 429) {
        const retryAfter = response.getHeaders()['Retry-After'] || 5;
        Logger.log(`[API] Rate limit hit. Sleeping for ${retryAfter}s...`);
        Utilities.sleep(parseInt(retryAfter) * 1000);
        continue;
      } else {
        Logger.log(`[API ERROR] Code: ${code}. Response: ${contentText}`);
        return null; 
      }
    } catch (e) {
      if (i === 2) {
        Logger.log(`[FETCH ERROR] ${e.message}`);
        return null;
      }
      Utilities.sleep(2000);
    }
  }
  return null;
}

// ==========================================
// SHEET HELPERS
// ==========================================

/**
 * Writes data to a specified sheet, creating it if necessary and clearing old data.
 * @param {Spreadsheet} spreadsheet - The Google Spreadsheet object.
 * @param {string} sheetName - The name of the sheet to write to.
 * @param {Array<string>} headers - Array of header names.
 * @param {Array<Array>} rows - 2D array of data rows.
 */
function writeParamsToSheet(spreadsheet, sheetName, headers, rows) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  } else {
    sheet.clear(); 
  }

  // Write Headers
  // Light green background #d9ead3 is typical for these reports
  sheet.getRange(1, 1, 1, headers.length)
       .setValues([headers])
       .setFontWeight("bold")
       .setBackground("#d9ead3");

  if (rows && rows.length > 0) {
    // Write Data
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    
    // Auto-detect numeric columns for simple formatting if needed, 
    // but typically we can format specific known columns from the caller or just leave as auto.
    // Here we'll apply a standard number format to potential numeric columns (Value, Shipping, Revenue)
    // assuming they are generally in columns 3, 4, 5 based on current usage.
    // To be safer, we could just format the whole data range as needed.
    
    // Auto-resize
    sheet.autoResizeColumns(1, headers.length);
  } else {
    Logger.log(`[Sheet] No data rows to write for ${sheetName}.`);
  }
}
