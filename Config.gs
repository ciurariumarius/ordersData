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
  const requiredKeys = ['shopifyDomain', 'shopifyAPIkey', 'gomagDomain', 'gomagAPIkey'];
  const missingKeys = requiredKeys.filter(k => !config[k]);

  if (missingKeys.length > 0) {
    Logger.log(`[Config] WARNING: Missing keys in config sheet: ${missingKeys.join(', ')}`);
  }

  return config;
}
