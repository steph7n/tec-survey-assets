/**
 * Utils.gs
 *
 * Shared utility functions used across the Survey Core system.
 * Currently includes:
 * - getConfigAsObject: read a key–value Config sheet into a JS object
 * - getCurrentYearConfigFileId: return the current Yearly Config file ID
 */

/**
 * Reads a key–value Config sheet (column A = key, column B = value)
 * into a plain JS object { key: value, ... }.
 */

function getConfigAsObject(sheet) {
  const range = sheet.getDataRange().getValues();
  const obj = {};
  for (let i = 0; i < range.length; i++) {
    const key = range[i][0];
    const val = range[i][1];
    if (key) obj[key] = val;
  }
  return obj;
}

/**
 * Reads the current Yearly Config File ID from the
 * Survey Core Database → Config sheet (cell B1).
 *
 * @return {string} The file ID, or an empty string if missing.
 */
function getCurrentConfigFileId() {
  try {
    const coreSS = SpreadsheetApp.getActive();
    const cfgSheet = coreSS.getSheetByName("Config");
    if (!cfgSheet) return "";

    const val = cfgSheet.getRange("B1").getValue();
    return val ? String(val).trim() : "";
  } catch (err) {
    return "";
  }
}