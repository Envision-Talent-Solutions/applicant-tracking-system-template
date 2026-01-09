/** @file Settings.gs - Manages reading configuration from the SETTINGS sheet. */

/**
 * Reads all dropdown list configurations from the SETTINGS sheet.
 * @returns {Map<string, string[]>} A map where the key is the header name and the value is an array of dropdown options.
 */
function getDropdownConfigurations() {
  const ss = SpreadsheetApp.getActive();
  const settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!settingsSheet) {
    logWarn('SETTINGS sheet not found. Cannot build dynamic validations.');
    return new Map();
  }

  const range = settingsSheet.getDataRange();
  const values = range.getValues();
  if (values.length < 1) return new Map();

  const headers = values[0];
  const configMap = new Map();

  for (let c = 0; c < headers.length; c++) {
    const headerName = headers[c];
    if (!headerName) continue;

    const options = [];
    for (let r = 1; r < values.length; r++) {
      const option = values[r][c];
      if (option && String(option).trim() !== '') {
        options.push(option);
      }
    }
    if (options.length > 0) {
      configMap.set(headerName, options);
    }
  }
  return configMap;
}