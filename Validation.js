/**
 * @file Validation.gs - Data validation management for dropdown menus.
 */

/**
 * Rebuilds all validations only if SETTINGS changed.
 * Uses hash check to skip unnecessary rebuilds.
 */
function rebuildAllValidations_() {
  const ss = SpreadsheetApp.getActive();
  const configurations = getDropdownConfigurations();

  if (configurations.size === 0) {
    logInfo('No configurations found in SETTINGS sheet.');
  }

  // Check if settings changed using hash comparison
  const newHash = JSON.stringify(
    Array.from(configurations.entries())
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([key, values]) => ({ key, values: values.sort() }))
  );
  const oldHash = StateManager.getSettingsHash();
  
  if (newHash === oldHash) {
    logInfo('Settings unchanged, skipping validation rebuild', { 
      configCount: configurations.size 
    });
  } else {
    logInfo('Settings changed, rebuilding validations', { 
      configCount: configurations.size 
    });

    const sheetsToValidate = [
      { sheetName: SHEET_REQUISITIONS, anchor: ANCHOR_HEADER_REQ },
      { sheetName: SHEET_ALL,          anchor: ANCHOR_HEADER_ALL },
      { sheetName: SHEET_ACTIVE,       anchor: ANCHOR_HEADER_ACT }
    ];

    let validationsApplied = 0;
    for (const { sheetName, anchor } of sheetsToValidate) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) continue;

      const headerInfo = getHeaderInfo(sheet, anchor);
      if (!headerInfo) continue;
      
      for (const header of Object.keys(headerInfo.headerMap)) {
        if (configurations.has(header)) {
          const optionsFromSettings = configurations.get(header);
          applyListValidation(sheet, headerInfo, header, optionsFromSettings);
          validationsApplied++;
        }
      }
    }
    
    // Store new hash
    StateManager.setSettingsHash(newHash);
    logInfo('Validation rebuild complete', { validationsApplied });
  }

  // Always rebuild Job ID dropdowns (dynamic from Requisitions)
  try {
    syncJobIdDropdowns_();
  } catch (e) {
    logWarn('Failed to sync Job ID dropdowns', { error: e.message });
  }
}

/**
 * Syncs Job ID dropdowns across sheets.
 * Only shows Job IDs for Open or On Hold positions.
 */
function syncJobIdDropdowns_() {
  const ss = SpreadsheetApp.getActive();
  const allSheet = ss.getSheetByName(SHEET_ALL);
  const actSheet = ss.getSheetByName(SHEET_ACTIVE);

  const jobIds = _getActiveJobIds_();

  // Handle empty requisitions case - clear validations completely
  if (jobIds.length === 0) {
    logInfo('No active Job IDs found in Requisitions, clearing dropdown validations');
    
    // Clear Candidate Database Job ID validation
    if (allSheet) {
      const allHeaderInfo = getHeaderInfo(allSheet, ANCHOR_HEADER_ALL);
      if (allHeaderInfo && allHeaderInfo.headerMap[H_ALL.JobID]) {
        _clearColumnValidation_(allSheet, allHeaderInfo, H_ALL.JobID);
      }
    }
    
    // Clear Active Candidates Job ID validation
    if (actSheet) {
      const actHeaderInfo = getHeaderInfo(actSheet, ANCHOR_HEADER_ACT);
      if (actHeaderInfo && actHeaderInfo.headerMap[H_ACT.JobID]) {
        _clearColumnValidation_(actSheet, actHeaderInfo, H_ACT.JobID);
      }
    }
    
    return;
  }
  
  // Apply to Candidate Database sheet
  if (allSheet) {
    const allHeaderInfo = getHeaderInfo(allSheet, ANCHOR_HEADER_ALL);
    if (allHeaderInfo && allHeaderInfo.headerMap[H_ALL.JobID]) {
      applyListValidation(allSheet, allHeaderInfo, H_ALL.JobID, jobIds);
      logInfo('Job ID dropdown synced to Candidate Database (active positions only)', { count: jobIds.length });
    }
  }
  
  // Apply to Active Candidates sheet
  if (actSheet) {
    const actHeaderInfo = getHeaderInfo(actSheet, ANCHOR_HEADER_ACT);
    if (actHeaderInfo && actHeaderInfo.headerMap[H_ACT.JobID]) {
      applyListValidation(actSheet, actHeaderInfo, H_ACT.JobID, jobIds);
      logInfo('Job ID dropdown synced to Active Candidates (active positions only)', { count: jobIds.length });
    }
  }
}

/**
 * Retrieves Job IDs for positions that are Open or On Hold.
 * @returns {string[]} Array of Job IDs for active positions
 */
function _getActiveJobIds_() {
  const ss = SpreadsheetApp.getActive();
  const reqSheet = ss.getSheetByName(SHEET_REQUISITIONS);
  
  if (!reqSheet) {
    logInfo('No Requisitions sheet found - returning empty Job ID list');
    return [];
  }
  
  const reqHeaderInfo = getHeaderInfo(reqSheet, ANCHOR_HEADER_REQ);
  if (!reqHeaderInfo) {
    logInfo('No header info found in Requisitions - returning empty Job ID list');
    return [];
  }
  
  const { rows } = buildReqIndex(reqSheet, reqHeaderInfo.headerMap, reqHeaderInfo.dataStartRow);
  
  // Filter to only Open or On Hold positions (exclude Closed, Pending Approval, Hired)
  const activeJobIds = rows
    .filter(r => {
      const status = (r.data[H_REQ.JobStatus] || '').toString().trim();
      return OPEN_STATUSES.has(status);
    })
    .map(r => r.data[H_REQ.JobID])
    .filter(id => id && String(id).trim() !== '')
    .map(String);
  
  return activeJobIds;
}

/**
 * Clears validation rules from a column.
 * @param {Sheet} sheet - The sheet to clear validations from
 * @param {Object} headerInfo - Header information object
 * @param {string} header - The header name to clear validation for
 */
function _clearColumnValidation_(sheet, headerInfo, header) {
  const col = headerInfo.headerMap[header];
  if (!col) return;

  const startRow = headerInfo.dataStartRow;
  if (startRow > sheet.getMaxRows()) return;
  
  try {
    const maxRows = sheet.getMaxRows();
    const numRowsToValidate = maxRows - startRow + 1;
    const rng = sheet.getRange(startRow, col, numRowsToValidate);
    
    // Clear any existing validation
    rng.clearDataValidations();
    
    logInfo('Cleared validation for column', { 
      sheet: sheet.getName(), 
      header 
    });
  } catch (e) {
    logWarn('Failed to clear column validation', { 
      sheet: sheet.getName(), 
      header, 
      error: e.message 
    });
  }
}

/**
 * Applies a list validation rule only if the values have changed
 * Preserves user-applied color formatting inside validation rules
 * @param {Sheet} sheet - The sheet to apply validation to
 * @param {Object} headerInfo - Header information object
 * @param {string} header - The header name to apply validation to
 * @param {string[]} newList - Array of valid values
 */
function applyListValidation(sheet, headerInfo, header, newList) {
  if (!newList || !Array.isArray(newList) || newList.length === 0) {
    logWarn('Attempted to apply empty validation list', { 
      sheet: sheet.getName(), 
      header 
    });
    return;
  }
  
  const col = headerInfo.headerMap[header];
  if (!col) return;

  const startRow = headerInfo.dataStartRow;
  if (startRow > sheet.getMaxRows()) return;
  
  const firstCell = sheet.getRange(startRow, col);
  const existingRule = firstCell.getDataValidation();
  
  const existingList = _extractAllowedValuesFromRule_(existingRule);

  if (_areListsEqual_(existingList, newList)) {
    logInfo('Validation unchanged - preserved formatting', { 
      sheet: sheet.getName(), 
      header 
    });
    return;
  }

  // Check if we're adding new values to an existing list (preserve colors)
  if (existingList && existingList.length > 0) {
    const onlyNewValues = newList.filter(v => !existingList.includes(v));
    const removedValues = existingList.filter(v => !newList.includes(v));
    
    if (onlyNewValues.length > 0 && removedValues.length === 0) {
      logInfo('Adding new values to existing validation - preserving colors', {
        sheet: sheet.getName(),
        header,
        newValues: onlyNewValues
      });
      
      const maxRows = sheet.getMaxRows();
      const numRowsToValidate = maxRows - startRow + 1;
      const rng = sheet.getRange(startRow, col, numRowsToValidate);
      
      const newRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(newList, true)
        .setAllowInvalid(true)
        .build();

      safeSetDataValidation(rng, newRule, { header });
      
      SpreadsheetApp.getActive().toast(
        `⚠️ Added new values to ${header} dropdown. Please manually re-add colors if needed.`,
        'Validation Updated',
        8
      );
      return;
    }
  }

  const maxRows = sheet.getMaxRows();
  const numRowsToValidate = maxRows - startRow + 1;
  const rng = sheet.getRange(startRow, col, numRowsToValidate);
  
  const newRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(newList, true)
    .setAllowInvalid(true)
    .build();

  safeSetDataValidation(rng, newRule, { header });
  logInfo('Validation rule updated', { 
    sheet: sheet.getName(), 
    header,
    valueCount: newList.length 
  });
}

/**
 * Public function wrapper with lock protection
 */
function Rebuild_Validations() {
  withLock(() => rebuildAllValidations_());
}

/**
 * Extracts allowed values from an existing validation rule.
 * @param {DataValidation} rule - The validation rule to extract from
 * @returns {string[]|null} Array of allowed values or null
 */
function _extractAllowedValuesFromRule_(rule) {
  try {
    if (!rule) return null;
    const criteria = rule.getCriteriaType();
    if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      const values = rule.getCriteriaValues();
      return values[0].map(String);
    }
  } catch (e) {
     logWarn('Could not extract values from existing validation rule', { 
       error: e.message 
     });
  }
  return null;
}

/**
 * Compares two lists for equality (order-independent)
 * @param {string[]} listA - First list
 * @param {string[]} listB - Second list
 * @returns {boolean} True if lists contain same values
 */
function _areListsEqual_(listA, listB) {
  if (!listA || !listB) return false;
  if (listA.length !== listB.length) return false;
  
  const sortedA = [...listA].sort();
  const sortedB = [...listB].sort();

  for (let i = 0; i < sortedA.length; i++) {
    if (sortedA[i] !== sortedB[i]) return false;
  }
  return true;
}