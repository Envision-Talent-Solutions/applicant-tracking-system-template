/** @file Triggers.gs - Menu setup, trigger installation, and edit handlers. */

/**
 * Required for Marketplace add-on installation.
 * Called when a user installs the add-on from the Marketplace.
 * @param {Object} e - The install event object
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Required for granular OAuth consent handling.
 * Called when a user grants file-level scope access.
 * @param {Object} e - The event object containing granted scopes
 */
function onFileScopeGranted(e) {
  // Re-run onOpen to refresh menus with newly available features
  onOpen(e);

  // Log the grant for debugging
  try {
    const grantedScopes = e && e.authMode ?
      ScriptApp.getAuthorizationInfo(e.authMode).getAuthorizedScopes() : [];
    logInfo('File scope granted', { scopes: grantedScopes });
  } catch (err) {
    // Non-critical - continue silently
  }
}

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();

  // Check critical authorization on startup (non-blocking, just warns)
  // onOpen runs in LIMITED mode - we can check but not prompt
  try {
    const authMode = e && e.authMode ? e.authMode : ScriptApp.AuthMode.LIMITED;
    checkCriticalAuthorization(authMode);
  } catch (err) {
    // If auth check itself fails, continue with menu setup
    // User will see errors when they try to use features
  }

  // Single consolidated menu for all system functions
  ui.createMenu('⚙️ System Tools')
    .addSubMenu(ui.createMenu('👥 Candidate Management')
      .addItem('📥 Import Candidates', 'showImportSidebar')
      .addItem('🔄 Sync All Data', 'Full_Resync')
      .addItem('📝 Setup Candidate Form', 'InitOrRepair_Form'))

    .addSubMenu(ui.createMenu('📄 Requisition Management')
      .addItem('📅 Update Days Open', 'Recompute_DaysOpen_All'))

    .addSubMenu(ui.createMenu('⚡ System Admin Functions')
      .addItem('🔧 Install/Repair Triggers', 'Install_Triggers')
      .addItem('✅ Rebuild Data Validations', 'Rebuild_Validations')
      .addItem('🧹 Clean Up Links', 'linkHygieneSweep_')
      .addItem('📋 Check Authorization Status', 'showAuthorizationStatus')
      .addItem('🔑 Authorize Script', 'promptForAuthorizationIfNeeded'))

    .addToUi();
}

function Install_Triggers() {
  const ui = SpreadsheetApp.getUi();

  // Use hard enforcement for triggers - this will prompt user if not authorized
  // Per Google's guidance: "Call requireScopes() before installing triggers"
  try {
    requireFeatureScopes('TRIGGERS');
  } catch (e) {
    // User cancelled the authorization prompt
    ui.alert(
      'Authorization Cancelled',
      'Trigger installation was cancelled because authorization was not granted.',
      ui.ButtonSet.OK
    );
    return;
  }

  // Re-check after requireScopes - with granular consent, user may have denied
  const authStatus = checkFeatureAuthorization('TRIGGERS');
  if (!authStatus.authorized) {
    ui.alert(
      'Authorization Incomplete',
      'The script trigger permission was not granted.\n\n' +
      'Triggers cannot be installed without this permission.',
      ui.ButtonSet.OK
    );
    return;
  }

  const ss = SpreadsheetApp.getActive();
  const id = ss.getId();

  const all = ScriptApp.getProjectTriggers();
  for (const t of all) ScriptApp.deleteTrigger(t);

  ScriptApp.newTrigger('onEdit_Installable').forSpreadsheet(id).onEdit().create();
  ScriptApp.newTrigger('onChange_Installable').forSpreadsheet(id).onChange().create();
  ScriptApp.newTrigger('processFormSubmission').forSpreadsheet(id).onFormSubmit().create();
  ScriptApp.newTrigger('refreshJobIdChoicesInForm').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('Recompute_DaysOpen_All').timeBased().atHour(3).nearMinute(10).everyDays(1).create();

  // Rebuild data validations to ensure dropdown menus are populated
  try {
    rebuildAllValidations_();
    logInfo('Data validations rebuilt during trigger installation');
  } catch (e) {
    logWarn('Failed to rebuild validations during install', { error: e.message });
  }

  logInfo('All triggers installed successfully.', {});
  SpreadsheetApp.getUi().alert(
    'Triggers Installed',
    'The following automated triggers have been set up:\n\n' +
    '• Edit Trigger - Syncs data when you make changes\n' +
    '• Change Trigger - Handles structural changes\n' +
    '• Form Submit - Processes new candidate submissions\n' +
    '• Hourly - Updates Job ID choices in the form\n' +
    '• Daily (3:10 AM) - Recalculates "Days Open" metrics\n\n' +
    'Dropdown menus have been configured with your settings.\n\n' +
    'Your ATS will now run automatically!',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function _collectJobIdsFromAll_(sh, hm, rows) {
  if (!hm || !hm[H_ALL.JobID]) return [];
  const ids = new Set();
  for (const r of rows) {
    try {
      const obj = getRowObjectByHeaders(sh, hm, r);
      const jid = (obj[H_ALL.JobID] || '').toString().trim();
      if (jid) ids.add(jid);
    } catch (e) {
      logWarn('Error collecting Job ID from Candidate Database', { row: r, error: e.message });
    }
  }
  return Array.from(ids);
}

function _collectJobIdsFromActive_(sh, hm, rows) {
  if (!hm || !hm[H_ACT.JobID]) return [];
  const ids = new Set();
  for (const r of rows) {
    try {
      const obj = getRowObjectByHeaders(sh, hm, r);
      const jid = (obj[H_ACT.JobID] || '').toString().trim();
      if (jid) ids.add(jid);
    } catch (e) {
      logWarn('Error collecting Job ID from Active Candidates', { row: r, error: e.message });
    }
  }
  return Array.from(ids);
}

function _collectJobIdsFromReqs_(sh, hm, rows) {
  if (!hm || !hm[H_REQ.JobID]) return [];
  const ids = new Set();
  for (const r of rows) {
    try {
      const obj = getRowObjectByHeaders(sh, hm, r);
      const jid = (obj[H_REQ.JobID] || '').toString().trim();
      if (jid) ids.add(jid);
    } catch (e) {
      logWarn('Error collecting Job ID from Requisitions', { row: r, error: e.message });
    }
  }
  return Array.from(ids);
}

function onEdit_Installable(e) {
  const range = e && e.range; 
  if (!range) return;
  
  const sh = range.getSheet();
  const name = sh.getName();

  const trackedSheets = [SHEET_ALL, SHEET_ACTIVE, SHEET_REQUISITIONS];
  if (!trackedSheets.includes(name)) {
    return;
  }

  const headerGuardPassed = handleHeaderGuard(e);
  if (!headerGuardPassed) return;

  const handler = {
    [SHEET_ALL]:          handleEditAllCandidates,
    [SHEET_ACTIVE]:       handleEditActiveCandidates,
    [SHEET_REQUISITIONS]: handleEditRequisitions,
  }[name];

  if (handler) {
    try {
      handler(e);
    } catch (error) {
      logWarn('Error in onEdit handler', { sheet: name, error: error.message, stack: error.stack });
    }
  }
}

function onChange_Installable(e) {
  const t = e && e.changeType; 
  if (!t) return;
  
  const structuralChanges = ['INSERT_ROW', 'REMOVE_ROW', 'INSERT_GRID', 'REMOVE_GRID'];
  if (!structuralChanges.includes(t)) {
    return;
  }
  
  withLock(() => {
    try {
      reconcileActiveMembership_All_();
    } catch (error) {
      logWarn('Error in onChange handler', { error: error.message });
    }
  }, 2000, () => {
    enqueueAndSchedule_Reconcile(['all']);
  });
}

function handleHeaderGuard(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  
  const anchors = {
    [SHEET_REQUISITIONS]: ANCHOR_HEADER_REQ,
    [SHEET_ALL]: ANCHOR_HEADER_ALL,
    [SHEET_ACTIVE]: ANCHOR_HEADER_ACT,
  };
  
  const anchor = anchors[sheetName];
  if (!anchor) return true;

  const headerInfo = getHeaderInfo(sheet, anchor);
  if (!headerInfo || range.getRow() !== headerInfo.headerRow) {
    return true;
  }

  const oldValue = e.oldValue;
  const newValue = e.value;
  
  if (String(oldValue).trim() === anchor && (!newValue || String(newValue).trim() !== anchor)) {
    SpreadsheetApp.getActive().toast(
      `The "${anchor}" header is critical and cannot be changed. Reverting edit.`, 
      'System Protection', 
      5
    );
    range.setValue(oldValue);
    return false;
  }
  
  invalidateHeaderCache(sheet);
  return true;
}

function handleEditAllCandidates(e) {
  withLock(() => {
    const range = e.range;
    const sh = range.getSheet();
    const headerInfo = getHeaderInfo(sh, ANCHOR_HEADER_ALL);
    
    if (!headerInfo) {
      logWarn('Could not get header info for Candidate Database sheet');
      return;
    }
    
    const { headerMap: hm, dataStartRow } = headerInfo;
    
    const rows = Array.from({ length: range.getNumRows() }, (_, i) => range.getRow() + i);
    const dataRows = rows.filter(r => r >= dataStartRow);
    
    if (dataRows.length === 0) return;
    
    if (hm[H_ALL.Email] && rangesIntersectColumns_(range, hm[H_ALL.Email])) {
      try {
        enforceUniqueEmail_(sh, dataRows);
      } catch (error) {
        logWarn('Error enforcing unique email', { error: error.message });
      }
    }
    
    try {
      stampCreatedAndUpdated_All_(sh, dataRows);
    } catch (error) {
      logWarn('Error stamping timestamps', { error: error.message });
    }
    
    if (hm[H_ALL.JobID] && rangesIntersectColumns_(range, hm[H_ALL.JobID])) {
      try {
        autopopulateAllFromJobId_(dataRows);
      } catch (error) {
        logWarn('Error auto-populating from Job ID', { error: error.message });
      }
    }
    
    const affectedJobIds = _collectJobIdsFromAll_(sh, hm, dataRows);
    if (affectedJobIds.length) {
      enqueueAndSchedule_Reconcile(affectedJobIds);
    }
  });
}

function handleEditActiveCandidates(e) {
  withLock(() => {
    const range = e.range;
    const sh = range.getSheet();
    const headerInfo = getHeaderInfo(sh, ANCHOR_HEADER_ACT);
    
    if (!headerInfo) {
      logWarn('Could not get header info for Active Candidates sheet');
      return;
    }
    
    const { headerMap: hmAct, dataStartRow } = headerInfo;

    const rows = Array.from({ length: range.getNumRows() }, (_, i) => range.getRow() + i);
    const dataRows = rows.filter(r => r >= dataStartRow);
    
    if (dataRows.length === 0) return;
    
    const affected = new Set();
    
    for (const row of dataRows) {
      try {
        const obj = getRowObjectByHeaders(sh, hmAct, row);
        const jobId = (obj[H_ACT.JobID] || '').toString().trim();
        const email = normEmail(obj[H_ACT.Email]);
        
        if (!jobId || !email) continue;
        
        const key = keyFor(jobId, email);
        markRecentEdit(SHEET_ACTIVE, key);
        logInfo('Marked Active row as recently edited', { key: key, row: row });
        
        upsertAllFromActive_(obj);
        affected.add(jobId);
      } catch (error) {
        logWarn('Error processing Active Candidates edit', { row, error: error.message });
      }
    }
    
    if (affected.size) {
      enqueueAndSchedule_Reconcile(Array.from(affected));
    }
  });
}

function handleEditRequisitions(e) {
  withLock(() => {
    const range = e.range;
    const sh = range.getSheet();
    const headerInfo = getHeaderInfo(sh, ANCHOR_HEADER_REQ);
    
    if (!headerInfo) {
      logWarn('Could not get header info for Requisitions sheet');
      return;
    }
    
    const { headerMap: hm, dataStartRow } = headerInfo;

    const rows = Array.from({ length: range.getNumRows() }, (_, i) => range.getRow() + i);
    const meaningfulRows = rows.filter(r => r >= dataStartRow);
    
    if (meaningfulRows.length === 0) return;
    
    try {
      ensureJobIds_();
    } catch (error) {
      logWarn('Error ensuring Job IDs', { error: error.message });
    }
    
    try {
      applyReqStatusTransitionsForRows_(meaningfulRows);
    } catch (error) {
      logWarn('Error applying status transitions', { error: error.message });
    }
    
    if (hm[H_REQ.JobID] && rangesIntersectColumns_(range, hm[H_REQ.JobID])) {
      try {
        syncJobIdDropdowns_();
      } catch (error) {
        logWarn('Error syncing Job ID dropdowns', { error: error.message });
      }
    }
    if (hm[H_REQ.JobStatus] && rangesIntersectColumns_(range, hm[H_REQ.JobStatus])) {
      try {
        syncJobIdDropdowns_();
      } catch (error) {
        logWarn('Error syncing Job ID dropdowns', { error: error.message });
      }
    }

    const jobIds = _collectJobIdsFromReqs_(sh, hm, meaningfulRows);
    if (jobIds.length > 0) {
      enqueueAndSchedule_Reconcile(jobIds);
    }
  });
}