/** @file CandidatesSync.gs - Two-way sync between Candidate Database and Active Candidates sheets. */

/** Recursion guard for hired flow */
let _processingHiredFlow = false;

// Local helpers
function _normStr_(v){ return (v==null ? '' : String(v)).trim(); }
function _valsEqual_(a,b){
  if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();
  const A = (a===null || typeof a==='undefined') ? '' : a;
  const B = (b===null || typeof b==='undefined') ? '' : b;
  const nA = (typeof A === 'number' || (/^-?\d+(\.\d+)?$/).test(String(A))) ? Number(A) : NaN;
  const nB = (typeof B === 'number' || (/^-?\d+(\.\d+)?$/).test(String(B))) ? Number(B) : NaN;
  if (!isNaN(nA) && !isNaN(nB)) return nA === nB;
  return String(A).trim() === String(B).trim();
}
function _diffFields_(sourceObj, targetObj, fields){
  const out = {};
  for (const f of fields) {
    const src = sourceObj[f] == null ? '' : sourceObj[f];
    const dst = targetObj[f] == null ? '' : targetObj[f];
    if (!_valsEqual_(src, dst)) out[f] = src;
  }
  return out;
}
function _canonReqStatus_(sRaw){
  const s = _normStr_(sRaw).toLowerCase();
  if (!s) return '';
  if (s === 'open') return 'Open';
  if (s === 'on hold' || s === 'on hold') return 'On Hold';
  if (s === 'closed') return 'Closed';
  if (s === 'pending approval') return 'Pending Approval';
  if (s === 'hired') return 'Hired';
  return s.charAt(0).toUpperCase() + s.slice(1);
}
function _isOpenForActive_(status){
  const canon = _canonReqStatus_(status);
  return canon === 'Open' || canon === 'On Hold';
}

/**
 * Applies formatting to a new row with proper empty sheet handling.
 * @param {Sheet} sheet - The sheet to apply formatting to
 * @param {number} templateRow - The row to copy formatting from
 * @param {number} newRow - The new row to format
 * @param {number} dataStartRow - The first valid data row
 */
function _applyFormattingToNewRow_(sheet, templateRow, newRow, dataStartRow) {
  if (newRow < 1) return;

  // Handle case where no valid template exists
  if (templateRow < dataStartRow) {
    logInfo('No valid template row - applying plain formatting', { 
      templateRow, 
      newRow, 
      dataStartRow 
    });
    
    // Clear inherited header formatting and apply plain style
    try {
      const destinationRange = sheet.getRange(newRow, 1, 1, sheet.getMaxColumns());
      destinationRange.setBackground('#ffffff');
      destinationRange.setFontColor('#000000');
      destinationRange.setFontWeight('normal');
      destinationRange.setFontSize(10);
    } catch (e) {
      logWarn('Failed to apply plain formatting', { 
        newRow, 
        error: e.message 
      });
    }
    return;
  }
  
  try {
    const templateRange = sheet.getRange(templateRow, 1, 1, sheet.getMaxColumns());
    const destinationRange = sheet.getRange(newRow, 1, 1, sheet.getMaxColumns());
    
    // Copy only format, not conditional formatting rules
    templateRange.copyTo(destinationRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    
    // Explicitly clear any data that may have been copied
    const currentValues = destinationRange.getValues()[0];
    const emptyValues = currentValues.map(() => '');
    destinationRange.setValues([emptyValues]);
    
  } catch (e) {
    logWarn('Could not apply formatting to new row', { 
      templateRow, 
      newRow, 
      error: e.message 
    });
  }
}

/**
 * Applies validations to an Active Candidates row with proper empty sheet handling.
 * Filters Job IDs to active positions only.
 * @param {Sheet} sheet - The Active Candidates sheet
 * @param {Object} headerInfo - Header information object
 * @param {number} row - The row to apply validations to
 */
function _applyValidationsToActiveRow_(sheet, headerInfo, row) {
  const dataStartRow = headerInfo.dataStartRow;

  // Check if there's a valid template row to copy from
  const templateRow = row - 1;
  
  if (templateRow < dataStartRow) {
    logInfo('No template row for validation copy - will rely on validation rebuild', { 
      row, 
      dataStartRow 
    });
    // Instead of copying, trigger a validation rebuild for this sheet
    try {
      const ss = SpreadsheetApp.getActive();
      const configurations = getDropdownConfigurations();
      const hmAct = headerInfo.headerMap;
      
      // Apply Settings-based validations to this specific row
      for (const header of Object.keys(hmAct)) {
        if (configurations.has(header)) {
          const col = hmAct[header];
          const optionsFromSettings = configurations.get(header);
          
          if (optionsFromSettings && optionsFromSettings.length > 0) {
            const rule = SpreadsheetApp.newDataValidation()
              .requireValueInList(optionsFromSettings, true)
              .setAllowInvalid(true)
              .build();
            
            const cell = sheet.getRange(row, col);
            safeSetDataValidation(cell, rule, { header });
          }
        }
      }
      
      // Apply Job ID validation from Requisitions (active positions only)
      if (hmAct[H_ACT.JobID]) {
        const jobIds = _getActiveJobIds_();
        
        if (jobIds.length > 0) {
          const rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(jobIds, true)
            .setAllowInvalid(true)
            .build();
          
          const cell = sheet.getRange(row, hmAct[H_ACT.JobID]);
          safeSetDataValidation(cell, rule, { header: H_ACT.JobID });
        }
      }
    } catch (e) {
      logWarn('Failed to apply fresh validations to first row', { error: e.message });
    }
    return;
  }
  
  // Copy validations from template row if it exists
  const maxCol = sheet.getMaxColumns();
  
  for (let col = 1; col <= maxCol; col++) {
    try {
      const templateCell = sheet.getRange(templateRow, col);
      const rule = templateCell.getDataValidation();
      
      if (rule) {
        const newCell = sheet.getRange(row, col);
        newCell.setDataValidation(rule);
      }
    } catch (e) {
      logWarn('Failed to copy validation rule', { row, col, error: e.message });
    }
  }
}

function _upsertActiveRow_(allRowObj, linkPayload) {
  const ss = SpreadsheetApp.getActive();
  const all = ss.getSheetByName(SHEET_ALL);
  const act = ss.getSheetByName(SHEET_ACTIVE);
  if (!all || !act) return;

  const allHeaderInfo = getHeaderInfo(all, ANCHOR_HEADER_ALL);
  const actHeaderInfo = getHeaderInfo(act, ANCHOR_HEADER_ACT);
  if (!allHeaderInfo || !actHeaderInfo) return;
  
  const hmAll = allHeaderInfo.headerMap;
  const hmAct = actHeaderInfo.headerMap;

  const jobId = _normStr_(allRowObj[H_ALL.JobID]);
  const email = normEmail(allRowObj[H_ALL.Email]);
  if (!jobId || !email) return;

  const { index: idxAct } = buildCandidateIndex(act, hmAct, actHeaderInfo.dataStartRow);
  const existingRow = idxAct.get(keyFor(jobId, email));

  const desired = {};
  for (const f of MIRRORED_FIELDS) {
    if (f === H_ALL.Resume || f === H_ALL.LinkedIn) continue;
    if (hmAct[f]) desired[f] = allRowObj[f] || '';
  }

  let targetRow = existingRow;
  if (existingRow) {
    const current = getRowObjectByHeaders(act, hmAct, existingRow);
    const diff = _diffFields_(desired, current, Object.keys(desired));
    if (Object.keys(diff).length) setRowValuesByHeaders(act, hmAct, existingRow, diff);
  } else {
    // Capture template formatting before inserting rows (handles empty sheets)
    const template = captureTemplateFormat(act, actHeaderInfo.dataStartRow);

    const lastRow = Math.max(act.getLastRow(), actHeaderInfo.headerRow);
    targetRow = lastRow + 1;
    act.insertRowAfter(lastRow);

    applyTemplateFormat(act, template, targetRow);
    setRowValuesByHeaders(act, hmAct, targetRow, desired);

    try {
      _applyValidationsToActiveRow_(act, actHeaderInfo, targetRow);
    } catch (e) {
      logWarn('Failed to apply validations to new Active row', { error: e.message });
    }
  }

  if (targetRow) {
    const allRow = allRowObj.__row || 0;
    const rUrl = (linkPayload && linkPayload.resumeUrl) || (hmAll[H_ALL.Resume] ? extractUrl(all, allRow, hmAll[H_ALL.Resume]) : '');
    const lUrl = (linkPayload && linkPayload.linkedinUrl) || (hmAll[H_ALL.LinkedIn] ? extractUrl(all, allRow, hmAll[H_ALL.LinkedIn]) : '');

    if (hmAct[H_ACT.Resume] && rUrl) setHyperlink(act, targetRow, hmAct[H_ACT.Resume], rUrl, 'Resume');
    if (hmAct[H_ACT.LinkedIn] && lUrl) setHyperlink(act, targetRow, hmAct[H_ACT.LinkedIn], lUrl, 'LinkedIn Profile');

    // Sync Email as mailto: hyperlink
    const email = normEmail(allRowObj[H_ALL.Email]);
    if (hmAct[H_ACT.Email] && email) {
      setHyperlink(act, targetRow, hmAct[H_ACT.Email], 'mailto:' + email, email);
    }

    // Sync Phone as tel: hyperlink
    const phone = normPhone(allRowObj[H_ALL.Phone]);
    if (hmAct[H_ACT.Phone] && phone && phone.length >= PHONE_MIN_DIGITS) {
      setHyperlink(act, targetRow, hmAct[H_ACT.Phone], 'tel:' + phone, allRowObj[H_ALL.Phone] || phone);
    }
  }
}

/**
 * Autopopulates job title and status from requisitions into Candidate Database
 * @param {number[]} rows - Array of row numbers to process
 */
function autopopulateAllFromJobId_(rows) {
  const ss = SpreadsheetApp.getActive();
  const all = ss.getSheetByName(SHEET_ALL);
  const req = ss.getSheetByName(SHEET_REQUISITIONS);
  if (!all || !req) return;

  const allHeaderInfo = getHeaderInfo(all, ANCHOR_HEADER_ALL);
  const reqHeaderInfo = getHeaderInfo(req, ANCHOR_HEADER_REQ);
  if (!allHeaderInfo || !reqHeaderInfo) return;
  
  const hAll = allHeaderInfo.headerMap;
  const hReq = reqHeaderInfo.headerMap;
  const { idx: reqIdx } = buildReqIndex(req, hReq, reqHeaderInfo.dataStartRow);

  for (const row of rows) {
    if (row < allHeaderInfo.dataStartRow) continue;
    const cur = getRowObjectByHeaders(all, hAll, row);
    const jobId = _normStr_(cur[H_ALL.JobID]);
    if (!jobId) continue;

    const reqRow = reqIdx.get(jobId);
    if (!reqRow) {
      logInfo('Autopopulate skipped – Job ID not found in Requisitions', { row, jobId });
      continue;
    }
    const rObj = getRowObjectByHeaders(req, hReq, reqRow);

    const want = {
      [H_ALL.JobTitle]:  rObj[H_REQ.JobTitle]  || '',
      [H_ALL.JobStatus]: _canonReqStatus_(rObj[H_REQ.JobStatus] || '')
    };
    const diff = _diffFields_(want, cur, [H_ALL.JobTitle, H_ALL.JobStatus]);
    if (Object.keys(diff).length) {
      diff[H_ALL.Updated] = nowDetroit();
      setRowValuesByHeaders(all, hAll, row, diff);
    }
  }
}

/**
 * Enforces unique emails with batched operations for performance.
 * @param {Sheet} sh - The sheet to check
 * @param {number[]} rows - The rows that were just edited
 */
function enforceUniqueEmail_(sh, rows) {
  const headerInfo = getHeaderInfo(sh, ANCHOR_HEADER_ALL);
  if (!headerInfo) return;
  const { headerMap: hm, dataStartRow } = headerInfo;
  if (!hm[H_ALL.Email]) return;

  const last = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (last < dataStartRow) return;

  // Read all data in one batch
  const data = sh.getRange(dataStartRow, 1, last-dataStartRow+1, lastCol).getValues();
  const seen = new Map();
  const rowsToClear = [];

  for (let i = 0; i < data.length; i++) {
    const rowIdx = dataStartRow + i;
    const email = normEmail(data[i][hm[H_ALL.Email]-1]);
    if (!email) continue;
    
    if (!seen.has(email)) {
      seen.set(email, rowIdx);
    } else if (rows.includes(rowIdx)) {
      rowsToClear.push(rowIdx);
      logWarn('Duplicate email blocked', { row: rowIdx, email });
    }
  }

  // Batch clear all duplicates at once
  if (rowsToClear.length > 0) {
    rowsToClear.sort((a, b) => a - b);
    
    let i = 0;
    while (i < rowsToClear.length) {
      const startRow = rowsToClear[i];
      let endRow = startRow;
      
      while (i + 1 < rowsToClear.length && rowsToClear[i + 1] === endRow + 1) {
        i++;
        endRow = rowsToClear[i];
      }
      
      const height = endRow - startRow + 1;
      if (height === 1) {
        sh.getRange(startRow, hm[H_ALL.Email]).setValue('');
      } else {
        const emptyValues = Array(height).fill(['']);
        sh.getRange(startRow, hm[H_ALL.Email], height, 1).setValues(emptyValues);
      }
      
      i++;
    }
    
    toast('Duplicate Email Address(es) blocked; please enter unique emails.');
  }
}

/**
 * Stamps created and updated timestamps, handles hire date logic
 * @param {Sheet} sh - The Candidate Database sheet
 * @param {number[]} rows - Array of row numbers to process
 */
function stampCreatedAndUpdated_All_(sh, rows) {
  const headerInfo = getHeaderInfo(sh, ANCHOR_HEADER_ALL);
  if (!headerInfo) return;
  const { headerMap: hm, dataStartRow } = headerInfo;
  const now = nowDetroit();

  for (const row of rows) {
    if (row < dataStartRow) continue;
    const obj = getRowObjectByHeaders(sh, hm, row);
    const meaningful = Object.keys(obj).some(h => !!obj[h] && String(obj[h]).trim() !== '');
    const updates = {};
    if (meaningful && hm[H_ALL.Created] && isBlank(obj[H_ALL.Created])) updates[H_ALL.Created] = now;
    if (meaningful && hm[H_ALL.Updated]) updates[H_ALL.Updated] = now;

    const stage  = _normStr_(obj[H_ALL.Stage]);
    const status = _normStr_(obj[H_ALL.JobStatus]);
    const isHired = (stage === 'Hired' || status === 'Hired');

    if (isHired) {
      if (hm[H_ALL.HiredDate] && isBlank(obj[H_ALL.HiredDate])) {
        updates[H_ALL.HiredDate] = now;
      }
      const jobId = _normStr_(obj[H_ALL.JobID]);
      const fullName = _normStr_(obj[H_ALL.FullName]);
      const email = normEmail(obj[H_ALL.Email]);
      if (jobId && fullName && email) {
        applyHiredFlow_(jobId, fullName, email);
      }
    } else {
      if (hm[H_ALL.HiredDate] && !isBlank(obj[H_ALL.HiredDate])) {
        updates[H_ALL.HiredDate] = '';
      }
    }

    if (Object.keys(updates).length) setRowValuesByHeaders(sh, hm, row, updates);
  }
}

/**
 * Syncs changes from Active Candidates back to Candidate Database
 * @param {Object} actRowObj - Row object from Active Candidates
 */
function upsertAllFromActive_(actRowObj) {
  const ss = SpreadsheetApp.getActive();
  const all = ss.getSheetByName(SHEET_ALL);
  const act = ss.getSheetByName(SHEET_ACTIVE);
  if (!all || !act) return;

  const allHeaderInfo = getHeaderInfo(all, ANCHOR_HEADER_ALL);
  const actHeaderInfo = getHeaderInfo(act, ANCHOR_HEADER_ACT);
  if (!allHeaderInfo || !actHeaderInfo) return;
  
  const hmAll = allHeaderInfo.headerMap;
  const hmAct = actHeaderInfo.headerMap;

  const jobId = _normStr_(actRowObj[H_ACT.JobID]);
  const email = normEmail(actRowObj[H_ACT.Email]);
  if (!jobId || !email) return;

  const { index: idxAll } = buildCandidateIndex(all, hmAll, allHeaderInfo.dataStartRow);
  const { index: idxAct } = buildCandidateIndex(act, hmAct, actHeaderInfo.dataStartRow);
  const allRow = idxAll.get(keyFor(jobId, email));
  const actRow = idxAct.get(keyFor(jobId, email));
  if (!allRow) return;

  const currentAll = getRowObjectByHeaders(all, hmAll, allRow);
  const desiredAll = {};
  for (const f of MIRRORED_FIELDS) {
    if (f === H_ALL.Resume || f === H_ALL.LinkedIn) continue;
    if (hmAll[f]) desiredAll[f] = actRowObj[f] || '';
  }
  desiredAll[H_ALL.Updated] = nowDetroit();

  const diff = _diffFields_(desiredAll, currentAll, Object.keys(desiredAll));
  if (Object.keys(diff).length) {
    setRowValuesByHeaders(all, hmAll, allRow, diff);
    stampCreatedAndUpdated_All_(all, [allRow]);
  }

  if (actRow) {
    if (hmAll[H_ALL.Resume] && hmAct[H_ACT.Resume]) {
      const url = extractUrl(act, actRow, hmAct[H_ACT.Resume]);
      if (url) setHyperlink(all, allRow, hmAll[H_ALL.Resume], url, 'Resume');
    }
    if (hmAll[H_ALL.LinkedIn] && hmAct[H_ACT.LinkedIn]) {
      const url = extractUrl(act, actRow, hmAct[H_ACT.LinkedIn]);
      if (url) setHyperlink(all, allRow, hmAll[H_ALL.LinkedIn], url, 'LinkedIn Profile');
    }
  }
}

/**
 * Handles the complete hiring workflow with recursion protection.
 * Updates the Requisitions sheet and auto-rejects other candidates for the same job.
 * @param {string} jobId - The Job ID that has been filled
 * @param {string} hiredFullName - The full name of the hired candidate
 * @param {string} hiredEmail - The normalized email of the hired candidate
 */
function applyHiredFlow_(jobId, hiredFullName, hiredEmail) {
  if (_processingHiredFlow) {
    logWarn('Hired flow already in progress, skipping nested call', { jobId });
    return;
  }
  
  _processingHiredFlow = true;
  
  try {
    const ss = SpreadsheetApp.getActive();
    
    // --- Part 1: Update the Requisitions Sheet ---
    const req = ss.getSheetByName(SHEET_REQUISITIONS);
    if (req) {
      const reqHeaderInfo = getHeaderInfo(req, ANCHOR_HEADER_REQ);
      if (reqHeaderInfo) {
        const hmReq = reqHeaderInfo.headerMap;
        const { idx: reqIdx } = buildReqIndex(req, hmReq, reqHeaderInfo.dataStartRow);
        const reqRow = reqIdx.get(jobId);
        if (reqRow) {
          const reqObj = getRowObjectByHeaders(req, hmReq, reqRow);
          const updates = {};

          if (_canonReqStatus_(reqObj[H_REQ.JobStatus]) !== 'Hired') {
            updates[H_REQ.JobStatus] = 'Hired';
          }
          if (hiredFullName && _normStr_(reqObj[H_REQ.HiredCandidateName]) !== _normStr_(hiredFullName)) {
            updates[H_REQ.HiredCandidateName] = hiredFullName;
          }
          
          if (Object.keys(updates).length > 0) {
            setRowValuesByHeaders(req, hmReq, reqRow, updates);
            applyReqStatusTransitionsForRows_([reqRow]);
          }
        }
      }
    }

    // --- Part 2: Auto-reject other candidates ---
    const all = ss.getSheetByName(SHEET_ALL);
    if (!all) return;
    
    const allHeaderInfo = getHeaderInfo(all, ANCHOR_HEADER_ALL);
    if (!allHeaderInfo) return;

    const hmAll = allHeaderInfo.headerMap;
    const { rows: allCandRows } = buildCandidateIndex(all, hmAll, allHeaderInfo.dataStartRow);

    for (const { row, data } of allCandRows) {
      const candJobId = _normStr_(data[H_ALL.JobID]);
      const candEmail = normEmail(data[H_ALL.Email]);

      if (candJobId === jobId && candEmail !== hiredEmail) {
        const currentStage = _normStr_(data[H_ALL.Stage]);
        const updates = {};
        
        if (currentStage !== 'Hired' && currentStage !== 'Rejected') {
          updates[H_ALL.Stage] = 'Rejected';
          updates[H_ALL.RejectedReason] = 'Hired a Different Candidate';
          
          if (Object.keys(updates).length > 0) {
            updates[H_ALL.Updated] = nowDetroit();
            setRowValuesByHeaders(all, hmAll, row, updates);
          }
        }
      }
    }
  } catch (e) {
    logWarn('applyHiredFlow_ failed', { jobId, error: e.message });
  } finally {
    _processingHiredFlow = false;
  }
}

/**
 * Reconciles Active Candidates membership for specific job IDs.
 * Pre-filters rows when possible for better performance.
 * @param {string[]} jobIds - Array of job IDs to reconcile, or empty for all
 */
function reconcileActiveMembership_ByJobIds_(jobIds) {
  const ss = SpreadsheetApp.getActive();
  const all = ss.getSheetByName(SHEET_ALL);
  const act = ss.getSheetByName(SHEET_ACTIVE);
  const req = ss.getSheetByName(SHEET_REQUISITIONS);
  if (!all || !act || !req) return;

  const allHeaderInfo = getHeaderInfo(all, ANCHOR_HEADER_ALL);
  const actHeaderInfo = getHeaderInfo(act, ANCHOR_HEADER_ACT);
  const reqHeaderInfo = getHeaderInfo(req, ANCHOR_HEADER_REQ);
  if (!allHeaderInfo || !actHeaderInfo || !reqHeaderInfo) return;

  const hAll = allHeaderInfo.headerMap;
  const hAct = actHeaderInfo.headerMap;
  const hReq = reqHeaderInfo.headerMap;

  const { index: idxAct } = buildCandidateIndex(act, hAct, actHeaderInfo.dataStartRow);
  const { index: idxAll, rows: allRows } = buildCandidateIndex(all, hAll, allHeaderInfo.dataStartRow);
  const { idx: reqIdx } = buildReqIndex(req, hReq, reqHeaderInfo.dataStartRow);

  const filterJob = (jid) => !jobIds || jobIds.length === 0 || jobIds.includes(jid);

  const toDelete = [];
  idxAct.forEach((row, key) => {
    const [jid] = key.split('|');
    if (!filterJob(jid)) return;
    const inAll = idxAll.has(key);
    let allowed = false;
    if (inAll) {
      const reqRow = reqIdx.get(jid);
      const status = reqRow ? getRowObjectByHeaders(req, hReq, reqRow)[H_REQ.JobStatus] : '';
      allowed = _isOpenForActive_(status);
    }
    if (!inAll || !allowed) toDelete.push(row);
  });
  toDelete.sort((a,b) => b - a).forEach(r => act.deleteRow(r));

  const relevantRows = (jobIds && jobIds.length > 0) 
    ? allRows.filter(r => {
        const jid = _normStr_(r.data[H_ALL.JobID]);
        return jobIds.includes(jid);
      })
    : allRows;

  const processedKeys = new Set();
  for (const { row, data } of relevantRows) {
    const jid = _normStr_(data[H_ALL.JobID]);
    const eml = normEmail(data[H_ALL.Email]);
    if (!jid || !eml || !filterJob(jid)) continue;

    const k = keyFor(jid, eml);
    if (processedKeys.has(k)) {
      logWarn('Duplicate candidate key encountered during reconcile; skipping duplicate', { key: k });
      continue;
    }
    processedKeys.add(k);

    const reqRow = reqIdx.get(jid);
    if (!reqRow) continue;

    const rObj = getRowObjectByHeaders(req, hReq, reqRow);
    const reqStatusCanon = _canonReqStatus_(rObj[H_REQ.JobStatus] || '');
    const reqTitle = rObj[H_REQ.JobTitle] || '';

    const curAll = data;
    const wantAll = {
      [H_ALL.JobTitle]:  reqTitle,
      [H_ALL.JobStatus]: reqStatusCanon
    };
    const diffAll = _diffFields_(wantAll, curAll, [H_ALL.JobTitle, H_ALL.JobStatus]);
    if (Object.keys(diffAll).length) {
      diffAll[H_ALL.Updated] = nowDetroit();
      setRowValuesByHeaders(all, hAll, row, diffAll);
      for (const [k2,v2] of Object.entries(diffAll)) data[k2] = v2;
    }

    const resumeUrl   = hAll[H_ALL.Resume]   ? extractUrl(all, row, hAll[H_ALL.Resume])     : '';
    const linkedinUrl = hAll[H_ALL.LinkedIn] ? extractUrl(all, row, hAll[H_ALL.LinkedIn])   : '';

    if (_isOpenForActive_(reqStatusCanon)) {
      const k = keyFor(jid, eml);
      const wasRecentlyEdited = isRecentlyEdited(SHEET_ACTIVE, k);
      
      if (processedKeys.size % 10 === 0 || processedKeys.size === 1) {
        logInfo('Reconcile progress', { 
          processed: processedKeys.size, 
          total: relevantRows.length,
          lastKey: k
        });
      }
      
      if (!wasRecentlyEdited) {
        const objWithRow = Object.assign({ __row: row }, data);
        _upsertActiveRow_(objWithRow, { resumeUrl, linkedinUrl });
      }
    }
  }
  
  logInfo('Reconcile complete', { 
    jobIds: jobIds && jobIds.length > 0 ? jobIds : 'all', 
    totalCandidates: relevantRows.length,
    synced: processedKeys.size,
    deleted: toDelete.length
  });
}

/**
 * Reconciles all Active Candidates (no job filter)
 */
function reconcileActiveMembership_All_() {
  reconcileActiveMembership_ByJobIds_([]);
}

/**
 * Sweeps Candidate Database and autopopulates job info from requisitions
 */
function Sweep_Autopopulate_All_From_Reqs_() {
  const ss = SpreadsheetApp.getActive();
  const all = ss.getSheetByName(SHEET_ALL);
  const req = ss.getSheetByName(SHEET_REQUISITIONS);
  if (!all || !req) return;

  const allHeaderInfo = getHeaderInfo(all, ANCHOR_HEADER_ALL);
  const reqHeaderInfo = getHeaderInfo(req, ANCHOR_HEADER_REQ);
  if (!allHeaderInfo || !reqHeaderInfo) return;

  const hAll = allHeaderInfo.headerMap;
  const hReq = reqHeaderInfo.headerMap;
  const { rows: allRows } = buildCandidateIndex(all, hAll, allHeaderInfo.dataStartRow);
  const { idx: reqIdx } = buildReqIndex(req, hReq, reqHeaderInfo.dataStartRow);

  let updates = 0;
  for (const { row, data } of allRows) {
    const jid = _normStr_(data[H_ALL.JobID]);
    if (!jid) continue;
    const r = reqIdx.get(jid);
    if (!r) continue;
    const rObj = getRowObjectByHeaders(req, hReq, r);
    const want = {
      [H_ALL.JobTitle]:  rObj[H_REQ.JobTitle]  || '',
      [H_ALL.JobStatus]: _canonReqStatus_(rObj[H_REQ.JobStatus] || '')
    };
    const diff = _diffFields_(want, data, [H_ALL.JobTitle, H_ALL.JobStatus]);
    if (Object.keys(diff).length) {
      diff[H_ALL.Updated] = nowDetroit();
      setRowValuesByHeaders(all, hAll, row, diff);
      updates++;
    }
  }
  logInfo('Sweep_Autopopulate_All_From_Reqs_ complete', { updatedRows: updates });
}