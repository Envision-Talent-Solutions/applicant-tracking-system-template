/** @file Requisitions.gs - Job ID generation, status transitions, days-open calculations. */

function _initJobSeqIfNeeded_(sh) {
  const yyyy = new Date().getFullYear();
  if (StateManager.isJobIdSequenceInitialized(yyyy)) return;
  
  const headerInfo = getHeaderInfo(sh, ANCHOR_HEADER_REQ);
  if (!headerInfo) return;
  const { headerMap: hm, dataStartRow } = headerInfo;

  let maxSeq = 0;
  const colId = hm[H_REQ.JobID];
  if (colId) {
    const last = sh.getLastRow();
    if (last >= dataStartRow) {
      const vals = sh.getRange(dataStartRow, colId, last - dataStartRow + 1, 1).getValues();
      for (let i = 0; i < vals.length; i++) {
        const v = (vals[i][0] || '').toString().trim();
        const m = v.match(/^(\d{4})-(\d{4})$/);
        if (m && +m[1] === yyyy) maxSeq = Math.max(maxSeq, +m[2]);
      }
    }
  }
  StateManager.setJobIdSequence(yyyy, maxSeq);
}

function _allocateNextJobIds_(count) {
  const yyyy = new Date().getFullYear();
  let last = StateManager.getJobIdSequence(yyyy);
  const out = [];
  for (let i = 0; i < count; i++) {
    last += 1;
    out.push(`${yyyy}-${String(last).padStart(4,'0')}`);
  }
  StateManager.setJobIdSequence(yyyy, last);
  return out;
}

/**
 * Applies status-driven date stamps based on job status transitions.
 * Handles proper date logic for Open, On Hold, Closed, and Hired statuses.
 */
function applyReqStatusTransitionsForRows_(rows) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_REQUISITIONS);
  if (!sh) return;

  const headerInfo = getHeaderInfo(sh, ANCHOR_HEADER_REQ);
  if (!headerInfo) return;
  const { headerMap: hm, dataStartRow } = headerInfo;
  const now = nowDetroit();

  for (const row of rows) {
    if (row < dataStartRow) continue;
    const rowObj = getRowObjectByHeaders(sh, hm, row);
    const up = {};
    const status = _normStr_(rowObj[H_REQ.JobStatus]);

    // When moving to a non-terminal state, clear the hired name field
    if (status !== 'Hired' && hm[H_REQ.HiredCandidateName] && rowObj[H_REQ.HiredCandidateName]) {
      up[H_REQ.HiredCandidateName] = '';
    }

    switch (status) {
      case 'Open':
        // Only set Opened Date if it's not already set
        if (hm[H_REQ.Opened] && !rowObj[H_REQ.Opened]) {
          up[H_REQ.Opened] = now;
        }
        // Clear On Hold Date when moving to Open (but keep Opened Date)
        if (hm[H_REQ.OnHoldDate] && rowObj[H_REQ.OnHoldDate]) {
          up[H_REQ.OnHoldDate] = '';
        }
        // Clear terminal dates
        if (hm[H_REQ.ClosedDate] && rowObj[H_REQ.ClosedDate]) {
          up[H_REQ.ClosedDate] = '';
        }
        if (hm[H_REQ.PositionHiredDate] && rowObj[H_REQ.PositionHiredDate]) {
          up[H_REQ.PositionHiredDate] = '';
        }
        break;

      case 'On Hold':
        // Only set On Hold Date if not already set
        if (hm[H_REQ.OnHoldDate] && !rowObj[H_REQ.OnHoldDate]) {
          up[H_REQ.OnHoldDate] = now;
        }
        // DO NOT set Opened Date if it doesn't exist
        // (This is the key fix - On Hold can be the first status)
        
        // Clear terminal dates
        if (hm[H_REQ.ClosedDate] && rowObj[H_REQ.ClosedDate]) {
          up[H_REQ.ClosedDate] = '';
        }
        if (hm[H_REQ.PositionHiredDate] && rowObj[H_REQ.PositionHiredDate]) {
          up[H_REQ.PositionHiredDate] = '';
        }
        break;

      case 'Closed':
        if (hm[H_REQ.ClosedDate] && !rowObj[H_REQ.ClosedDate]) {
          up[H_REQ.ClosedDate] = now;
        }
        break;

      case 'Hired':
        if (hm[H_REQ.PositionHiredDate] && !rowObj[H_REQ.PositionHiredDate]) {
          up[H_REQ.PositionHiredDate] = now;
        }
        // Also set the Closed Date when a position is Hired
        if (hm[H_REQ.ClosedDate] && !rowObj[H_REQ.ClosedDate]) {
          up[H_REQ.ClosedDate] = now;
        }
        break;
    }

    if (Object.keys(up).length > 0) {
      setRowValuesByHeaders(sh, hm, row, up);
    }
  }
}

function ensureJobIds_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_REQUISITIONS);
  if (!sh) return;

  const headerInfo = getHeaderInfo(sh, ANCHOR_HEADER_REQ);
  if (!headerInfo) return;
  const { headerMap: hm, dataStartRow } = headerInfo;

  const last = sh.getLastRow();
  if (last < dataStartRow) return;

  const colId = hm[H_REQ.JobID], colTitle = hm[H_REQ.JobTitle], colStatus = hm[H_REQ.JobStatus];
  if (!colId || !colTitle || !colStatus) return;

  _initJobSeqIfNeeded_(sh);

  const values = sh.getRange(dataStartRow, 1, last - dataStartRow + 1, sh.getLastColumn()).getValues();

  let need = 0;
  for (let i = 0; i < values.length; i++) {
    const jobId = (values[i][colId-1] || '').toString().trim();
    const title = (values[i][colTitle-1] || '').toString().trim();
    const status = (values[i][colStatus-1] || '').toString().trim();
    if ((title || status) && !jobId) {
      need++;
    }
  }
  const newIds = need > 0 ? _allocateNextJobIds_(need) : [];
  
  let newIdPtr = 0;
  for (let i = 0; i < values.length; i++) {
    const row = dataStartRow + i;
    const rowObj = {};
    Object.entries(hm).forEach(([h, c]) => {
      if(c-1 < values[i].length) rowObj[h] = values[i][c-1];
    });
    
    const updates = {};
    const title = _normStr_(rowObj[H_REQ.JobTitle]);
    const status = _normStr_(rowObj[H_REQ.JobStatus]);

    if ((title || status) && !_normStr_(rowObj[H_REQ.JobID])) {
      if (newIdPtr < newIds.length) {
        updates[H_REQ.JobID] = newIds[newIdPtr++];
        if(hm[H_REQ.Created] && !rowObj[H_REQ.Created]) {
           updates[H_REQ.Created] = nowDetroit();
        }
      }
    }
    
    if (Object.keys(updates).length > 0) {
      setRowValuesByHeaders(sh, hm, row, updates);
    }
  }
}

function Recompute_DaysOpen_Rows_(rows) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_REQUISITIONS);
  if (!sh || !rows || !rows.length) return;
  
  const headerInfo = getHeaderInfo(sh, ANCHOR_HEADER_REQ);
  if (!headerInfo) return;
  const { headerMap: hm, dataStartRow } = headerInfo;

  if (!hm[H_REQ.DaysOpen] || !hm[H_REQ.Opened] || !hm[H_REQ.JobStatus]) return;

  for (const row of rows) {
    if (row < dataStartRow) continue;
    const vals = getRowObjectByHeaders(sh, hm, row);
    const status = _normStr_(vals[H_REQ.JobStatus]);
    const opened = vals[H_REQ.Opened] || '';
    const closed = vals[H_REQ.ClosedDate] || '';
    const hired  = vals[H_REQ.PositionHiredDate] || '';
    
    let endDate = new Date();
    if (status !== 'Open' && status !== 'On Hold') {
      if (closed) endDate = new Date(closed);
      else if (hired) endDate = new Date(hired);
    }
    const days = opened ? businessDaysBetween(opened, endDate) : 0;
    if (String(days) !== String(vals[H_REQ.DaysOpen])) {
      setRowValuesByHeaders(sh, hm, row, { [H_REQ.DaysOpen]: days });
    }
  }
}

function Recompute_DaysOpen_All() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_REQUISITIONS);
  if (!sh) return;
  
  const headerInfo = getHeaderInfo(sh, ANCHOR_HEADER_REQ);
  if (!headerInfo) return;
  const { dataStartRow } = headerInfo;

  const last = sh.getLastRow();
  if (last < dataStartRow) return;

  const allRows = Array.from({length: last - dataStartRow + 1}, (_, i) => dataStartRow + i);
  Recompute_DaysOpen_Rows_(allRows);
}