/** @file Diagnostics.gs — quick checks for headers & triggers; logs to SYS_LOGS and toast. */

function Diagnose_Setup() {
  const ss = SpreadsheetApp.getActive();
  const checks = [
    { name: SHEET_REQUISITIONS, anchor: ANCHOR_HEADER_REQ, required: [H_REQ.JobID, H_REQ.JobStatus, H_REQ.JobTitle, H_REQ.Opened, H_REQ.Created] },
    { name: SHEET_ALL,          anchor: ANCHOR_HEADER_ALL, required: [H_ALL.JobID, H_ALL.Email, H_ALL.JobStatus, H_ALL.JobTitle] },
    { name: SHEET_ACTIVE,       anchor: ANCHOR_HEADER_ACT, required: [H_ACT.JobID, H_ACT.Email, H_ACT.JobStatus, H_ACT.JobTitle] },
  ];

  for (const c of checks) {
    const sh = ss.getSheetByName(c.name);
    if (!sh) { logWarn('Diagnose: sheet missing', { sheet: c.name }); continue; }

    const headerInfo = getHeaderInfo(sh, c.anchor);
    if (!headerInfo) {
      logWarn('Diagnose: could not find headers', { sheet: c.name, anchor: c.anchor });
      continue;
    }

    const hm = headerInfo.headerMap;
    const missing = c.required.filter(h => !hm[h]);
    if (missing.length) logWarn('Diagnose: headers missing', { sheet: c.name, missing });
    else logInfo('Diagnose: headers OK', { sheet: c.name });

    const headerRowText = sh.getRange(headerInfo.headerRow, 1, 1, Math.max(1, sh.getLastColumn())).getDisplayValues()[0];
    logInfo('Diagnose: header row snapshot', { sheet: c.name, headerRow: headerInfo.headerRow, cells: headerRowText.slice(0, 15) });
  }

  const trig = ScriptApp.getProjectTriggers().map(t => ({ handler: t.getHandlerFunction(), type: String(t.getEventType()) }));
  logInfo('Diagnose: triggers', { triggers: trig });

  toast('Diagnosis complete — see SYS_LOGS.');
}