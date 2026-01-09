/** @file LinkHygiene.gs - Link normalization and cleanup utilities. */

// Full sweep (menu or Full Resync)
function linkHygieneSweep_() {
  const ss = SpreadsheetApp.getActive();
  const sheets = [
    { name: SHEET_ALL,    anchor: ANCHOR_HEADER_ALL, resumeH: H_ALL.Resume,  linkedinH: H_ALL.LinkedIn },
    { name: SHEET_ACTIVE, anchor: ANCHOR_HEADER_ACT, resumeH: H_ACT.Resume,  linkedinH: H_ACT.LinkedIn }
  ];
  for (const { name, anchor, resumeH, linkedinH } of sheets) {
    const sh = ss.getSheetByName(name);
    if (!sh) continue;

    const headerInfo = getHeaderInfo(sh, anchor);
    if (!headerInfo) continue;
    const { headerMap: hm, dataStartRow } = headerInfo;

    const lastRow = sh.getLastRow();
    if (lastRow < dataStartRow) continue;
    
    if (hm[resumeH])  _normalizeColumnLinks_URLOnly_(sh, hm[resumeH], dataStartRow, lastRow, 'Resume');
    if (hm[linkedinH]) _normalizeColumnLinks_URLOnly_(sh, hm[linkedinH], dataStartRow, lastRow, 'LinkedIn Profile');
  }
}

function _normalizeBlock_URLOnly_(sh, startRow, col, height, label) {
  // Collect contiguous runs of rows that actually contain a URL; only those runs are written.
  const rowsWithUrl = [];
  for (let i = 0; i < height; i++) {
    const row = startRow + i;
    const url = _extractUrlSafe_(sh, row, col);
    if (url) rowsWithUrl.push({ row, url });
  }
  if (!rowsWithUrl.length) return;

  // Group into contiguous runs for batched setFormulas
  let i = 0;
  while (i < rowsWithUrl.length) {
    let runStart = rowsWithUrl[i].row, runEnd = runStart, parts = [rowsWithUrl[i].url];
    while (i+1 < rowsWithUrl.length && rowsWithUrl[i+1].row === runEnd + 1) {
      i++; runEnd = rowsWithUrl[i].row; parts.push(rowsWithUrl[i].url);
    }
    const h = runEnd - runStart + 1;
    const formulas = new Array(h);
    for (let k = 0; k < h; k++) {
      const u = _escapeForFormula_(_normalizeUrl_(parts[k]));
      const l = _escapeForFormula_(label);
      formulas[k] = [`=HYPERLINK("${u}","${l}")`];
    }
    
    try {
      sh.getRange(runStart, col, h, 1).setFormulas(formulas);
    } catch (e) {
      logWarn('Failed to apply hyperlink formulas', {
        sheet: sh.getName(),
        startRow: runStart,
        height: h,
        error: e.message
      });
    }
    i++;
  }
}

function _normalizeColumnLinks_URLOnly_(sh, col, startRow, endRow, label) {
  const height = Math.max(0, endRow - startRow + 1);
  if (!height) return;
  _normalizeBlock_URLOnly_(sh, startRow, col, height, label);
}

// -------------------- Helpers --------------------

function _extractUrlSafe_(sh, row, col) {
  try {
    const r = sh.getRange(row, col);
    const rtv = r.getRichTextValue();
    if (rtv) {
      const u = rtv.getLinkUrl();
      if (u) return String(u);
    }
    const f = r.getFormula();
    if (f && /^=HYPERLINK\(/i.test(f)) {
      const m = f.match(/=HYPERLINK\(\s*"([^"]+)"/i);
      if (m) return String(m[1] || '');
    }
    const v = r.getValue();
    if (typeof v === 'string' && /^https?:\/\//i.test(v)) return v.trim();
  } catch (e) {
    logWarn('extract url safe failed', { row, col, msg: e && e.message });
  }
  return '';
}

function _normalizeUrl_(u) {
  const s = (u || '').toString().trim();
  if (!s) return s;
  if (/^www\./i.test(s)) return 'https://' + s;
  if (/^linkedin\.com\//i.test(s)) return 'https://' + s;
  return s;
}