/** @file Util.gs - Logging, locks, header cache, range helpers, URL helpers, row readers/writers. */

// ---------- Logging
/**
 * Logs an informational message to console and SYS_LOGS sheet
 * @param {string} msg - The message to log
 * @param {Object} ctx - Optional context object
 */
function logInfo(msg, ctx) {
  try {
    console.log(`${LOG_PREFIX} ${msg}`, ctx || '');
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(SHEET_SYS_LOG);
    if (!sh) {
      sh = ss.insertSheet(SHEET_SYS_LOG);
      sh.hideSheet();
      sh.getRange(1,1,1,3).setValues([['Timestamp','Level','Message']]);
    }
    sh.appendRow([new Date(), 'INFO', `${msg} ${ctx ? JSON.stringify(ctx) : ''}`]);
  } catch (_) {}
}

/**
 * Logs a warning message to console and SYS_LOGS sheet
 * @param {string} msg - The message to log
 * @param {Object} ctx - Optional context object
 */
function logWarn(msg, ctx) {
  try {
    console.warn(`${LOG_PREFIX} ${msg}`, ctx || '');
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(SHEET_SYS_LOG);
    if (!sh) {
      sh = ss.insertSheet(SHEET_SYS_LOG);
      sh.hideSheet();
      sh.getRange(1,1,1,3).setValues([['Timestamp','Level','Message']]);
    }
    sh.appendRow([new Date(), 'WARN', `${msg} ${ctx ? JSON.stringify(ctx) : ''}`]);
  } catch (_) {}
}

// ---------- Lock + recursion guard (non-blocking; optional onBusy callback)
/**
 * Executes a function with a document lock to prevent concurrent modifications
 * @param {Function} fn - The function to execute
 * @param {number} timeoutMs - Lock timeout in milliseconds
 * @param {Function} onBusy - Optional callback if lock cannot be acquired
 * @returns {*} Result of fn()
 */
function withLock(fn, timeoutMs = LOCK_TIMEOUT_MS, onBusy) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(timeoutMs)) {
    logWarn('Skipped run: lock busy', {});
    if (typeof onBusy === 'function') {
      try { onBusy(); } catch (err) { logWarn('onBusy failed', { message: String(err && err.message || err) }); }
    }
    return;
  }
  try {
    if (StateManager.isRecursionGuardSet()) {
      logWarn('Recursion blocked', {});
      return;
    }
    StateManager.setRecursionGuard();
    return fn();
  } finally {
    StateManager.deleteRecursionGuard();
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

// ---------- Time + keys
/**
 * Returns current date/time in Detroit timezone
 * @returns {Date} Current date in Detroit timezone
 */
function nowDetroit() { 
  return new Date(Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss")); 
}

/**
 * Normalizes email addresses with Gmail-specific rules.
 * - Converts to lowercase and trims whitespace
 * - For Gmail: removes dots from local part and ignores +tags
 * - For other domains: just lowercases and trims
 *
 * Note: Non-Gmail addresses with +tags are NOT normalized,
 * meaning john@company.com and john+tag@company.com are treated
 * as different candidates. This is intentional to avoid false
 * matches across different email systems.
 *
 * @param {string} s - Email address to normalize
 * @returns {string} Normalized email address
 */
function normEmail(s) {
  const email = (s || '').toString().trim().toLowerCase();
  if (!email || !email.includes('@')) return email;
  
  const parts = email.split('@');
  if (parts.length !== 2) return email; // Invalid format
  
  const [local, domain] = parts;
  
  // Only apply Gmail-specific rules to Gmail addresses
  if (domain === 'gmail.com' || domain === 'googlemail.com') {
    const cleanLocal = local.replace(/\./g, '').split('+')[0];
    return `${cleanLocal}@gmail.com`;
  }
  
  // For all other domains, return as-is (already lowercase and trimmed)
  return email;
}

/**
 * Normalizes phone numbers by removing all non-digit characters
 * @param {string} s - Phone number to normalize
 * @returns {string} Phone number with only digits
 */
function normPhone(s) { 
  return (s || '').toString().replace(/\D/g, ''); 
}

/**
 * Creates composite key for candidate indexing.
 * Handles empty values gracefully to prevent key collisions.
 * @param {string} jobId - The job ID
 * @param {string} email - The candidate's email
 * @returns {string} Composite key in format "jobId|normalizedEmail"
 */
function keyFor(jobId, email) { 
  const jid = (jobId || '').toString().trim();
  const em = normEmail(email);
  
  // If both are empty, return a unique marker to prevent collisions
  if (!jid && !em) return '_EMPTY_KEY_' + Date.now();
  
  return `${jid}|${em}`;
}

// ---------- DYNAMIC HEADER MAPPING ----------
/**
 * Retrieves or caches header information for a sheet
 * @param {Sheet} sheet - The sheet to analyze
 * @param {string} anchorHeader - The anchor header to search for
 * @returns {Object|null} Header info object with headerRow, dataStartRow, headerMap
 */
function getHeaderInfo(sheet, anchorHeader) {
  if (!sheet || !anchorHeader) return null;
  const cache = CacheService.getScriptCache();
  const key = CACHE_KEYS.HDR_INFO(sheet.getSheetId());
  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);

  const searchRange = sheet.getRange(1, 1, Math.min(sheet.getMaxRows(), MAX_HEADER_SEARCH_ROWS), sheet.getLastColumn() || 1);
  const values = searchRange.getValues();
  
  let headerRow = -1;
  let headerValues = [];
  for (let i = 0; i < values.length; i++) {
    const rowContainsAnchor = values[i].some(cell => String(cell).trim() === anchorHeader);
    if (rowContainsAnchor) {
      headerRow = i + 1; // 1-based index
      headerValues = values[i];
      break;
    }
  }

  if (headerRow === -1) {
    logWarn('Could not find anchor header in sheet', { sheet: sheet.getName(), anchor: anchorHeader });
    return null;
  }
  
  const headerMap = {};
  headerValues.forEach((h, i) => {
    const t = (h || '').toString().trim();
    if (t) headerMap[t] = i + 1; // 1-based column index
  });
  
  const headerInfo = {
    headerRow: headerRow,
    dataStartRow: headerRow + 1,
    headerMap: headerMap,
  };
  
  cache.put(key, JSON.stringify(headerInfo), CACHE_TTL_MUTATION);
  return headerInfo;
}

/**
 * Invalidates cached header information for a sheet
 * @param {Sheet} sheet - The sheet to invalidate cache for
 */
function invalidateHeaderCache(sheet){
  if (!sheet) return;
  CacheService.getScriptCache().remove(CACHE_KEYS.HDR_INFO(sheet.getSheetId()));
}

// ---------- Recent-edit guard (Active)
/**
 * Marks a row as recently edited to prevent immediate sync conflicts
 * @param {string} sheetName - Name of the sheet
 * @param {string} compositeKey - Composite key for the row
 */
function markRecentEdit(sheetName, compositeKey) {
  CacheService.getScriptCache().put(CACHE_KEYS.MUTE_ROW(sheetName, compositeKey), '1', CACHE_TTL_SHORT);
}

/**
 * Checks if a row was recently edited
 * @param {string} sheetName - Name of the sheet
 * @param {string} compositeKey - Composite key for the row
 * @returns {boolean} True if recently edited
 */
function isRecentlyEdited(sheetName, compositeKey) {
  return !!CacheService.getScriptCache().get(CACHE_KEYS.MUTE_ROW(sheetName, compositeKey));
}

// ---------- Safe range ops
/**
 * Safely executes a range operation with error logging
 * @param {Sheet} sh - The sheet
 * @param {string} headerName - Header name for logging
 * @param {Range} range - The range to operate on
 * @param {string} opName - Name of the operation for logging
 * @param {Function} callback - Function to execute on the range
 * @returns {*} Result of callback
 */
function safeRangeOp(sh, headerName, range, opName, callback) {
  try { 
    return callback(range); 
  } catch (err) {
    logWarn(`Range op failed: ${opName}`, {
      sheet: sh ? sh.getName() : '(unknown)',
      header: headerName || '(unknown)',
      a1: range ? range.getA1Notation() : '(unknown)',
      message: String(err && err.message || err)
    });
  }
}

/**
 * Safely sets data validation on a range
 * @param {Range} range - The range to set validation on
 * @param {DataValidation} rule - The validation rule
 * @param {Object} ctx - Context object with header name
 */
function safeSetDataValidation(range, rule, ctx) {
  const sh = range.getSheet();
  return safeRangeOp(sh, ctx && ctx.header, range, 'setDataValidation', r => r.setDataValidation(rule));
}

// ---------- Optimized Row Writers ----------
/**
 * Sets multiple values in a row by header names, using batched operations
 * @param {Sheet} sh - The sheet
 * @param {Object} hm - Header map (header name -> column number)
 * @param {number} row - Row number to update
 * @param {Object} updates - Object with header names as keys and values to set
 */
function setRowValuesByHeaders(sh, hm, row, updates){
  if (!sh || !hm || !updates) {
    logWarn('setRowValuesByHeaders called with invalid parameters', { 
      hasSheet: !!sh, 
      hasHeaderMap: !!hm, 
      hasUpdates: !!updates 
    });
    return;
  }
  
  const updatesByCol = [];
  for (const [header, value] of Object.entries(updates)) {
    if (hm[header]) {
      updatesByCol.push({ col: hm[header], value: value, header: header });
    }
  }

  if (!updatesByCol.length) return;

  updatesByCol.sort((a, b) => a.col - b.col);

  let i = 0;
  while (i < updatesByCol.length) {
    let startCol = updatesByCol[i].col;
    let headersInRun = [updatesByCol[i].header];
    let valuesInRun = [updatesByCol[i].value];
    
    while (i + 1 < updatesByCol.length && updatesByCol[i + 1].col === startCol + valuesInRun.length) {
      i++;
      headersInRun.push(updatesByCol[i].header);
      valuesInRun.push(updatesByCol[i].value);
    }

    if (valuesInRun.length > 1) {
      const range = sh.getRange(row, startCol, 1, valuesInRun.length);
      safeRangeOp(sh, headersInRun.join(', '), range, 'setValues', r => r.setValues([valuesInRun]));
    } else {
      const range = sh.getRange(row, startCol);
      safeRangeOp(sh, headersInRun[0], range, 'setValue', r => r.setValue(valuesInRun[0]));
    }
    i++;
  }
}

/**
 * Safely sets a single value with error handling
 * @param {Sheet} sh - The sheet
 * @param {string} header - Header name for logging
 * @param {Range} range - The range to set value in
 * @param {*} value - The value to set
 */
function safeSetValue(sh, header, range, value) {
  return safeRangeOp(sh, header, range, 'setValue', r => r.setValue(value));
}

// ---------- Row readers
/**
 * Reads a row and returns an object with header names as keys
 * @param {Sheet} sh - The sheet
 * @param {Object} hm - Header map (header name -> column number)
 * @param {number} row - Row number to read
 * @returns {Object} Object with header names as keys and cell values as values
 */
function getRowObjectByHeaders(sh, hm, row){
  if (!sh || !hm || !row || row < 1) {
    logWarn('getRowObjectByHeaders called with invalid parameters', {
      hasSheet: !!sh,
      hasHeaderMap: !!hm,
      row: row
    });
    return {};
  }
  
  const lastCol = sh.getLastColumn();
  if (Object.keys(hm).length === 0) return {};
  
  const vals = sh.getRange(row, 1, 1, lastCol).getValues()[0];
  const obj = {};
  for (const [h, c] of Object.entries(hm)) {
    if (c - 1 < vals.length) obj[h] = vals[c - 1];
  }
  return obj;
}

// ---------- Index builders
/**
 * Builds an index of candidates from a sheet
 * @param {Sheet} sh - The sheet to index
 * @param {Object} hm - Header map
 * @param {number} dataStartRow - First row of data
 * @returns {Object} Object with index (Map) and rows (Array)
 */
function buildCandidateIndex(sh, hm, dataStartRow){
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const index = new Map(), rows = [];
  if (lastRow < dataStartRow) return { index, rows };

  const numRows = lastRow - dataStartRow + 1;
  const rng = sh.getRange(dataStartRow, 1, numRows, lastCol).getValues();

  for (let i = 0; i < rng.length; i++){
    const r = dataStartRow + i;
    const obj = {};
    for (const [h, c] of Object.entries(hm)) {
      if (c - 1 < rng[i].length) obj[h] = rng[i][c - 1];
    }
    const jid = (obj[H_ALL?.JobID] || obj[H_ACT?.JobID] || '').toString().trim();
    const em  = normEmail(obj[H_ALL?.Email] || obj[H_ACT?.Email] || '');
    if (jid || em) { 
      index.set(keyFor(jid, em), r);
      rows.push({row: r, data: obj});
    }
  }
  return { index, rows };
}

/**
 * Builds an index of requisitions from a sheet
 * @param {Sheet} sh - The sheet to index
 * @param {Object} hm - Header map
 * @param {number} dataStartRow - First row of data
 * @returns {Object} Object with idx (Map) and rows (Array)
 */
function buildReqIndex(sh, hm, dataStartRow){
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const idx = new Map(), rows = [];
  if (lastRow < dataStartRow) return { idx, rows };

  const numRows = lastRow - dataStartRow + 1;
  const rng = sh.getRange(dataStartRow, 1, numRows, lastCol).getValues();

  for (let i = 0; i < rng.length; i++){
    const r = dataStartRow + i;
    const obj = {};
    for (const [h, c] of Object.entries(hm)) {
       if (c - 1 < rng[i].length) obj[h] = rng[i][c - 1];
    }
    const jid = (obj[H_REQ.JobID] || '').toString().trim();
    if (jid) idx.set(jid, r);
    rows.push({row: r, data: obj});
  }
  return { idx, rows };
}

// ---------- URL helpers
/**
 * Extracts URL from a cell (supports formulas, rich text, and plain text)
 * @param {Sheet} sh - The sheet
 * @param {number} row - Row number
 * @param {number} col - Column number
 * @returns {string} The extracted URL or empty string
 */
function extractUrl(sh, row, col){
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
    logWarn('extractUrl failed', { row, col, err: e.message });
  }
  return '';
}

/**
 * Sets a hyperlink with URL validation.
 * @param {Sheet} sh - The sheet
 * @param {number} row - Row number
 * @param {number} col - Column number
 * @param {string} url - The URL to link to
 * @param {string} label - The display label
 */
function setHyperlink(sh, row, col, url, label){
  const u = (url || '').toString().trim();
  if (!u) return;

  // Validate URL format - only allow http/https/mailto/tel protocols
  if (!u.match(/^(https?:\/\/.+|mailto:.+|tel:.+)/i)) {
    logWarn('Invalid URL format, skipping hyperlink', {
      url: u,
      row,
      col,
      sheet: sh.getName()
    });
    return;
  }
  
  // Escape special characters to prevent formula injection
  const escapedUrl = _escapeForFormula_(u);
  const escapedLabel = _escapeForFormula_(label);
  
  safeRangeOp(sh, null, sh.getRange(row, col), 'setFormula',
    rg => rg.setFormula(`=HYPERLINK("${escapedUrl}","${escapedLabel}")`));
}

/**
 * Escapes quotes in strings for safe use in formulas
 * @param {string} s - String to escape
 * @returns {string} Escaped string
 */
function _escapeForFormula_(s) { 
  return String(s).replace(/"/g, '""'); 
}

// ---------- US Holidays for business days calculation
const US_HOLIDAYS = [
  // 2024
  '2024-01-01', // New Year's Day
  '2024-01-15', // MLK Day
  '2024-02-19', // Presidents Day
  '2024-05-27', // Memorial Day
  '2024-07-04', // Independence Day
  '2024-09-02', // Labor Day
  '2024-10-14', // Columbus Day
  '2024-11-11', // Veterans Day
  '2024-11-28', // Thanksgiving
  '2024-12-25', // Christmas
  // 2025
  '2025-01-01', // New Year's Day
  '2025-01-20', // MLK Day
  '2025-02-17', // Presidents Day
  '2025-05-26', // Memorial Day
  '2025-07-04', // Independence Day
  '2025-09-01', // Labor Day
  '2025-10-13', // Columbus Day
  '2025-11-11', // Veterans Day
  '2025-11-27', // Thanksgiving
  '2025-12-25', // Christmas
  // 2026
  '2026-01-01', // New Year's Day
  '2026-01-19', // MLK Day
  '2026-02-16', // Presidents Day
  '2026-05-25', // Memorial Day
  '2026-07-03', // Independence Day (observed)
  '2026-09-07', // Labor Day
  '2026-10-12', // Columbus Day
  '2026-11-11', // Veterans Day
  '2026-11-26', // Thanksgiving
  '2026-12-25', // Christmas
];

/**
 * Calculates business days between two dates, excluding weekends and US holidays.
 * Uses optimized algorithm for large date ranges.
 * @param {Date|string} a - Start date
 * @param {Date|string} b - End date
 * @returns {number} Number of business days
 */
function businessDaysBetween(a, b) {
  if (!a || !b) return 0;
  
  const s = new Date(a), e = new Date(b);
  s.setHours(0, 0, 0, 0); 
  e.setHours(0, 0, 0, 0);
  
  if (e <= s) return 0;
  
  const totalDays = Math.floor((e - s) / (1000 * 60 * 60 * 24));
  
  // Optimize for date ranges > 14 days using full week calculation
  if (totalDays > 14) {
    const fullWeeks = Math.floor(totalDays / 7);
    const remainingDays = totalDays % 7;
    
    let businessDays = fullWeeks * 5; // 5 business days per week
    
    // Add remaining days (excluding weekends)
    const remainingStart = new Date(s);
    remainingStart.setDate(remainingStart.getDate() + (fullWeeks * 7));
    
    for (let i = 0; i < remainingDays; i++) {
      const day = new Date(remainingStart);
      day.setDate(day.getDate() + i);
      const dayOfWeek = day.getDay();
      if (dayOfWeek !== 0 && dayOfWeek !== 6) businessDays++;
    }
    
    // Subtract holidays that fall on business days
    for (const holiday of US_HOLIDAYS) {
      const hDate = new Date(holiday);
      if (hDate >= s && hDate < e) {
        const dayOfWeek = hDate.getDay();
        if (dayOfWeek !== 0 && dayOfWeek !== 6) businessDays--;
      }
    }
    
    return businessDays;
  }
  
  // For short ranges, use simple day-by-day (still fast enough)
  let n = 0;
  for (let d = new Date(s); d < e; d.setDate(d.getDate() + 1)) { 
    const k = d.getDay(); 
    if (k !== 0 && k !== 6) {
      // Check if it's a holiday
      const dateStr = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
      if (!US_HOLIDAYS.includes(dateStr)) {
        n++;
      }
    }
  }
  return n;
}

// ---------- Misc helpers
/**
 * Checks if a value is blank (null, undefined, or empty string)
 * @param {*} v - Value to check
 * @returns {boolean} True if blank
 */
function isBlank(v) { 
  return v === null || v === '' || typeof v === 'undefined'; 
}

/**
 * Displays a toast notification
 * @param {string} msg - Message to display
 * @param {string} title - Optional title
 * @param {number} timeout - Optional timeout in seconds
 */
function toast(msg, title, timeout) { 
  try { 
    const ss = SpreadsheetApp.getActive();
    if (title && timeout !== undefined) {
      ss.toast(msg, title, timeout);
    } else if (title) {
      ss.toast(msg, title);
    } else {
      ss.toast(msg);
    }
  } catch(_) { } 
}

/**
 * Checks if a range intersects with a specific column
 * @param {Range} range - The range to check
 * @param {number} col - The column number to check against
 * @returns {boolean} True if range intersects the column
 */
function rangesIntersectColumns_(range, col) {
  const a = range.getColumn(), b = range.getLastColumn();
  return col >= a && col <= b;
}

// ---------- Template Formatting System (handles empty sheet case)
/**
 * Captures template formatting BEFORE inserting rows.
 * This is critical for empty sheets where the user has pre-formatted
 * the first data row - we must capture it before row insertion shifts it.
 *
 * @param {Sheet} sheet - The sheet to capture from
 * @param {number} dataStartRow - The first valid data row position
 * @returns {Object} Template object with type and formatting data
 */
function captureTemplateFormat(sheet, dataStartRow) {
  const lastRow = sheet.getLastRow();
  const maxCols = sheet.getMaxColumns();

  // Case 1: Sheet has data rows - use last row as template (copy after insert)
  if (lastRow >= dataStartRow) {
    return { type: 'row', row: lastRow };
  }

  // Case 2: Empty sheet - capture formatting from first data row position
  // Users may have pre-formatted this row in the template
  if (sheet.getMaxRows() >= dataStartRow) {
    try {
      const range = sheet.getRange(dataStartRow, 1, 1, maxCols);
      return {
        type: 'saved',
        backgrounds: range.getBackgrounds(),
        fontColors: range.getFontColors(),
        fontFamilies: range.getFontFamilies(),
        fontSizes: range.getFontSizes(),
        fontWeights: range.getFontWeights(),
        fontStyles: range.getFontStyles(),
        horizontalAlignments: range.getHorizontalAlignments(),
        verticalAlignments: range.getVerticalAlignments(),
        numberFormats: range.getNumberFormats(),
        wraps: range.getWraps()
      };
    } catch (e) {
      logWarn('Failed to capture template formatting', { error: e.message });
      return { type: 'none' };
    }
  }

  // Case 3: No template available
  return { type: 'none' };
}

/**
 * Applies captured template formatting to a destination row.
 * Handles both live row reference (for sheets with data) and
 * saved formatting (for empty sheets with pre-formatted template row).
 *
 * @param {Sheet} sheet - The sheet to apply formatting to
 * @param {Object} template - Template object from captureTemplateFormat()
 * @param {number} destRow - The destination row number
 */
function applyTemplateFormat(sheet, template, destRow) {
  if (template.type === 'none') return;

  const maxCols = sheet.getMaxColumns();
  const destRange = sheet.getRange(destRow, 1, 1, maxCols);

  try {
    if (template.type === 'row') {
      // Copy from existing data row
      const srcRange = sheet.getRange(template.row, 1, 1, maxCols);
      srcRange.copyTo(destRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    } else if (template.type === 'saved') {
      // Apply saved formatting (from pre-formatted empty row)
      destRange.setBackgrounds(template.backgrounds);
      destRange.setFontColors(template.fontColors);
      destRange.setFontFamilies(template.fontFamilies);
      destRange.setFontSizes(template.fontSizes);
      destRange.setFontWeights(template.fontWeights);
      destRange.setFontStyles(template.fontStyles);
      destRange.setHorizontalAlignments(template.horizontalAlignments);
      destRange.setVerticalAlignments(template.verticalAlignments);
      destRange.setNumberFormats(template.numberFormats);
      destRange.setWraps(template.wraps);
    }
  } catch (e) {
    logWarn('Failed to apply template formatting', { destRow, error: e.message });
  }
}

/**
 * @deprecated Use captureTemplateFormat() and applyTemplateFormat() instead.
 * Kept for backward compatibility during transition.
 */
function _findTemplateRow_(sheet, dataStartRow, currentLastRow) {
  if (currentLastRow >= dataStartRow) {
    return currentLastRow;
  }
  // NEW: Check if first data row exists (may have pre-set formatting)
  if (sheet.getMaxRows() >= dataStartRow) {
    return dataStartRow;
  }
  return -1;
}

// ---------- Anchor Header Helper (consolidated from DebounceQueue and LinkHygiene)
/**
 * Gets the anchor header for a sheet name.
 * @param {string} sheetName - Name of the sheet
 * @returns {string} The anchor header for that sheet
 */
function getAnchorForSheet(sheetName) {
  const anchors = {
    [SHEET_REQUISITIONS]: ANCHOR_HEADER_REQ,
    [SHEET_ALL]: ANCHOR_HEADER_ALL,
    [SHEET_ACTIVE]: ANCHOR_HEADER_ACT,
  };
  return anchors[sheetName];
}

