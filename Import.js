/** @file Import.gs - Bulk import routine for candidate data. */

function Bulk_Import_AllCandidates(rows) {
  if (!Array.isArray(rows)) {
    throw new Error(
      'Import failed: Invalid data format.\n\n' +
      'The import data must be a list of candidate records. Please check your import file format.'
    );
  }

  return withLock(() => {
    const ss = SpreadsheetApp.getActive();
    const all = ss.getSheetByName(SHEET_ALL);
    if (!all) {
      throw new Error(
        'Import failed: "Candidate Database" sheet not found.\n\n' +
        'Please ensure your spreadsheet has a sheet named "Candidate Database".'
      );
    }

    const headerInfo = getHeaderInfo(all, ANCHOR_HEADER_ALL);
    if (!headerInfo) {
      throw new Error(
        'Import failed: Cannot find headers in "Candidate Database" sheet.\n\n' +
        'Please ensure the sheet has a header row with "Full Name" in column A.'
      );
    }
    const { headerMap: hm, dataStartRow, headerRow } = headerInfo;

    if (!rows[0] || !rows[0][H_ALL.FullName] || !rows[0][H_ALL.Email]) {
      throw new Error(
        'Import failed: Your data is missing required columns.\n\n' +
        'Each candidate record must include "Full Name" and "Email Address" fields.'
      );
    }

    const { rows: existing } = buildCandidateIndex(all, hm, dataStartRow);
    const emails = new Set();
    for (const { data } of existing) {
      const em = normEmail(data[H_ALL.Email]);
      if (em) emails.add(em);
    }

    const toAppend = [];
    let skippedCount = 0;
    let skippedNoEmail = 0;

    rows.forEach((r) => {
      const rowObj = {};
      for (const [h, v] of Object.entries(r)) {
        if (!hm[h]) continue;
        rowObj[h] = v;
      }
      if (rowObj[H_ALL.FullName]) rowObj[H_ALL.FullName] = String(rowObj[H_ALL.FullName]).trim();
      if (rowObj[H_ALL.Email])   rowObj[H_ALL.Email]   = String(rowObj[H_ALL.Email]).trim();

      const em = normEmail(rowObj[H_ALL.Email]);

      // Track why records are skipped
      if (!em) {
        skippedNoEmail++;
        return;
      }
      if (emails.has(em)) {
        skippedCount++;
        return;
      }

      emails.add(em);

      rowObj[H_ALL.Source] = 'Resume Import';
      rowObj[H_ALL.Created] = nowDetroit();
      rowObj[H_ALL.Updated] = nowDetroit();

      toAppend.push(rowObj);
    });

    if (toAppend.length > 0) {
      const lastRow = Math.max(all.getLastRow(), headerRow);

      // Capture template formatting before inserting rows
      const template = captureTemplateFormat(all, dataStartRow);

      all.insertRowsAfter(lastRow, toAppend.length);

      const newRowIndices = [];
      for(let i = 0; i < toAppend.length; i++) {
        const rowObj = toAppend[i];
        const rowIdx = lastRow + 1 + i;
        newRowIndices.push(rowIdx);

        applyTemplateFormat(all, template, rowIdx);

        setRowValuesByHeaders(all, hm, rowIdx, rowObj);

        if (hm[H_ALL.Resume] && rowObj[H_ALL.Resume]) {
          setHyperlink(all, rowIdx, hm[H_ALL.Resume], rowObj[H_ALL.Resume], 'Resume');
        }
        if (hm[H_ALL.LinkedIn] && rowObj[H_ALL.LinkedIn]) {
          setHyperlink(all, rowIdx, hm[H_ALL.LinkedIn], rowObj[H_ALL.LinkedIn], 'LinkedIn Profile');
        }
      }
      
      // Apply data validations to new rows
      try {
        _applyValidationsToRows_(all, headerInfo, newRowIndices);
      } catch (e) {
        logWarn('Failed to apply validations after import', { error: e.message });
      }
      
      // Collect unique Job IDs for batch processing
      const jobIdsToReconcile = toAppend
        .map(r => r[H_ALL.JobID])
        .filter(id => id && String(id).trim() !== '');
      const uniqueJobIds = Array.from(new Set(jobIdsToReconcile));

      // Autopopulate job details and reconcile Active Candidates
      if (newRowIndices.length > 0) {
        autopopulateAllFromJobId_(newRowIndices);
      }
      if (uniqueJobIds.length > 0) {
        reconcileActiveMembership_ByJobIds_(uniqueJobIds);
      }
    }
    
    return {
      added: toAppend.length,
      skippedDuplicate: skippedCount,
      skippedNoEmail: skippedNoEmail,
      total: rows.length
    };
  });
}

/**
 * Applies data validations to specific rows based on SETTINGS sheet.
 * @param {Sheet} sheet - The sheet to apply validations to
 * @param {Object} headerInfo - Header information object
 * @param {number[]} rows - Array of row numbers to apply validations to
 */
function _applyValidationsToRows_(sheet, headerInfo, rows) {
  const configurations = getDropdownConfigurations();
  if (configurations.size === 0) {
    logInfo('No configurations to apply from SETTINGS sheet');
  }
  
  const hm = headerInfo.headerMap;
  
  // Apply SETTINGS-based validations
  for (const header of Object.keys(hm)) {
    if (configurations.has(header)) {
      const col = hm[header];
      const optionsFromSettings = configurations.get(header);
      
      if (optionsFromSettings && optionsFromSettings.length > 0) {
        const rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(optionsFromSettings, true)
          .setAllowInvalid(true)
          .build();
        
        for (const row of rows) {
          const cell = sheet.getRange(row, col);
          safeSetDataValidation(cell, rule, { header });
        }
      }
    }
  }
  
  // Apply Job ID validation from Requisitions (active positions only)
  if (hm[H_ALL.JobID]) {
    const jobIds = _getActiveJobIds_();
    
    if (jobIds.length > 0) {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(jobIds, true)
        .setAllowInvalid(true)
        .build();
      
      for (const row of rows) {
        const cell = sheet.getRange(row, hm[H_ALL.JobID]);
        safeSetDataValidation(cell, rule, { header: H_ALL.JobID });
      }
    } else {
      logInfo('No active Job IDs available - skipping Job ID validation during import');
    }
  }
}