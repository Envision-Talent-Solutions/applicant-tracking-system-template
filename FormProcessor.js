/**
 * @file FormProcessor.gs - Handles incoming Google Form submissions.
 */

/**
 * Triggered by a form submission. Processes the submitted data and adds it to the Candidate Database sheet.
 * @param {Object} e The event object from the onFormSubmit trigger.
 */
function processFormSubmission(e) {
  withLock(() => {
    logInfo('Form submission received', e ? e.namedValues : 'No event data');
    if (!e || !e.namedValues) {
      logWarn('processFormSubmission called without valid event data.');
      return;
    }

    const ss = SpreadsheetApp.getActive();
    const allSheet = ss.getSheetByName(SHEET_ALL);
    if (!allSheet) {
      logWarn('processFormSubmission failed: Candidate Database sheet not found.');
      return;
    }

    const headerInfo = getHeaderInfo(allSheet, ANCHOR_HEADER_ALL);
    if (!headerInfo) {
      logWarn('processFormSubmission failed: Could not find headers in Candidate Database sheet.');
      return;
    }
    const { headerMap: hm, dataStartRow, headerRow } = headerInfo;

    // De-duplication check
    const submittedEmail = normEmail((e.namedValues[H_ALL.Email] || [])[0]);
    const submittedJobId = ((e.namedValues[H_ALL.JobID] || [])[0] || '').toString().trim();
    
    if (!submittedEmail) {
      logWarn('Form submission blocked: Email address is missing.');
      return;
    }
    
    if (!submittedJobId) {
      logWarn('Form submission blocked: Job ID is missing.');
      return;
    }

    const { index: candIndex } = buildCandidateIndex(allSheet, hm, dataStartRow);
    
    // Check for Job ID + Email combo, not just email
    const compositeKey = keyFor(submittedJobId, submittedEmail);
    const existingCandidate = candIndex.has(compositeKey);

    if (existingCandidate) {
      logWarn('Form submission blocked: Duplicate job+email combination.', { 
        email: submittedEmail, 
        jobId: submittedJobId 
      });
      toast('Submission blocked: You have already applied for this position.');
      return;
    }

    // Prepare new row
    const newRowObj = {};
    for (const header of Object.keys(hm)) {
      const formValue = (e.namedValues[header] || [])[0];
      if (formValue !== undefined && formValue !== null) {
        newRowObj[header] = String(formValue).trim();
      }
    }

    // Automatically set the candidate source for all form submissions
    newRowObj[H_ALL.Source] = 'Career Site (Form)';

    // Add timestamps
    newRowObj[H_ALL.Created] = nowDetroit();
    newRowObj[H_ALL.Updated] = nowDetroit();

    // Append to sheet
    const template = captureTemplateFormat(allSheet, dataStartRow);
    const lastRow = Math.max(allSheet.getLastRow(), headerRow);
    allSheet.insertRowAfter(lastRow);
    const newRowIdx = lastRow + 1;
    applyTemplateFormat(allSheet, template, newRowIdx);

    setRowValuesByHeaders(allSheet, hm, newRowIdx, newRowObj);

    // Set hyperlinks for link fields (with null checks)
    if (hm[H_ALL.Resume] && newRowObj[H_ALL.Resume]) {
      setHyperlink(allSheet, newRowIdx, hm[H_ALL.Resume], newRowObj[H_ALL.Resume], 'Resume');
    }
    if (hm[H_ALL.LinkedIn] && newRowObj[H_ALL.LinkedIn]) {
      setHyperlink(allSheet, newRowIdx, hm[H_ALL.LinkedIn], newRowObj[H_ALL.LinkedIn], 'LinkedIn Profile');
    }

    // Apply data validations to the new row
    try {
      _applyValidationsToFormRow_(allSheet, headerInfo, newRowIdx);
    } catch (e) {
      logWarn('Failed to apply validations to form submission row', { error: e.message });
    }

    logInfo('Successfully processed and added new candidate.', { email: submittedEmail, row: newRowIdx, jobId: submittedJobId });

    // Autopopulate job details and reconcile
    autopopulateAllFromJobId_([newRowIdx]);
    const jobId = (newRowObj[H_ALL.JobID] || '').toString().trim();
    if (jobId) {
      reconcileActiveMembership_ByJobIds_([jobId]);
    }

    toast('New candidate successfully submitted and added to the ATS.');
  });
}

/**
 * Applies data validations from SETTINGS sheet to a form submission row.
 * @param {Sheet} sheet - The sheet to apply validations to
 * @param {Object} headerInfo - Header information object
 * @param {number} row - The row number to apply validations to
 */
function _applyValidationsToFormRow_(sheet, headerInfo, row) {
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
          .setAllowInvalid(false)
          .build();
        
        const cell = sheet.getRange(row, col);
        safeSetDataValidation(cell, rule, { header });
      }
    }
  }
  
  // Also apply Job ID dropdown from Requisitions
  if (hm[H_ALL.JobID]) {
    const ss = SpreadsheetApp.getActive();
    const reqSheet = ss.getSheetByName(SHEET_REQUISITIONS);
    if (reqSheet) {
      const reqHeaderInfo = getHeaderInfo(reqSheet, ANCHOR_HEADER_REQ);
      if (reqHeaderInfo) {
        const { rows } = buildReqIndex(reqSheet, reqHeaderInfo.headerMap, reqHeaderInfo.dataStartRow);
        const jobIds = rows
          .map(r => r.data[H_REQ.JobID])
          .filter(id => id && String(id).trim() !== '')
          .map(String);
        
        if (jobIds.length > 0) {
          const rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(jobIds, true)
            .setAllowInvalid(true)
            .build();
          
          const cell = sheet.getRange(row, hm[H_ALL.JobID]);
          safeSetDataValidation(cell, rule, { header: H_ALL.JobID });
        } else {
          logInfo('No Job IDs available - skipping Job ID validation for form submission');
        }
      }
    }
  }
}

/**
 * Time-driven function to refresh the Job ID choices in the live Google Form.
 */
function refreshJobIdChoicesInForm() {
  withLock(() => {
    const formId = StateManager.getFormId();
    if (!formId) {
      logInfo('No form configured - skipping Job ID refresh');
      return;
    }

    let form;
    try {
      form = FormApp.openById(formId);
    } catch (err) {
      // Check if this is a permissions error
      if (isAuthorizationError && isAuthorizationError(err.message)) {
        logWarn('Forms permission not granted - cannot refresh Job IDs in form. ' +
                'Grant Forms access via System Admin > Authorization if you want automatic Job ID updates.',
                { formId, error: err.message });
      } else {
        logWarn('Could not open Google Form to refresh Job IDs. The form may have been deleted or moved.',
                { formId, error: err.message });
      }
      return;
    }

    const jobIdItem = form.getItems().find(item => item.getTitle() === H_ALL.JobID);
    if (!jobIdItem || jobIdItem.getType() !== FormApp.ItemType.LIST) {
      logWarn('Could not find a dropdown/list item with the title "Job ID" in the form.');
      return;
    }
    const jobIdListItem = jobIdItem.asListItem();

    // Fetch open jobs from Requisitions sheet
    const ss = SpreadsheetApp.getActive();
    const reqSheet = ss.getSheetByName(SHEET_REQUISITIONS);
    if (!reqSheet) {
      logWarn('Cannot refresh Job IDs in form: Requisitions sheet not found.');
      return;
    }
    
    const reqHeaderInfo = getHeaderInfo(reqSheet, ANCHOR_HEADER_REQ);
    if (!reqHeaderInfo) return;

    const { rows } = buildReqIndex(reqSheet, reqHeaderInfo.headerMap, reqHeaderInfo.dataStartRow);
    const choices = rows
      .filter(r => OPEN_STATUSES.has((r.data[H_REQ.JobStatus] || '').toString().trim()))
      .map(r => r.data[H_REQ.JobID])
      .filter(Boolean)
      .map(String);

    if (choices.length > 0) {
      jobIdListItem.setChoiceValues(choices);
      logInfo('Successfully refreshed Job ID choices in the live Google Form.', { count: choices.length });
    } else {
      // It's valid to have no open jobs, so clear the list if that's the case.
      jobIdListItem.setChoiceValues([]);
      logInfo('No open jobs found. Cleared Job ID list in the form.');
    }
  });
}