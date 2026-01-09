/** @file Main.gs - Admin commands with progress feedback. */

function Full_Resync() {
  const ss = SpreadsheetApp.getActive();
  
  const reqSheet = ss.getSheetByName(SHEET_REQUISITIONS);
  const allSheet = ss.getSheetByName(SHEET_ALL);
  
  const reqHeaderInfo = getHeaderInfo(reqSheet, ANCHOR_HEADER_REQ);
  const allHeaderInfo = getHeaderInfo(allSheet, ANCHOR_HEADER_ALL);

  if (!reqHeaderInfo || !reqHeaderInfo.headerMap[H_REQ.JobID] || !reqHeaderInfo.headerMap[H_REQ.JobStatus]) {
    ss.toast('Resync failed: Critical headers are missing from the Requisitions sheet.', 'Error', 10);
    return;
  }
  if (!allHeaderInfo || !allHeaderInfo.headerMap[H_ALL.JobID] || !allHeaderInfo.headerMap[H_ALL.Email]) {
    ss.toast('Resync failed: Critical headers are missing from the Candidate Database sheet.', 'Error', 10);
    return;
  }
  
  ss.toast('Starting full system resync...', 'Please Wait', -1);
  
  withLock(() => {
    invalidateHeaderCache(reqSheet);
    invalidateHeaderCache(allSheet);
    invalidateHeaderCache(ss.getSheetByName(SHEET_ACTIVE));

    const steps = [
      { fn: ensureJobIds_, name: 'Job IDs', desc: 'Assigning IDs' },
      { fn: Recompute_DaysOpen_All, name: 'Days Open', desc: 'Calculating days open' },
      { fn: rebuildAllValidations_, name: 'Validations', desc: 'Applying dropdowns' },
      { fn: Sweep_Autopopulate_All_From_Reqs_, name: 'Job Sync', desc: 'Syncing job details' },
      { fn: reconcileActiveMembership_All_, name: 'Active Sync', desc: 'Updating Active sheet' },
      { fn: linkHygieneSweep_, name: 'Links', desc: 'Cleaning links' }
    ];
    
    const totalSteps = steps.length;
    
    for (let i = 0; i < totalSteps; i++) {
      const step = steps[i];
      const stepNum = i + 1;
      
      ss.toast(
        `${step.desc}...`, 
        `Step ${stepNum}/${totalSteps}: ${step.name}`, 
        -1
      );
      
      try {
        step.fn();
        logInfo(`Resync step completed: ${step.name}`);
      } catch (e) {
        logWarn(`Resync step failed: ${step.name}`, { error: e.message });
        ss.toast(
          `${step.name} encountered an error but sync is continuing. Check SYS_LOGS for details.`,
          'Warning',
          5
        );
      }
    }
    
    ss.toast('Full system resync complete!', 'Success', 5);
    logInfo('Full resync completed successfully', { stepsCompleted: totalSteps });
    
  }, LOCK_TIMEOUT_LONG_MS, () => {
    enqueueAndSchedule_Reconcile(['all']);
    ss.toast('Resync is queued and will run shortly.', 'System Busy', 5);
  });
}

function resetForDistribution() {
  PropertiesService.getDocumentProperties().deleteAllProperties();
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  // Note: Cache entries will expire naturally (short TTL)
  SpreadsheetApp.getActive().toast('Reset complete', 'Ready', 5);
}