/** @file DebounceQueue.gs - Debounced queues for reconciliation and link hygiene. */

/**
 * Adds Job IDs to a queue with race condition protection.
 * @param {string[]} jobIds - Array of Job IDs to process
 */
function enqueueAndSchedule_Reconcile(jobIds) {
  // Use a mini-lock for queue operations to prevent race conditions
  const scriptLock = LockService.getScriptLock();
  const gotLock = scriptLock.tryLock(2000);
  
  if (!gotLock) {
    logWarn('Could not acquire queue lock, skipping enqueue', { jobIds });
    return;
  }
  
  try {
    const dp = PropertiesService.getDocumentProperties();
    const key = PROP_QUEUE_JOBIDS;

    // If a full resync is already queued ('all'), we don't need to add specific IDs.
    const existing = dp.getProperty(key);
    if (existing === 'all') {
      _scheduleReconcileTrigger();
      return;
    }
    
    // Parse existing queue atomically
    let existingIds = [];
    try {
      existingIds = JSON.parse(existing || '[]');
      if (!Array.isArray(existingIds)) {
        logWarn('Queue corrupted, resetting', { existing });
        existingIds = [];
      }
    } catch (e) {
      logWarn('Queue parse error, resetting', { existing, error: e.message });
      existingIds = [];
    }
    
    // Add new job IDs to the existing queue, ensuring uniqueness
    const newIdSet = new Set([...existingIds, ...jobIds]);
    
    // Write atomically
    try {
      dp.setProperty(key, JSON.stringify(Array.from(newIdSet)));
    } catch (e) {
      logWarn('Failed to update queue', { error: e.message });
      return;
    }

    _scheduleReconcileTrigger();
  } finally {
    scriptLock.releaseLock();
  }
}

/**
 * Worker function that processes the queued reconciliation.
 */
function Debounced_Reconcile_() {
  withLock(() => {
    const dp = PropertiesService.getDocumentProperties();
    const key = PROP_QUEUE_JOBIDS;
    const scheduledKey = PROP_QUEUE_SCHEDULED;
    
    try {
      const jobIdsRaw = dp.getProperty(key);
      if (!jobIdsRaw) {
        logInfo('Queue empty, nothing to reconcile', {});
        return;
      }

      if (jobIdsRaw === 'all') {
        logInfo('Running debounced Full Resync.', {});
        reconcileActiveMembership_All_();
      } else {
        let jobIds = [];
        try {
          jobIds = JSON.parse(jobIdsRaw);
        } catch (e) {
          logWarn('Could not parse queue, skipping reconcile', { jobIdsRaw, error: e.message });
          return;
        }
        
        if (!Array.isArray(jobIds) || jobIds.length === 0) {
          logInfo('Queue empty or invalid, nothing to reconcile', { jobIdsRaw });
          return;
        }

        logInfo('Running debounced reconcile', { jobIds: jobIds, count: jobIds.length });
        
        // Improved Days Open recompute
        try {
          const ss = SpreadsheetApp.getActive();
          const reqSheet = ss.getSheetByName(SHEET_REQUISITIONS);
          if (reqSheet) {
            const reqHeaderInfo = getHeaderInfo(reqSheet, ANCHOR_HEADER_REQ);
            if (reqHeaderInfo) {
              const { idx } = buildReqIndex(reqSheet, reqHeaderInfo.headerMap, reqHeaderInfo.dataStartRow);
              const rowsToUpdate = jobIds.map(id => idx.get(id)).filter(Boolean);
              if (rowsToUpdate.length > 0) {
                Recompute_DaysOpen_Rows_(rowsToUpdate);
              }
            }
          }
        } catch (e) {
          logWarn('Days Open recompute failed', { error: e.message });
        }

        // Run reconciliation
        try {
          reconcileActiveMembership_ByJobIds_(jobIds);
        } catch (e) {
          logWarn('Reconciliation failed', { error: e.message, jobIds });
        }
      }
    } catch (e) {
      logWarn('Debounced_Reconcile_ failed', { error: e.message });
    } finally {
      // Always clean up, even if errors occurred
      try {
        dp.deleteProperty(key);
        dp.deleteProperty(scheduledKey);
        _deleteReconcileTrigger();
      } catch (e) {
        logWarn('Cleanup failed', { error: e.message });
      }
    }
  }, 8000); // Longer timeout for reconciliation
}

/**
 * Schedules a reconcile trigger with duplicate prevention.
 */
function _scheduleReconcileTrigger() {
  const dp = PropertiesService.getDocumentProperties();
  const scheduledKey = PROP_QUEUE_SCHEDULED;

  // Clean up any orphaned triggers before scheduling
  _cleanupOrphanedTriggers_('Debounced_Reconcile_');
  
  // Check if already scheduled to prevent duplicate triggers
  const existingTriggerId = dp.getProperty(scheduledKey);
  if (existingTriggerId) {
    // Verify trigger still exists
    const triggers = ScriptApp.getProjectTriggers();
    const triggerExists = triggers.some(t => t.getUniqueId() === existingTriggerId);
    if (triggerExists) {
      logInfo('Reconcile trigger already scheduled', { triggerId: existingTriggerId });
      return;
    } else {
      // Orphaned trigger ID - clean up
      dp.deleteProperty(scheduledKey);
      logInfo('Removed orphaned trigger ID from properties', { triggerId: existingTriggerId });
    }
  }

  try {
    const trigger = ScriptApp.newTrigger('Debounced_Reconcile_')
      .timeBased()
      .after(DEBOUNCE_RECONCILE_MS)
      .create();

    dp.setProperty(scheduledKey, trigger.getUniqueId());
    logInfo('Scheduled reconcile trigger', { triggerId: trigger.getUniqueId() });
  } catch (e) {
    // Check if this is an authorization error
    if (typeof isAuthorizationError === 'function' && isAuthorizationError(e.message)) {
      logWarn('Cannot schedule trigger - Script Triggers permission not granted. ' +
              'Background sync is disabled. Grant permission via System Admin > Authorization.',
              { error: e.message });
    } else {
      logWarn('Failed to schedule trigger', { error: e.message });
    }
  }
}

/**
 * Deletes the scheduled reconcile trigger.
 */
function _deleteReconcileTrigger() {
  const dp = PropertiesService.getDocumentProperties();
  const scheduledKey = PROP_QUEUE_SCHEDULED;
  const triggerId = dp.getProperty(scheduledKey);
  if (!triggerId) return;

  try {
    const triggers = ScriptApp.getProjectTriggers();
    let deleted = false;
    
    for (const trigger of triggers) {
      if (trigger.getUniqueId() === triggerId) {
        ScriptApp.deleteTrigger(trigger);
        deleted = true;
        break;
      }
    }
    
    if (deleted) {
      logInfo('Deleted reconcile trigger', { triggerId });
    }
  } catch (e) {
    logWarn('Failed to delete trigger', { error: e.message, triggerId });
  } finally {
    dp.deleteProperty(scheduledKey);
  }
}


/**
 * Schedules link hygiene with duplicate prevention.
 */
function scheduleDebouncedLinkHygiene_() {
  const dp = PropertiesService.getDocumentProperties();
  const key = 'ATS:queue:link:scheduled';

  // Clean up any orphaned link hygiene triggers
  _cleanupOrphanedTriggers_('Debounced_LinkHygiene_');
  
  // Check if already scheduled
  const existing = dp.getProperty(key);
  if (existing) {
    const scheduledTime = Number(existing);
    const now = Date.now();
    // If scheduled less than 2 seconds ago, don't schedule again
    if (!isNaN(scheduledTime) && (now - scheduledTime) < 2000) {
      return;
    }
  }
  
  try {
    ScriptApp.newTrigger('Debounced_LinkHygiene_').timeBased().after(DEBOUNCE_LINK_HYGIENE_MS).create();
    dp.setProperty(key, String(Date.now()));
    logInfo('Scheduled Debounced_LinkHygiene_', {});
  } catch (e) {
    logWarn('Failed to schedule link hygiene', { error: e.message });
  }
}

/**
 * Worker function that processes dirty link cleanup.
 */
function Debounced_LinkHygiene_() {
  withLock(() => {
    const dp = PropertiesService.getDocumentProperties();
    const key = 'ATS:queue:link:scheduled';
    
    try {
      _consumeLinkDirtySet_(SHEET_ALL, H_ALL.Resume, 'Resume');
      _consumeLinkDirtySet_(SHEET_ALL, H_ALL.LinkedIn, 'LinkedIn Profile');
      _consumeLinkDirtySet_(SHEET_ACTIVE, H_ACT.Resume, 'Resume');
      _consumeLinkDirtySet_(SHEET_ACTIVE, H_ACT.LinkedIn, 'LinkedIn Profile');

      // Check if more work remains
      if (_hasAnyLinkDirty_()) {
        scheduleDebouncedLinkHygiene_();
      }
    } catch (e) {
      logWarn('Link hygiene failed', { error: e.message });
    } finally {
      dp.deleteProperty(key);
    }
  }, 5000);
}

/**
 * Processes dirty links for a specific sheet and header with batching.
 */
function _consumeLinkDirtySet_(sheetName, header, label) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;

  const headerInfo = getHeaderInfo(sh, getAnchorForSheet(sheetName));
  if (!headerInfo || !headerInfo.headerMap[header]) return;

  const col = headerInfo.headerMap[header];
  const dirtyProps = StateManager.getDirtyLinkProperties();
  const prefix = `ATS:linkdirty:${sheetName}:${header}:`;
  
  const rowsToProcess = [];
  for (const key in dirtyProps) {
    if (key.startsWith(prefix)) {
      const row = Number(key.substring(prefix.length));
      if (!isNaN(row) && row >= headerInfo.dataStartRow) {
        rowsToProcess.push(row);
      }
      StateManager.deleteLinkDirtyProperty(key);
    }
  }

  if (rowsToProcess.length === 0) return;

  // Batch process contiguous rows
  rowsToProcess.sort((a, b) => a - b);
  
  let i = 0;
  while (i < rowsToProcess.length) {
    const startRow = rowsToProcess[i];
    let endRow = startRow;
    
    // Find contiguous block
    while (i + 1 < rowsToProcess.length && rowsToProcess[i + 1] === endRow + 1) {
      i++;
      endRow = rowsToProcess[i];
    }
    
    // Process this block
    try {
      _normalizeColumnLinks_URLOnly_(sh, col, startRow, endRow, label);
    } catch (e) {
      logWarn('Link normalization failed for block', { 
        sheet: sheetName, 
        startRow, 
        endRow, 
        error: e.message 
      });
    }
    
    i++;
  }
}

/**
 * Check if any dirty link flags exist
 */
function _hasAnyLinkDirty_() {
  const dirtyProps = StateManager.getDirtyLinkProperties();
  return Object.keys(dirtyProps).length > 0;
}

/**
 * Cleans up orphaned triggers for a specific function.
 * Triggers are considered orphaned if they exist but are not tracked in properties
 * @param {string} handlerFunction - The handler function name to clean up
 */
function _cleanupOrphanedTriggers_(handlerFunction) {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const dp = PropertiesService.getDocumentProperties();
    
    // Get tracked trigger IDs
    const trackedIds = new Set();
    if (handlerFunction === 'Debounced_Reconcile_') {
      const id = dp.getProperty(PROP_QUEUE_SCHEDULED);
      if (id) trackedIds.add(id);
    } else if (handlerFunction === 'Debounced_LinkHygiene_') {
      const key = 'ATS:queue:link:scheduled';
      const timestamp = dp.getProperty(key);
      // For link hygiene, we track by timestamp, so we can't directly match trigger IDs
      // We'll just check if the trigger was created recently
    }
    
    let cleanedCount = 0;
    const now = new Date().getTime();
    const maxAge = 5 * 60 * 1000; // 5 minutes in milliseconds
    
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === handlerFunction) {
        const triggerId = trigger.getUniqueId();
        
        // For reconcile triggers, check if tracked
        if (handlerFunction === 'Debounced_Reconcile_' && !trackedIds.has(triggerId)) {
          // Check if trigger is old (created more than 5 minutes ago suggests it's stuck)
          const triggerAge = now - trigger.getTriggerSource();
          if (triggerAge > maxAge) {
            ScriptApp.deleteTrigger(trigger);
            cleanedCount++;
            logInfo('Cleaned up orphaned trigger', { 
              function: handlerFunction, 
              triggerId,
              ageMinutes: Math.round(triggerAge / 60000)
            });
          }
        }
      }
    }
    
    if (cleanedCount > 0) {
      logInfo('Orphaned trigger cleanup complete', { 
        function: handlerFunction, 
        cleaned: cleanedCount 
      });
    }
  } catch (e) {
    logWarn('Failed to clean up orphaned triggers', { 
      function: handlerFunction, 
      error: e.message 
    });
  }
}