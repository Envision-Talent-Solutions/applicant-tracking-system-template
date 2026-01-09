/**
 * @file ResumeLinker.gs - Processes uploaded resumes and creates/links candidate records.
 * Resume text is extracted client-side in the browser before being sent to this module.
 */

const EMAIL_REGEX = /([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/gi;
const PHONE_REGEX = /(\+?\d{1,2}[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}/g;
const LINKEDIN_REGEX = /(?:https?:\/\/)?(?:www\.)?linkedin\.com\/in\/[a-zA-Z0-9_-]+\/?/gi;

/**
 * Processes uploaded resume files and creates/links candidates.
 * Called from the ImportSidebar after client-side text extraction.
 *
 * @param {Array<{filename: string, text: string, driveUrl?: string}>} resumeData - Array of resume objects
 * @returns {Object} Result object with linked, created, failed, and total counts
 */
function processUploadedResumes(resumeData) {
  if (!Array.isArray(resumeData) || resumeData.length === 0) {
    throw new Error('No resume data provided. Please select files to upload.');
  }

  return withLock(() => {
    const ss = SpreadsheetApp.getActive();
    const allSheet = ss.getSheetByName(SHEET_ALL);
    if (!allSheet) {
      throw new Error(
        'Cannot find the "Candidate Database" sheet.\n\n' +
        'Please ensure you have a sheet named "Candidate Database" in this spreadsheet.'
      );
    }

    const headerInfo = getHeaderInfo(allSheet, ANCHOR_HEADER_ALL);
    if (!headerInfo) {
      throw new Error(
        'Cannot find the header row in "Candidate Database" sheet.\n\n' +
        'Please ensure the sheet has a header row containing "Full Name" in column A.'
      );
    }
    const hm = headerInfo.headerMap;

    // Build index of existing candidates by email and phone
    const { rows: candRows } = buildCandidateIndex(allSheet, hm, headerInfo.dataStartRow);
    const emailMap = new Map();
    const phoneMap = new Map();

    for (const { row, data } of candRows) {
      const email = normEmail(data[H_ALL.Email]);
      const phone = normPhone(data[H_ALL.Phone]);
      if (email) emailMap.set(email, row);
      if (phone) phoneMap.set(phone, row);
    }

    // Capture template formatting before inserting rows
    const template = captureTemplateFormat(allSheet, headerInfo.dataStartRow);

    // Track results
    let linkedCount = 0;
    let createdCount = 0;
    const failedFiles = []; // Track which files failed and why
    const jobIdsToReconcile = new Set();

    // Process each resume
    for (const resume of resumeData) {
      try {
        const { filename, text, driveUrl } = resume;

        if (!text || text.trim() === '') {
          failedFiles.push({ filename, reason: 'Could not extract text from file' });
          continue;
        }

        // Extract contact info from resume text
        const emails = (text.match(EMAIL_REGEX) || []).map(normEmail).filter(Boolean);
        const phones = (text.match(PHONE_REGEX) || []).map(normPhone).filter(Boolean);
        const linkedinUrls = (text.match(LINKEDIN_REGEX) || []).map(url => {
          // Normalize LinkedIn URLs to full https format
          let normalized = url.trim();
          if (!normalized.startsWith('http')) {
            normalized = 'https://' + normalized;
          }
          return normalized;
        });

        // Try to find existing candidate by email first, then phone
        let existingRow = null;
        for (const email of emails) {
          if (emailMap.has(email)) {
            existingRow = emailMap.get(email);
            break;
          }
        }
        if (!existingRow) {
          for (const phone of phones) {
            if (phoneMap.has(phone)) {
              existingRow = phoneMap.get(phone);
              break;
            }
          }
        }

        if (existingRow) {
          // Match found - update existing candidate's timestamp
          safeSetValue(allSheet, H_ALL.Updated, allSheet.getRange(existingRow, hm[H_ALL.Updated]), nowDetroit());
          linkedCount++;
        } else if (emails.length > 0 || phones.length > 0) {
          // No match - create new candidate
          const newRowObj = {};
          newRowObj[H_ALL.FullName] = _cleanFileNameForName_(filename);
          newRowObj[H_ALL.Email] = emails[0] || '';
          newRowObj[H_ALL.Phone] = phones[0] || '';
          newRowObj[H_ALL.LinkedIn] = linkedinUrls[0] || '';
          newRowObj[H_ALL.Resume] = driveUrl || '';  // Store Drive URL if available
          newRowObj[H_ALL.Source] = 'Resume Import';
          newRowObj[H_ALL.Created] = nowDetroit();
          newRowObj[H_ALL.Updated] = nowDetroit();

          const currentLastRow = Math.max(allSheet.getLastRow(), headerInfo.headerRow);
          const newRowIdx = currentLastRow + 1;
          allSheet.insertRowAfter(currentLastRow);

          applyTemplateFormat(allSheet, template, newRowIdx);

          setRowValuesByHeaders(allSheet, hm, newRowIdx, newRowObj);

          // Set email as clickable mailto: link
          if (hm[H_ALL.Email] && emails[0]) {
            setHyperlink(allSheet, newRowIdx, hm[H_ALL.Email], 'mailto:' + emails[0], emails[0]);
          }

          // Set phone as clickable tel: link
          if (hm[H_ALL.Phone] && phones[0] && phones[0].length >= PHONE_MIN_DIGITS) {
            setHyperlink(allSheet, newRowIdx, hm[H_ALL.Phone], 'tel:' + phones[0], newRowObj[H_ALL.Phone]);
          }

          // Set LinkedIn as clickable link
          if (hm[H_ALL.LinkedIn] && linkedinUrls[0]) {
            setHyperlink(allSheet, newRowIdx, hm[H_ALL.LinkedIn], linkedinUrls[0], 'LinkedIn Profile');
          }

          // Set Resume as clickable link if Drive URL is available
          if (hm[H_ALL.Resume] && driveUrl) {
            setHyperlink(allSheet, newRowIdx, hm[H_ALL.Resume], driveUrl, 'Resume');
          }

          // Apply data validations to the new row
          try {
            _applyValidationsToNewRow_(allSheet, headerInfo, newRowIdx);
          } catch (e) {
            logWarn('Failed to apply validations to new resume row', { error: e.message });
          }

          // Update maps for deduplication within this batch
          if (emails[0]) emailMap.set(emails[0], newRowIdx);
          if (phones[0]) phoneMap.set(phones[0], newRowIdx);

          // Collect job ID for batch reconciliation
          if (newRowObj[H_ALL.JobID]) {
            jobIdsToReconcile.add(String(newRowObj[H_ALL.JobID]).trim());
          }

          createdCount++;
        } else {
          // No contact info found in resume
          failedFiles.push({ filename, reason: 'No email or phone number found in resume' });
        }
      } catch (err) {
        logWarn('Failed to process resume', { filename: resume.filename, error: err.message });
        failedFiles.push({ filename: resume.filename, reason: 'Processing error: ' + err.message });
      }
    }

    // Batch reconciliation for all created candidates
    if (jobIdsToReconcile.size > 0) {
      reconcileActiveMembership_ByJobIds_(Array.from(jobIdsToReconcile));
    }

    // Show toast notification in the spreadsheet
    const failedCount = failedFiles.length;
    const summary = `Resume upload complete: ${linkedCount} linked, ${createdCount} created, ${failedCount} failed`;
    toast(summary);

    return {
      linked: linkedCount,
      created: createdCount,
      failed: failedCount,
      failedFiles: failedFiles,
      total: resumeData.length
    };
  });
}

/**
 * Applies data validations from SETTINGS sheet to a newly created row.
 * @param {Sheet} sheet - The sheet to apply validations to
 * @param {Object} headerInfo - Header information object
 * @param {number} row - The row number to apply validations to
 */
function _applyValidationsToNewRow_(sheet, headerInfo, row) {
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

        const cell = sheet.getRange(row, col);
        safeSetDataValidation(cell, rule, { header });
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

      const cell = sheet.getRange(row, hm[H_ALL.JobID]);
      safeSetDataValidation(cell, rule, { header: H_ALL.JobID });
    } else {
      logInfo('No active Job IDs available - skipping Job ID validation for resume import');
    }
  }
}

/**
 * Processes pasted resume URLs and creates candidate records.
 * Called from ImportSidebar when users paste resume links directly.
 * Since we don't have file content, candidates are created with minimal info.
 *
 * @param {Array<{filename: string, text: string, driveUrl: string}>} resumeData - Array of URL data objects
 * @returns {Object} Result object with linked, created, failed, and total counts
 */
function processResumeLinks(resumeData) {
  if (!Array.isArray(resumeData) || resumeData.length === 0) {
    throw new Error('No URLs provided. Please paste at least one resume link.');
  }

  return withLock(() => {
    const ss = SpreadsheetApp.getActive();
    const allSheet = ss.getSheetByName(SHEET_ALL);
    if (!allSheet) {
      throw new Error(
        'Cannot find the "Candidate Database" sheet.\n\n' +
        'Please ensure you have a sheet named "Candidate Database" in this spreadsheet.'
      );
    }

    const headerInfo = getHeaderInfo(allSheet, ANCHOR_HEADER_ALL);
    if (!headerInfo) {
      throw new Error(
        'Cannot find the header row in "Candidate Database" sheet.\n\n' +
        'Please ensure the sheet has a header row containing "Full Name" in column A.'
      );
    }
    const hm = headerInfo.headerMap;

    // Build index of existing candidates by resume URL to avoid duplicates
    const { rows: candRows } = buildCandidateIndex(allSheet, hm, headerInfo.dataStartRow);
    const urlMap = new Map();

    for (const { row, data } of candRows) {
      const resumeUrl = data[H_ALL.Resume];
      if (resumeUrl) {
        // Normalize URL for comparison
        const normalizedUrl = resumeUrl.toLowerCase().trim();
        urlMap.set(normalizedUrl, row);
      }
    }

    // Capture template formatting before inserting rows
    const template = captureTemplateFormat(allSheet, headerInfo.dataStartRow);

    // Track results
    let linkedCount = 0;
    let createdCount = 0;
    const failedFiles = [];

    // Process each URL
    for (const resume of resumeData) {
      try {
        const { filename, driveUrl } = resume;

        if (!driveUrl || !driveUrl.trim()) {
          failedFiles.push({ filename: filename || 'Unknown', reason: 'No URL provided' });
          continue;
        }

        const normalizedUrl = driveUrl.toLowerCase().trim();

        // Check if this URL already exists in the database
        if (urlMap.has(normalizedUrl)) {
          // URL already exists - skip to avoid duplicates
          linkedCount++;
          continue;
        }

        // Create new candidate with minimal info
        const newRowObj = {};
        newRowObj[H_ALL.FullName] = _cleanFileNameForName_(filename) || 'Unknown Candidate';
        newRowObj[H_ALL.Email] = '';  // No email available from URL only
        newRowObj[H_ALL.Phone] = '';  // No phone available from URL only
        newRowObj[H_ALL.Resume] = driveUrl;
        newRowObj[H_ALL.Source] = 'Resume Link Import';
        newRowObj[H_ALL.Created] = nowDetroit();
        newRowObj[H_ALL.Updated] = nowDetroit();

        const currentLastRow = Math.max(allSheet.getLastRow(), headerInfo.headerRow);
        const newRowIdx = currentLastRow + 1;
        allSheet.insertRowAfter(currentLastRow);

        applyTemplateFormat(allSheet, template, newRowIdx);

        setRowValuesByHeaders(allSheet, hm, newRowIdx, newRowObj);

        // Set Resume as clickable link
        if (hm[H_ALL.Resume] && driveUrl) {
          setHyperlink(allSheet, newRowIdx, hm[H_ALL.Resume], driveUrl, 'Resume');
        }

        // Apply data validations to the new row
        try {
          _applyValidationsToNewRow_(allSheet, headerInfo, newRowIdx);
        } catch (e) {
          logWarn('Failed to apply validations to new resume link row', { error: e.message });
        }

        // Update URL map for deduplication within this batch
        urlMap.set(normalizedUrl, newRowIdx);

        createdCount++;
      } catch (err) {
        logWarn('Failed to process resume link', { filename: resume.filename, error: err.message });
        failedFiles.push({ filename: resume.filename || 'Unknown', reason: 'Processing error: ' + err.message });
      }
    }

    // Show toast notification in the spreadsheet
    const failedCount = failedFiles.length;
    const summary = `Resume links processed: ${linkedCount} already existed, ${createdCount} created, ${failedCount} failed`;
    toast(summary);

    return {
      linked: linkedCount,
      created: createdCount,
      failed: failedCount,
      failedFiles: failedFiles,
      total: resumeData.length
    };
  });
}

/**
 * Cleans a filename for use as a candidate's full name.
 * Removes file extensions and common noise words.
 * @param {string} fileName - The original filename
 * @returns {string} Cleaned name suitable for display
 */
function _cleanFileNameForName_(fileName) {
  if (!fileName) return 'Unknown Candidate';

  // Remove file extension
  let name = fileName.replace(/\.(pdf|docx?|txt)$/i, '');

  // Check for "Name - Resume" or "Resume - Name" patterns
  const lastHyphenIndex = name.lastIndexOf(' - ');
  if (lastHyphenIndex !== -1) {
    const afterHyphen = name.substring(lastHyphenIndex + 3).trim();
    const beforeHyphen = name.substring(0, lastHyphenIndex).trim();

    // If the part after the hyphen looks like a name (not "Resume" or "CV")
    if (afterHyphen.toLowerCase() !== 'resume' && afterHyphen.toLowerCase() !== 'cv') {
      return afterHyphen;
    }
    // If the part before looks like a name
    if (beforeHyphen.toLowerCase() !== 'resume' && beforeHyphen.toLowerCase() !== 'cv') {
      return beforeHyphen;
    }
  }

  const firstHyphenIndex = name.indexOf(' - ');
  if (firstHyphenIndex !== -1) {
    const potentialName = name.substring(0, firstHyphenIndex).trim();
    if (potentialName.toLowerCase() !== 'resume' && potentialName.toLowerCase() !== 'cv') {
      return potentialName;
    }
  }

  // Remove common noise words (including job titles/functions often appended to names)
  const noiseWords = /\b(resume|cv|candidate|application|precision|final|draft|updated|v\d+|recops|recruiting|recruiter|hr|talent|sourcer|sourcing|manager|director|coordinator|specialist|analyst|consultant|engineer|developer|admin|executive|assistant|associate|senior|junior|lead|intern)\b/gi;
  name = name.replace(/_/g, ' ').replace(noiseWords, '');

  // Clean up whitespace
  name = name.replace(/\s+/g, ' ').trim();

  return name || 'Unknown Candidate';
}
