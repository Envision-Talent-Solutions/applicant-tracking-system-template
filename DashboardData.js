/** @file DashboardData.gs - Dashboard data foundation with metrics and historical tracking. */

/**
 * Main setup function - creates dashboard data foundation
 */
function setupDashboardData() {
  const ss = SpreadsheetApp.getActive();
 
  ss.toast('Setting up dashboard data foundation...', 'Please Wait', -1);
  
  // Delete old data sheet if exists
  const oldData = ss.getSheetByName('Dashboard_Data');
  if (oldData) {
    try {
      ss.deleteSheet(oldData);
    } catch (e) {
      logWarn('Could not delete old Dashboard_Data sheet', { error: e.message });
      ss.toast('Please manually delete the old Dashboard_Data sheet and try again.', 'Error', 10);
      return;
    }
  }
  
  // Create new data sheet
  let dataSheet;
  try {
    dataSheet = ss.insertSheet('Dashboard_Data');
  } catch (e) {
    logWarn('Could not create Dashboard_Data sheet', { error: e.message });
    ss.toast('Failed to create Dashboard_Data sheet. You may not have permission to add sheets to this spreadsheet.', 'Error', 10);
    return;
  }
  
  // Get header info with validation
  const reqHeaderInfo = getHeaderInfo(ss.getSheetByName(SHEET_REQUISITIONS), ANCHOR_HEADER_REQ);
  const allHeaderInfo = getHeaderInfo(ss.getSheetByName(SHEET_ALL), ANCHOR_HEADER_ALL);
  const settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  
  if (!reqHeaderInfo || !allHeaderInfo || !settingsSheet) {
    const missing = [];
    if (!reqHeaderInfo) missing.push('Requisitions');
    if (!allHeaderInfo) missing.push('Candidate Database');
    if (!settingsSheet) missing.push('Settings');

    ss.toast(`Setup failed: Missing required sheet(s): ${missing.join(', ')}`, 'Error', 10);
    logWarn('Dashboard setup failed - missing sheets', {
      hasReq: !!reqHeaderInfo,
      hasAll: !!allHeaderInfo,
      hasSettings: !!settingsSheet
    });
    return;
  }
  
  // Populate the data sheet
  try {
    _buildDataSheet(dataSheet, reqHeaderInfo, allHeaderInfo, settingsSheet);
  } catch (e) {
    logWarn('Failed to build data sheet', { error: e.message });
    ss.toast('Dashboard setup failed. Check SYS_LOGS for details.', 'Error', 10);
    return;
  }
  
  // Create named ranges
  try {
    _createNamedRanges(ss, dataSheet);
  } catch (e) {
    logWarn('Failed to create named ranges', { error: e.message });
    ss.toast('Named ranges creation failed, but data sheet is ready.', 'Warning', 8);
  }
  
  // Basic formatting only
  try {
    dataSheet.getRange('A1:Z1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
    dataSheet.hideSheet();
  } catch (e) {
    logWarn('Failed to format/hide sheet', { error: e.message });
  }
  
  ss.toast('✅ Dashboard data ready! All metrics and historical tracking complete.', 'Complete', 8);
  logInfo('Dashboard setup completed successfully');
}

/**
 * Validates sheet names to prevent QUERY injection.
 */
function _validateSheetName(name) {
  if (!name || typeof name !== 'string') {
    throw new Error('Sheet name must be a non-empty string');
  }
  
  if (!/^[a-zA-Z0-9 _-]+$/.test(name)) {
    throw new Error(`Invalid sheet name: "${name}"`);
  }
  
  return name;
}

/**
 * Builds the complete data sheet with all metrics and historical tracking.
 */
function _buildDataSheet(dataSheet, reqHeaderInfo, allHeaderInfo, settingsSheet) {
  const hmReq = reqHeaderInfo.headerMap;
  const hmAll = allHeaderInfo.headerMap;
  
  const safeReqSheetName = _validateSheetName(SHEET_REQUISITIONS);
  const safeAllSheetName = _validateSheetName(SHEET_ALL);
  
  const reqSheetName = `'${safeReqSheetName}'!`;
  const allSheetName = `'${safeAllSheetName}'!`;
  
  // Validate required columns
  if (!hmReq[H_REQ.JobStatus] || !hmReq[H_REQ.DaysOpen] || !hmReq[H_REQ.JobID]) {
    throw new Error('Missing required Requisitions headers');
  }
  if (!hmAll[H_ALL.Stage] || !hmAll[H_ALL.HiredDate] || !hmAll[H_ALL.Created]) {
    throw new Error('Missing required All Candidates headers');
  }
  
  // Build column references
  const reqStatusCol = reqSheetName + _colToLetter(hmReq[H_REQ.JobStatus]) + ":" + _colToLetter(hmReq[H_REQ.JobStatus]);
  const reqDaysOpenCol = reqSheetName + _colToLetter(hmReq[H_REQ.DaysOpen]) + ":" + _colToLetter(hmReq[H_REQ.DaysOpen]);
  const reqJobIdCol = reqSheetName + _colToLetter(hmReq[H_REQ.JobID]) + ":" + _colToLetter(hmReq[H_REQ.JobID]);
  const reqCreatedCol = reqSheetName + _colToLetter(hmReq[H_REQ.Created]) + ":" + _colToLetter(hmReq[H_REQ.Created]);
  const allStageCol = allSheetName + _colToLetter(hmAll[H_ALL.Stage]) + ":" + _colToLetter(hmAll[H_ALL.Stage]);
  const allHiredDateCol = allSheetName + _colToLetter(hmAll[H_ALL.HiredDate]) + ":" + _colToLetter(hmAll[H_ALL.HiredDate]);
  const allCreatedDateCol = allSheetName + _colToLetter(hmAll[H_ALL.Created]) + ":" + _colToLetter(hmAll[H_ALL.Created]);
  const allFullNameCol = allSheetName + _colToLetter(hmAll[H_ALL.FullName]) + ":" + _colToLetter(hmAll[H_ALL.FullName]);
  const allJobIdCol = allSheetName + _colToLetter(hmAll[H_ALL.JobID]) + ":" + _colToLetter(hmAll[H_ALL.JobID]);
  
  // ===== SECTION A: CURRENT METRICS =====
  dataSheet.getRange('A1').setValue('CURRENT METRICS');
  const currentMetrics = [
    ['Metric Name', 'Value', 'Named Range', 'Description'],
    ['Open Reqs', `=COUNTIF(${reqStatusCol}, "Open") + COUNTIF(${reqStatusCol}, "On Hold")`, 'OpenReqs', 'Total open requisitions including on hold'],
    ['Avg Days Open', `=IFERROR(ROUND(AVERAGEIFS(${reqDaysOpenCol}, ${reqStatusCol}, "Open")), 0)`, 'AvgDaysOpen', 'Average days requisitions have been open'],
    ['Hires This Month', `=COUNTIFS(${allHiredDateCol}, ">="&EOMONTH(TODAY(),-1)+1, ${allHiredDateCol}, "<="&EOMONTH(TODAY(),0))`, 'HiresThisMonth', 'Hires in current calendar month'],
    ['Active Candidates', `=COUNTIFS(${allStageCol}, "<>New Applicant", ${allStageCol}, "<>Rejected", ${allStageCol}, "<>Hired", ${allStageCol}, "<>")`, 'ActiveCandidates', 'Candidates actively in pipeline'],
    ['Total Hires', `=COUNTIF(${allStageCol}, "Hired")`, 'TotalHires', 'All-time hires'],
    ['Total Candidate Profiles', `=IFERROR(COUNTA(${allFullNameCol})-1, 0)`, 'TotalCandidates', 'Total candidate profiles ever'],
    ['In Interview', `=COUNTIF(${allStageCol}, "*Interview*")`, 'InInterview', 'Candidates in any interview stage'],
    ['Hires This Week', `=COUNTIFS(${allHiredDateCol}, ">="&TODAY()-WEEKDAY(TODAY(),2)+1)`, 'HiresThisWeek', 'Hires in current week'],
    ['Hires This Year', `=COUNTIFS(${allHiredDateCol}, ">="&DATE(YEAR(TODAY()),1,1))`, 'HiresThisYear', 'Hires in current year'],
    ['Reqs Over 30 Days', `=COUNTIFS(${reqDaysOpenCol}, ">30", ${reqStatusCol}, "Open")`, 'ReqsOver30Days', 'Open reqs older than 30 days'],
    ['Reqs No Candidates', `=IFERROR(ROWS(FILTER(${reqJobIdCol}, SEARCH("Open", ${reqStatusCol}), ISNA(MATCH(${reqJobIdCol}, ${allJobIdCol}, 0)))), 0)`, 'ReqsNoCandidates', 'Open reqs with zero applicants'],
    ['New Applicants', `=COUNTIF(${allStageCol}, "New Applicant")`, 'NewApplicants', 'Applicants in "New Applicant" stage'],
    ['Rejected Total', `=COUNTIF(${allStageCol}, "Rejected")`, 'RejectedTotal', 'Total rejected candidates'],
    ['Time to Fill Avg', `=IFERROR(ROUND(AVERAGE(FILTER(${reqDaysOpenCol}, ${reqStatusCol}="Hired")), 0), 0)`, 'TimeToFill', 'Average days to fill a position'],
    ['Filled Reqs', `=COUNTIF(${reqStatusCol}, "Hired")`, 'FilledReqs', 'Total filled requisitions'],
    ['On Hold Reqs', `=COUNTIF(${reqStatusCol}, "On Hold")`, 'OnHoldReqs', 'Requisitions on hold'],
    ['Profiles Last 7 Days', `=COUNTIFS(${allCreatedDateCol}, ">="&TODAY()-7, ${allCreatedDateCol}, "<="&TODAY())`, 'ProfilesLast7Days', 'Candidate profiles added in last 7 days'],
    ['Profiles Last 30 Days', `=COUNTIFS(${allCreatedDateCol}, ">="&TODAY()-30, ${allCreatedDateCol}, "<="&TODAY())`, 'ProfilesLast30Days', 'Candidate profiles added in last 30 days'],
    ['Avg Candidates Per Req', `=IFERROR(ROUND(B7/B2, 1), 0)`, 'AvgCandidatesPerReq', 'Average candidates per requisition'],
  ];
  dataSheet.getRange(1, 1, currentMetrics.length, 4).setValues(currentMetrics);
  
  // ===== SECTION B: HISTORICAL COMPARISON (EXPANDED) =====
  const histStart = currentMetrics.length + 2;
  dataSheet.getRange(histStart, 1).setValue('HISTORICAL COMPARISONS');
  const historicalMetrics = [
    ['Metric Name', 'Value', 'Named Range', 'Description'],
    ['Open Reqs Last Month', `=COUNTIFS(${reqCreatedCol}, ">="&EOMONTH(TODAY(),-2)+1, ${reqCreatedCol}, "<="&EOMONTH(TODAY(),-1))`, 'OpenReqsLastMonth', 'Open reqs from last month'],
    ['Change in Open Reqs', `=B2-B${histStart+1}`, 'OpenReqsChange', 'Change from last month'],
    ['Active Candidates Last Week', `=COUNTIFS(${allCreatedDateCol}, ">="&TODAY()-14, ${allCreatedDateCol}, "<"&TODAY()-7)`, 'ActiveCandidatesLastWeek', 'Active candidates from last week'],
    ['Change in Active Candidates', `=B5-B${histStart+3}`, 'ActiveCandidatesChange', 'Change from last week'],
    ['Hires Last Month', `=COUNTIFS(${allHiredDateCol}, ">="&EOMONTH(TODAY(),-2)+1, ${allHiredDateCol}, "<="&EOMONTH(TODAY(),-1))`, 'HiresLastMonth', 'Hires from last month'],
    ['Change in Hires', `=B4-B${histStart+5}`, 'HiresChange', 'Change from last month'],
    ['Avg Days Open Last Month', `=IFERROR(ROUND(AVERAGEIFS(${reqDaysOpenCol}, ${reqCreatedCol}, ">="&EOMONTH(TODAY(),-2)+1, ${reqCreatedCol}, "<="&EOMONTH(TODAY(),-1))), 0)`, 'AvgDaysOpenLastMonth', 'Avg days open last month'],
    ['Change in Avg Days', `=B3-B${histStart+7}`, 'AvgDaysChange', 'Change from last month'],
    ['In Interview Last Week', `=COUNTIFS(${allStageCol}, "*Interview*", ${allCreatedDateCol}, ">="&TODAY()-14, ${allCreatedDateCol}, "<"&TODAY()-7)`, 'InInterviewLastWeek', 'In interview last week'],
    ['Change in Interview', `=B8-B${histStart+9}`, 'InInterviewChange', 'Change from last week'],
    ['New Applicants Last Week', `=COUNTIFS(${allStageCol}, "New Applicant", ${allCreatedDateCol}, ">="&TODAY()-7, ${allCreatedDateCol}, "<"&TODAY())`, 'NewApplicantsLastWeek', 'New applicants last week'],
    ['Change in New Applicants', `=B13-B${histStart+11}`, 'NewApplicantsChange', 'Change from last week'],
    ['Profiles Added Last Week', `=COUNTIFS(${allCreatedDateCol}, ">="&TODAY()-14, ${allCreatedDateCol}, "<"&TODAY()-7)`, 'ProfilesAddedLastWeek', 'Profiles added last week'],
    ['Change in Profiles', `=B18-B${histStart+13}`, 'ProfilesChange', 'Change from last week'],
  ];
  dataSheet.getRange(histStart, 1, historicalMetrics.length, 4).setValues(historicalMetrics);
  
  // ===== SECTION C: SPARKLINE DATA (EXPANDED) =====
  const sparkStart = histStart + historicalMetrics.length + 2;
  dataSheet.getRange(sparkStart, 1).setValue('SPARKLINE DATA');
  const sparklineData = [
    ['Metric', 'Last Period', 'Current', 'For Sparkline Formula'],
    ['Open Reqs Trend', `=B${histStart+1}`, '=B2', '=SPARKLINE({B' + (sparkStart+1) + ':C' + (sparkStart+1) + '}, {"charttype","column"})'],
    ['Active Candidates Trend', `=B${histStart+3}`, '=B5', '=SPARKLINE({B' + (sparkStart+2) + ':C' + (sparkStart+2) + '}, {"charttype","column"})'],
    ['Hires Trend', `=B${histStart+5}`, '=B4', '=SPARKLINE({B' + (sparkStart+3) + ':C' + (sparkStart+3) + '}, {"charttype","column"})'],
    ['Days Open Trend', `=B${histStart+7}`, '=B3', '=SPARKLINE({B' + (sparkStart+4) + ':C' + (sparkStart+4) + '}, {"charttype","column"})'],
    ['Interview Trend', `=B${histStart+9}`, '=B8', '=SPARKLINE({B' + (sparkStart+5) + ':C' + (sparkStart+5) + '}, {"charttype","column"})'],
    ['New Applicants Trend', `=B${histStart+11}`, '=B13', '=SPARKLINE({B' + (sparkStart+6) + ':C' + (sparkStart+6) + '}, {"charttype","column"})'],
    ['Profiles Trend', `=B${histStart+13}`, '=B18', '=SPARKLINE({B' + (sparkStart+7) + ':C' + (sparkStart+7) + '}, {"charttype","column"})'],
  ];
  dataSheet.getRange(sparkStart, 1, sparklineData.length, 4).setValues(sparklineData);
  
  // ===== SECTION D: CONVERSION METRICS =====
  const conversionStart = sparkStart + sparklineData.length + 2;
  dataSheet.getRange(conversionStart, 1).setValue('CONVERSION METRICS');
  const conversionMetrics = [
    ['Metric Name', 'Value', 'Named Range', 'Description'],
    ['Overall Conversion %', `=IFERROR(B6/B7, 0)`, 'ConversionOverall', 'Total Hires / Total Candidate Profiles'],
    ['Interview to Hire %', `=IFERROR(B6/B8, 0)`, 'ConversionInterview', 'Total Hires / In Interview'],
    ['Pipeline Fill Rate', `=IFERROR(B5/B2, 0)`, 'PipelineFillRate', 'Active Candidates / Open Reqs'],
    ['Screening Rate %', `=IFERROR((B7-B13)/B7, 0)`, 'ScreeningRate', '(Profiles - New) / Total Profiles'],
    ['Rejection Rate %', `=IFERROR(B14/B7, 0)`, 'RejectionRate', 'Rejected / Total Candidate Profiles'],
    ['Fill Rate %', `=IFERROR(B16/(B16+B2), 0)`, 'FillRate', 'Filled / (Filled + Open)'],
    ['Time to Hire (Days)', `=B15`, 'TimeToHire', 'Same as Time to Fill Avg'],
  ];
  dataSheet.getRange(conversionStart, 1, conversionMetrics.length, 4).setValues(conversionMetrics);
  dataSheet.getRange(conversionStart + 1, 2, conversionMetrics.length - 1, 1).setNumberFormat('0.0%');
  
  // ===== SECTION E: CANDIDATE FUNNEL =====
  const funnelStart = conversionStart + conversionMetrics.length + 2;
  dataSheet.getRange(funnelStart, 1).setValue('CANDIDATE FUNNEL');
  dataSheet.getRange(funnelStart + 1, 1, 1, 2).setValues([['Stage', 'Count']]);
  
  let funnelStages = [];
  let stageColFound = false;
  
  try {
    const headerWithBreak = "Candidate\nWorkflow Status";
    const settingsVals = settingsSheet.getDataRange().getValues();
    let stageHeaderIdx = settingsVals[0].indexOf(headerWithBreak);
    
    if (stageHeaderIdx === -1) {
      stageHeaderIdx = settingsVals[0].indexOf('Candidate Workflow Status');
    }
    
    if (stageHeaderIdx === -1) {
      const funnelHeader = Object.keys(allHeaderInfo.headerMap).find(h => h.trim() === H_ALL.Stage.trim());
      stageHeaderIdx = settingsVals[0].indexOf(funnelHeader);
    }
    
    if (stageHeaderIdx !== -1) {
      funnelStages = settingsVals.slice(1).map(row => row[stageHeaderIdx]).filter(String);
      if (funnelStages.length > 0) {
        stageColFound = true;
      }
    }
  } catch (e) {
    logWarn('Error reading funnel stages from Settings', { error: e.message });
  }
  
  if (stageColFound && funnelStages.length > 0) {
    const funnelFormulas = funnelStages.map(stage => [stage, `=COUNTIF(${allStageCol}, "${stage}")`]);
    dataSheet.getRange(funnelStart + 2, 1, funnelFormulas.length, 2).setValues(funnelFormulas);
  } else {
    dataSheet.getRange(funnelStart + 2, 1).setFormula(
      `=IFERROR(SORT(UNIQUE(FILTER(${allStageCol}, ${allStageCol}<>"")), 1, TRUE), "No stages found")`
    );
    dataSheet.getRange(funnelStart + 2, 2).setFormula(
      `=IF(A${funnelStart + 2}="No stages found", "", COUNTIF(${allStageCol}, A${funnelStart + 2}))`
    );
    for (let i = 1; i < 15; i++) {
      dataSheet.getRange(funnelStart + 2 + i, 2).setFormula(
        `=IF(A${funnelStart + 2 + i}="", "", COUNTIF(${allStageCol}, A${funnelStart + 2 + i}))`
      );
    }
  }
  
  // ===== SECTION F: HIRES BY MONTH =====
  const hiresStart = funnelStart;
  dataSheet.getRange(hiresStart, 4).setValue('HIRES BY MONTH (Last 12)');
  dataSheet.getRange(hiresStart + 1, 4, 1, 2).setValues([['Month', 'Hires']]);
  
  const hiresByMonth = Array.from({length: 12}).map((_, i) => {
    const monthFormula = `=TEXT(EOMONTH(TODAY(), -${i}), "MMM YYYY")`;
    const countFormula = `=COUNTIFS(${allHiredDateCol}, ">="&EOMONTH(TODAY(),-${i}-1)+1, ${allHiredDateCol}, "<="&EOMONTH(TODAY(),-${i}))`;
    return [monthFormula, countFormula];
  });
  dataSheet.getRange(hiresStart + 2, 4, 12, 2).setFormulas(hiresByMonth);
  
  // ===== SECTION G: CANDIDATE PROFILES BY WEEK =====
  const profilesWeekStart = hiresStart + 15;
  dataSheet.getRange(profilesWeekStart, 4).setValue('CANDIDATE PROFILES BY WEEK (Last 12)');
  dataSheet.getRange(profilesWeekStart + 1, 4, 1, 2).setValues([['Week Ending', 'New Profiles']]);
  
  const profilesByWeek = Array.from({length: 12}).map((_, i) => {
    const countFormula = `=IFERROR(COUNTIFS(${allCreatedDateCol}, ">="&TODAY()-${(i+1)*7}+1, ${allCreatedDateCol}, "<="&TODAY()-${i*7}), 0)`;
    const weekLabel = `=TEXT(TODAY()-${i*7}, "MMM DD")`;
    return [weekLabel, countFormula];
  });
  dataSheet.getRange(profilesWeekStart + 2, 4, 12, 2).setFormulas(profilesByWeek);
  
  // ===== SECTION H: OLDEST OPEN REQS =====
  const oldestStart = hiresStart;
  dataSheet.getRange(oldestStart, 7).setValue('OLDEST OPEN REQUISITIONS');
  dataSheet.getRange(oldestStart + 1, 7, 1, 3).setValues([['Job Title', 'Job ID', 'Days Open']]);
  
  const oldestFormula = `=IFERROR(SORTN(QUERY('${safeReqSheetName}'!A:Z, "SELECT B, C, R WHERE (A='Open' OR A='On Hold') and B is not null", 0), 15, 0, 3, FALSE), {"No data", "", ""})`;
  dataSheet.getRange(oldestStart + 2, 7).setFormula(oldestFormula);
  
  // ===== DOCUMENTATION SECTION =====
  _addDocumentation(dataSheet);
  
  logInfo('Dashboard data sheet built successfully');
}

/**
 * Adds comprehensive documentation to the data sheet.
 */
function _addDocumentation(dataSheet) {
  const docStart = 2;
  const docCol = 12;
  
  dataSheet.getRange(docStart, docCol, 1, 3).merge()
    .setValue('📚 DASHBOARD DATA REFERENCE GUIDE')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  dataSheet.getRange(docStart + 2, docCol).setValue('TREND TEXT FORMULAS')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground('#e8f0fe');
  
  dataSheet.getRange(docStart + 3, docCol, 1, 2).setValues([['Metric', 'Formula']]);
  
  const trendFormulas = [
    ['Open Reqs', '=IF(OpenReqsChange>0,"▲ ","▼ ")&ABS(OpenReqsChange)&" from last month"'],
    ['Active Candidates', '=IF(ActiveCandidatesChange>0,"▲ ","▼ ")&ABS(ActiveCandidatesChange)&" from last week"'],
    ['Hires', '=IF(HiresChange>0,"▲ ","▼ ")&ABS(HiresChange)&" from last month"'],
    ['Avg Days Open', '=IF(AvgDaysChange>0,"▲ ","▼ ")&ABS(AvgDaysChange)&" days vs last month"'],
    ['In Interview', '=IF(InInterviewChange>0,"▲ ","▼ ")&ABS(InInterviewChange)&" from last week"'],
    ['New Applicants', '=IF(NewApplicantsChange>0,"▲ ","▼ ")&ABS(NewApplicantsChange)&" from last week"'],
    ['Profiles', '=IF(ProfilesChange>0,"▲ ","▼ ")&ABS(ProfilesChange)&" from last week"'],
  ];
  
  dataSheet.getRange(docStart + 4, docCol, trendFormulas.length, 2).setValues(trendFormulas);
  
  const condStart = docStart + 4 + trendFormulas.length + 2;
  dataSheet.getRange(condStart, docCol).setValue('CONDITIONAL FORMATTING')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground('#e8f0fe');
  
  dataSheet.getRange(condStart + 1, docCol, 1, 2).setValues([['Green (Good)', 'Red (Bad)']]);
  
  const condFormulas = [
    ['=INDIRECT("OpenReqsChange")>0', '=INDIRECT("OpenReqsChange")<0'],
    ['=INDIRECT("HiresChange")>0', '=INDIRECT("HiresChange")<0'],
    ['=INDIRECT("AvgDaysChange")<0', '=INDIRECT("AvgDaysChange")>0'],
    ['=INDIRECT("InInterviewChange")>0', '=INDIRECT("InInterviewChange")<0'],
  ];
  
  dataSheet.getRange(condStart + 2, docCol, condFormulas.length, 2).setValues(condFormulas);
}

/**
 * Creates all named ranges for dashboard formulas.
 */
function _createNamedRanges(ss, dataSheet) {
  const namedRanges = [
    // Current metrics
    {name: 'OpenReqs', cell: 'B2'},
    {name: 'AvgDaysOpen', cell: 'B3'},
    {name: 'HiresThisMonth', cell: 'B4'},
    {name: 'ActiveCandidates', cell: 'B5'},
    {name: 'TotalHires', cell: 'B6'},
    {name: 'TotalCandidates', cell: 'B7'},
    {name: 'InInterview', cell: 'B8'},
    {name: 'HiresThisWeek', cell: 'B9'},
    {name: 'HiresThisYear', cell: 'B10'},
    {name: 'ReqsOver30Days', cell: 'B11'},
    {name: 'ReqsNoCandidates', cell: 'B12'},
    {name: 'NewApplicants', cell: 'B13'},
    {name: 'RejectedTotal', cell: 'B14'},
    {name: 'TimeToFill', cell: 'B15'},
    {name: 'FilledReqs', cell: 'B16'},
    {name: 'OnHoldReqs', cell: 'B17'},
    {name: 'ProfilesLast7Days', cell: 'B18'},
    {name: 'ProfilesLast30Days', cell: 'B19'},
    {name: 'AvgCandidatesPerReq', cell: 'B20'},
    // Historical comparisons
    {name: 'OpenReqsLastMonth', cell: 'B23'},
    {name: 'OpenReqsChange', cell: 'B24'},
    {name: 'ActiveCandidatesLastWeek', cell: 'B25'},
    {name: 'ActiveCandidatesChange', cell: 'B26'},
    {name: 'HiresLastMonth', cell: 'B27'},
    {name: 'HiresChange', cell: 'B28'},
    {name: 'AvgDaysOpenLastMonth', cell: 'B29'},
    {name: 'AvgDaysChange', cell: 'B30'},
    {name: 'InInterviewLastWeek', cell: 'B31'},
    {name: 'InInterviewChange', cell: 'B32'},
    {name: 'NewApplicantsLastWeek', cell: 'B33'},
    {name: 'NewApplicantsChange', cell: 'B34'},
    {name: 'ProfilesAddedLastWeek', cell: 'B35'},
    {name: 'ProfilesChange', cell: 'B36'},
    // Conversion metrics (row 48+)
    {name: 'ConversionOverall', cell: 'B48'},
    {name: 'ConversionInterview', cell: 'B49'},
    {name: 'PipelineFillRate', cell: 'B50'},
    {name: 'ScreeningRate', cell: 'B51'},
    {name: 'RejectionRate', cell: 'B52'},
    {name: 'FillRate', cell: 'B53'},
    {name: 'TimeToHire', cell: 'B54'},
  ];
  
  let successCount = 0;
  let failCount = 0;
  
  namedRanges.forEach(range => {
    try {
      const existing = ss.getNamedRanges().find(nr => nr.getName() === range.name);
      if (existing) existing.remove();
      
      ss.setNamedRange(range.name, dataSheet.getRange(range.cell));
      successCount++;
    } catch (e) {
      logWarn('Error creating named range', { name: range.name, error: e.message });
      failCount++;
    }
  });
  
  logInfo('Named ranges created', { success: successCount, failed: failCount });
}

/**
 * Converts column number to letter
 */
function _colToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}