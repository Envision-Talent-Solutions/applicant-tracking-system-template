/** @file FormInit.gs - Creates and links a Google Form for candidate submission. */

function InitOrRepair_Form() {
  // Check Forms authorization before proceeding
  if (!AuthGuards.forms()) {
    SpreadsheetApp.getUi().alert(
      'Authorization Required',
      'Google Forms authorization is required to create the candidate submission form.\n\n' +
      'Please go to System Admin > Authorization > Authorize Script to grant permissions.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  withLock(() => {
    const ss = SpreadsheetApp.getActive();
    const all = ss.getSheetByName(SHEET_ALL);
    if (!all) {
      SpreadsheetApp.getUi().alert(
        'Setup Failed',
        'Cannot create form: "Candidate Database" sheet not found.\n\n' +
        'Please ensure your spreadsheet has a sheet named "Candidate Database" before setting up the form.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    // --- Create Form & Store ID ---
    let form = FormApp.create('Candidate Submission (Envision ATS)');
    const formId = form.getId();
    StateManager.setFormId(formId);
    
    form.setDescription('Submit your details and resume. Entries will be automatically added to the Envision ATS.');
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

    // --- Build Form Items ---
    form.addTextItem().setTitle(H_ALL.FullName).setRequired(true);
    form.addTextItem().setTitle(H_ALL.Email).setRequired(true);
    form.addTextItem().setTitle(H_ALL.Phone);
    form.addTextItem().setTitle(H_ALL.City);
    form.addTextItem().setTitle(H_ALL.State);

    const req = ss.getSheetByName(SHEET_REQUISITIONS);
    if (req) {
      const reqHeaderInfo = getHeaderInfo(req, ANCHOR_HEADER_REQ);
      if (reqHeaderInfo) {
        const hReq = reqHeaderInfo.headerMap;
        const { rows } = buildReqIndex(req, hReq, reqHeaderInfo.dataStartRow);
        const choices = rows
          .filter(r => OPEN_STATUSES.has((r.data[H_REQ.JobStatus] || '').toString().trim()))
          .map(r => r.data[H_REQ.JobID])
          .filter(Boolean)
          .map(String);
        
        if (choices.length > 0) {
          const dd = form.addListItem().setTitle(H_ALL.JobID).setRequired(true);
          dd.setChoices(choices.map(c => dd.createChoice(c)));
        } else {
          form.addTextItem().setTitle(H_ALL.JobID).setRequired(true).setHelpText('Enter the Job ID you are applying for.');
        }
      } else {
        form.addTextItem().setTitle(H_ALL.JobID).setRequired(true).setHelpText('Enter the Job ID you are applying for.');
      }
    } else {
      form.addTextItem().setTitle(H_ALL.JobID).setRequired(true).setHelpText('Enter the Job ID you are applying for.');
    }

    form.addTextItem().setTitle(H_ALL.Resume).setHelpText('Paste a web link to your resume (Google Drive, Dropbox, URL).');
    form.addTextItem().setTitle(H_ALL.LinkedIn).setHelpText('Paste a link to your LinkedIn Profile.');

    // --- Log & Notify ---
    const formUrl = form.getPublishedUrl();
    logInfo('Candidate submission form created & linked.', { formId: formId, url: formUrl });
    toast(`Form created! The "Form Responses 1" sheet is now linked.`, 7);
    SpreadsheetApp.getUi().alert(`Candidate Form Ready`, `The candidate submission form has been created and linked to this spreadsheet.\n\nShare this link with candidates:\n${formUrl}`, SpreadsheetApp.getUi().ButtonSet.OK);
  });
}