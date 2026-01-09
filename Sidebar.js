/** @file Sidebar.gs - Functions to manage and display UI sidebars. */

/**
 * Creates an HTML output from a file and displays it as a sidebar in the spreadsheet.
 */
function showImportSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('ImportSidebar.html')
      .setTitle('Importing Toolkit')
      .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}