/**
 * @OnlyCurrentDoc
 *
 * The above comment specifies that this automation will only
 * attempt to read or modify the spreadsheet this script is bound to.
 * The authorization request message presented to users reflects the
 * limited scope.
 */

/**
 * Creates a custom menu in the Google Sheets UI when the document is opened.
 *
 * @param {object} e The event parameter for a simple onOpen trigger.
 */


function onOpen(e) {
  console.info('onOpen', 'e.authMode', e && e.authMode)
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Settings', 'showSidebar')
      .addItem('Install trigger on form submission', 'installTrigger')
      .addItem('Delete triggers', 'deleteTriggers')
      .addToUi();
}

/**
 * Runs when the add-on is installed; calls onOpen() to ensure menu creation and
 * any other initializion work is done immediately.
 *
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e) {
  console.info('onInstall', 'e.authMode', e && e.authMode)
  onOpen(e);
}

/**
 * Opens a sidebar. The sidebar structure is described in the Sidebar.html
 * project file.
 */
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('SideBar')
      .evaluate()
      .setTitle(APP_TITLE)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(ui);
}

function installTrigger() {
  ScriptApp.newTrigger('onFormSubmitFunction')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onFormSubmit()
      .create();
  SpreadsheetApp.getUi().alert("Trigger installed");
}

function deleteTriggers(){
  // Deletes all triggers in the current project.
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  SpreadsheetApp.getUi().alert("Triggers deleted");
}
