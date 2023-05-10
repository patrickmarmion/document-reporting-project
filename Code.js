const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs");
const statusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Document_status");
const errorsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Errors");
const headers = statusSheet.getRange(1, 1, 1, statusSheet.getLastColumn()).getValues();
const scriptProperties = PropertiesService.getScriptProperties();
const properties = scriptProperties.getProperties();
const propertiesKeys = Object.keys(properties);
scriptProperties.setProperty('stopFlag', 'false');


const onOpen = () => {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('PandaDoc')
    .addItem('Initial Spreadsheet Setup', 'indexSetup')
    .addToUi();
};

const indexSetup = () => {
  const lastRow = statusSheet.getLastRow();
  if (lastRow > 1) {
      errorHandler.alert("WARNING: Setup can only occur with no data in the Document_status Sheet.");
      return;
  }
  if (!propertiesKeys.includes("token")){
    errorHandler.alert("You must have at least one Bearer Token saved as a script property");
    return;
  }
  scriptProperties.setProperty("increment", 1);
  scriptProperties.setProperty("createDate", "2021-01-01T01:01:01.000000Z");
  triggers.createTriggers();
};


// ----IDEAS-----
//Full testing
//How does the code not keep on repeating workspace results?
//Sort docs by created date or sort workspaces into their own sheet?
//Handle throttling error
//When setting up triggers, could I not sort all of the script properties to be the same each time?
