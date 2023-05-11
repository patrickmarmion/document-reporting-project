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
    .addSeparator()
    .addItem('Refresh Documents', 'indexRecovery')
    .addToUi();
};

const indexSetup = () => {
  const lastRow = statusSheet.getLastRow();
  if (lastRow > 1) {
    errorHandler.alert("WARNING: Setup can only occur with no data in the Document_status Sheet.");
    return;
  };
  let arr = [];
  propertiesKeys.forEach(item => {
    if (item.startsWith("token")) {
      arr.push(item);
    }
  });
  if (!arr.length) {
    errorHandler.alert("You must have at least one Bearer Token saved as a script property");
    return;
  };
  scriptProperties.setProperty("increment", 1);
  scriptProperties.setProperty("createDate", "2021-01-01T01:01:01.000000Z");
  triggers.createTriggers();
};

const indexRecovery = () => {
  const lastRow = statusSheet.getLastRow();
  let arr = [];

  if (lastRow < 2) {
    errorHandler.alert("WARNING: Recovery can only occur with data in the Document_status Sheet.");
    return;
  };
  propertiesKeys.forEach(item => {
    if (item.startsWith("key")) {
      arr.push(item);
    };
    //Return all hasKeyBeenIterated back to false
    if (item.startsWith("hasKeyBeenIterated")){
      scriptProperties.setProperty(item, "false");
    }
  });
  if (!arr.length) {
    errorHandler.alert("You must have at least one API Key saved as a script property");
    return;
  };

  recovery.recoveryIndex();
}

// ----IDEAS-----
//Recovery Script: manually triggered. 
//Will need to create a pauseForTime, which creates another time based trigger.
//For loop through apiKeys in script properties => match this to an increment (hasKeyBeenIterated), which at the end of the recovery return each back to false
//Could add here to sort the sheet by CreateDate?

//Sort docs by created date or sort workspaces into their own sheet?
