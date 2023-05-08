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
  triggers.createTriggers();
};


// ----IDEAS-----
//Beautify Provider / Form process
//Grand Total
//Bulk add rows
//End of the script run, check for any rows with the same IDs