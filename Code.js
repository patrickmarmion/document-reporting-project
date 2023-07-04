const statusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Document_status");
const errorsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Errors");
const headers = statusSheet.getRange(1, 1, 1, statusSheet.getLastColumn()).getValues();
const scriptProperties = PropertiesService.getScriptProperties();
let initialProperties = scriptProperties.getProperties();
const propertiesKeys = Object.keys(initialProperties);
scriptProperties.setProperty('stopFlag', 'false');

/**
 * Handles the onOpen event triggered when the spreadsheet is opened.
 * Creates a custom menu in the UI.
 *
 * @returns {void}
 */
const onOpen = () => {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('PandaDoc')
    .addItem('Initial Spreadsheet Setup', 'indexSetup')
    .addSeparator()
    .addItem('Refresh Documents', 'indexRecovery')
    .addToUi();
};

/**
 * Performs the initial setup for the spreadsheet.
 * Checks for existing data and required script properties.
 *
 * @returns {void}
 */
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
  scriptProperties.setProperty("createDate", "2021-01-01T01:01:01.000000Z");
  triggers.createTriggers();
};

const indexRecovery = () => {
  const lastRow = statusSheet.getLastRow();
  if (lastRow < 2) {
    errorHandler.alert("WARNING: Recovery can only occur with data in the Document_status Sheet.");
    return;
  };

  let apiKeyFound = false;
  propertiesKeys.forEach(item => {
    if (item.startsWith("apiKey")) {
      apiKeyFound = true;
    }
    if (item.startsWith("hasKeyBeenIterated")) {
      scriptProperties.setProperty(item, "false");
    }
  });
  if (!apiKeyFound) {
    errorHandler.alert("You must have at least one API Key saved as a script property");
    return;
  };

  recovery.recoveryIndex();
};

//Catch webhook
const doPost = (e) => {
  try {
    const workspaceName = handleWebook.verifyWebhookSignature(e);
    if (!workspaceName) {
      errorHandler.logAPIError("Error: No workspace name returned");
      errorHandler.logAPIError(workspaceName);
      return;
    };

    const postData = JSON.parse(e.postData.contents);
    const {
      event,
      data
    } = postData[0];
    let dataArr = [];
    dataArr.push(data);

    const lastRow = statusSheet.getLastRow();
    const values = statusSheet.getRange(`A1:A${lastRow}`).getValues();
    const rowIndex = values.findIndex(data.id) > 1 ? values.findIndex(data.id) + 1 : statusSheet.getLastRow() + 1;

    if (data.name.startsWith("[DEV]")) return;
    if (event === "document_state_changed" && data.status === "document.completed") return;

    if (event === "recipient_completed") {
      if (rowIndex !== statusSheet.getLastRow() + 1) {
        handleWebook.webhookRecipientCompleted(dataArr, rowIndex);
      };
      return;
    };
    if (event === "document_deleted") {
      handleWebook.documentDeleted(data.id, rowIndex);
      return;
    };

    statusSheet.insertRows(statusSheet.getLastRow() + 1, 1);
    rowIndex === statusSheet.getLastRow() + 1 ? handleWebook.webhookAddRow(dataArr, workspaceName, rowIndex) : handleWebook.webhookUpdateRow(dataArr, rowIndex);

  } catch (error) {
    errorHandler.logAPIError(error);
  }
  return HtmlService.createHtmlOutput("doPost received");
};

Array.prototype.findIndex = function (search) {
  for (let i = 0; i < this.length; i++) {
    if (this[i][0] === search) {
      return i;
    }
  }
  return -1;
};