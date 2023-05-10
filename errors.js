// Log any errors to Error Sheet
const logError = (error) => {
    errorsSheet.appendRow([error]);
};

//Alert pop up
const sendAlert = (alertMessage) => {
    let ui = SpreadsheetApp.getUi();
    ui.alert(alertMessage);
};

const errorHandler = {
    logAPIError: logError,
    alert: sendAlert
};