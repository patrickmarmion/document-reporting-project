//To be done:
//+ 1 second to last list doc date after the script finishes running
//Forms?
//Auto Set Recovery Script (this could ultimately be linked to a report export?): https://developers.google.com/apps-script/guides/triggers/installable#time_driven_triggers
//Document Versions...?
//Build test scripts 
//Split Workspaces into different Sheets & Group?
//Document what this script achieves
//Long term: Grand Total reporting, Template Reporting, Product reporting, Renewal Reporting, Expiration Reporting 

// Constants
let page = 1;
const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs");
const statusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Document_status");
const errorsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Errors");
const scriptProperties = PropertiesService.getScriptProperties();
const properties = scriptProperties.getProperties();
const propertiesKeys = Object.keys(properties);

//App: Dropdown menu
const onOpen = () => {
    let ui = SpreadsheetApp.getUi();
    ui.createMenu('PandaDoc')
        .addItem('Initial Spreadsheet Setup', 'setup')
        .addSeparator()
        .addItem('Refresh Documents', 'selectWorkspace')
        .addToUi();
}
//App: Custom diolgue with workspace options
const selectWorkspace = () => {
    let html = HtmlService.createHtmlOutputFromFile('Workspace')
        .setWidth(400)
        .setHeight(200);
    SpreadsheetApp.getUi()
        .showModalDialog(html, 'Please Select Your Workspace');
}

// Log any errors to Error Sheet
const logError = (error) => {
    errorsSheet.appendRow([error]);
};

//Alert pop up
const setupAlert = (alertMessage) => {
    let ui = SpreadsheetApp.getUi();
    ui.alert(alertMessage);
};

//---WEBHOOK PROCESS FUNCTIONS---
//Catch webhook
const doPost = (e) => {
    try {
        const workspaceName = getWorkspaceName(e);
        if (!workspaceName) {
            logError("Webhook signature does not match. Key changed or payload has been modified!");
            throw new Error("Webhook signature does not match. Key changed or payload has been modified!");
        }

        const postData = JSON.parse(e.postData.contents)[0];
        const {
            event,
            data
        } = postData;

        if (event === "document_state_changed" && data.status === "document.completed") return;
        if (data.name.startsWith("[DEV]")) return;
        if (event === "document_deleted") {
            documentDeleted(data.id);
            return
        }

        addLastRow();
        // Write to the Log Sheet
        logs(data)

        // Write to Document Status Sheet    
        const rowIndex = searchId(data.id)
        documentStatus(data, rowIndex, workspaceName, event);

    } catch (error) {
        logError(error);
    }
    return HtmlService.createHtmlOutput("doPost received");
};

//Webhook Verification
const getGeneratedSignature = (input, secret) => {
    const signatureBytes = Utilities.computeHmacSha256Signature(input, secret);
    const generatedSignature = signatureBytes.reduce((str, byte) => {
        byte = (byte < 0 ? byte + 256 : byte).toString(16);
        return str + (byte.length === 1 ? "0" : "") + byte;
    }, "");
    return generatedSignature;
};

//Verification and determine workspace from Script Properties
const getWorkspaceName = (e) => {
    const {
        signature
    } = e.parameter;
    const input = e.postData.contents;

    for (const key of propertiesKeys) {
        if (!key.startsWith("sharedKey")) continue;

        const generatedSignature = getGeneratedSignature(input, properties[key]);

        if (signature === generatedSignature) {
            const workspaceProperty = getWorkspaceProperty(9, "name", key);
            return workspaceProperty;
        }
    }
    throw new Error("Webhook signature does not match. Key changed or payload has been modified!");
};

//Every trigger write to Log Sheet
const logs = (data) => {
    const values = [data.id, data.workspaceName, data.name, data.status, data.date_created, data.date_modified, data.expiration_date, data.created_by.email, data.grand_total.currency, data.grand_total.amount];
    if (data.template) values.push(data.template.id, data.template.name);
    logSheet.appendRow(values);
}

//Delete row when doc gets deleted
const documentDeleted = (id) => {
    const rowIndex = searchId(id);
    if (rowIndex === statusSheet.getLastRow() + 1) return
    statusSheet.deleteRow(rowIndex);
};

//---SHARED FUNCTIONS---
//Return Value from Script Properties
const getWorkspaceProperty = (num, prop, property) => {
    const workspaceName = property.slice(num);
    const workspaceNameValue = propertiesKeys
        .filter((key) => key.includes(workspaceName) && key.includes(prop))
        .reduce((cur, key) => {
            return Object.assign(cur, {
                [key]: properties[key]
            });
        }, {});
    const valueArr = Object.values(workspaceNameValue);
    return valueArr[0];
};

//Library of Column Headers in the Document Status Sheet
const columnHeaders = (headers) => {
    return {
        status: headers[0].indexOf("Status") + 1,
        dateSent: headers[0].indexOf("Date Sent") + 1,
        dateViewed: headers[0].indexOf("Date Viewed") + 1,
        dateCompleted: headers[0].indexOf("Date Completed") + 1,
        dateSentForApproval: headers[0].indexOf("Date Sent For Approval") + 1,
        dateApproved: headers[0].indexOf("Date Approved") + 1,
        timeCreatedToCompleted: headers[0].indexOf("Time Created to Completed (HH:MM:SS)") + 1,
        timeSentToCompleted: headers[0].indexOf("Time Sent to Completed (HH:MM:SS)") + 1,
        timeViewedToCompleted: headers[0].indexOf("Time Viewed to Completed (HH:MM:SS)") + 1,
        timeCreatedToSent: headers[0].indexOf("Time Created to Sent (HH:MM:SS)") + 1,
        timeToApproveDoc: headers[0].indexOf("Total Time to Approve (HH:MM:SS)") + 1,
        timeSentToViewed: headers[0].indexOf("Time Sent to First View (HH:MM:SS)") + 1,
        id: headers[0].indexOf("ID") + 1,
        workspace: headers[0].indexOf("Workspace Name") + 1,
        name: headers[0].indexOf("Document Name") + 1,
        dateCreated: headers[0].indexOf("Date Created") + 1,
        statusUnformat: headers[0].indexOf("Status Unformatted") + 1,
    }
};

//Sets the two status columns with both the formatted and unformatted status
const setStatusValue = (row, column, status, unformattedStatus) => {
    statusSheet.getRange(row, column.status).setValue(status);
    statusSheet.getRange(row, column.statusUnformat).setValue(unformattedStatus);
};

//The following functions handle all the different statuses, and updates the sheet accordingly 
const handleDocumentDraft = (data, row, columns) => {
    setStatusValue(row, columns, "Draft", "document.draft");
};
const handleDocumentSent = (data, row, columns) => {
    setStatusValue(row, columns, "Sent", "document.sent");
    const createToSent = timeTo(data.date_created, data.status === 1 ? data.date_sent : data.date_modified);
    statusSheet.getRange(row, columns.dateSent).setValue(data.status === 1 ? data.date_sent : data.date_modified);
    statusSheet.getRange(row, columns.timeCreatedToSent).setValue(createToSent);
};
const handleDocumentCompleted = (data, row, columns, values, event) => {
    setStatusValue(row, columns, "Completed", "document.completed");

    if (event === "Setup") {
        const createToSent = timeTo(data.date_created, data.date_sent);
        const sentToComplete = timeTo(data.date_sent, data.date_completed);
        const createToComplete = timeTo(data.date_created, data.date_completed);
        const rangeValues = [
            [data.date_sent, data.date_completed, createToSent, sentToComplete, createToComplete]
        ];
        const range = statusSheet.getRange(row, columns.dateSent, 1, rangeValues[0].length);
        range.setValues(rangeValues);
    } else {
        if (values[0][4]) {
            const sentToComplete = timeTo(values[0][4], data.date_modified);
            statusSheet.getRange(row, columns.timeSentToCompleted).setValue(sentToComplete);
        }
        if (values[0][5]) {
            const viewToComplete = timeTo(values[0][5], data.date_modified);
            statusSheet.getRange(row, columns.timeViewedToCompleted).setValue(viewToComplete);
        }
        const createToComplete = timeTo(data.date_created, data.date_modified);
        const rangeValues = [
            [data.date_modified, createToComplete]
        ];
        const range = statusSheet.getRange(row, columns.dateCompleted, 1, rangeValues[0].length);
        range.setValues(rangeValues);
        statusSheet.getRange(row, columns.timeCreatedToCompleted).setValue(createToComplete);
    }
};
const handleDocumentViewed = (data, row, columns, values, event) => {
    const col = statusSheet.getLastColumn() + 1;

    if (event === "recipient_completed" && data.status === "document.viewed") {
        const recipCompletedTime = timeTo(data.date_created, data.date_modified);
        statusSheet.getRange(row, col, 1, 3).setValues([
            [data.action_by.email, data.action_date, recipCompletedTime]
        ]);
        statusSheet.getRange(1, col, 1, 3).setValues([
            ["Recipient Email", "Recipient Complete Date", "Recipient Time to Complete (HH:MM:SS)"]
        ]);
        return;
    }

    setStatusValue(row, columns, "Viewed", "document.viewed");

    let sentToViewed = null;
    let dateViewed = null;

    if (data.status === 5) {
        sentToViewed = timeTo(data.date_sent, data.date_status_changed);
        dateViewed = data.date_status_changed;
    } else {
        dateViewed = data.date_modified;

        if (values[0][4]) {
            sentToViewed = timeTo(values[0][4], data.date_modified);
        }
    }

    if (sentToViewed !== null) statusSheet.getRange(row, columns.timeSentToViewed).setValue(sentToViewed);
    if (dateViewed !== null) statusSheet.getRange(row, columns.dateViewed).setValue(dateViewed);
};
const handleDocumentWaitingApproval = (data, row, columns) => {
    setStatusValue(row, columns, "Waiting For Approval", "document.waiting_approval");
    if (data.status !== 6) statusSheet.getRange(row, dateSentForApproval).setValue(data.date_modified);
};
const handleDocumentRejected = (data, row, columns) => {
    setStatusValue(row, columns, "Rejected", "document.rejected");
};
const handleDocumentApproved = (data, row, columns, values, event) => {
    setStatusValue(row, columns, "Approved", "document.approved");
    if (data.status === "document.approved" && event !== "Recovery") {
        const timeToApprove = timeTo(values[0][7], data.date_modified);
        statusSheet.getRange(row, columns.timeToApproveDoc).setValue(timeToApprove);
        statusSheet.getRange(row, columns.dateApproved).setValue(data.date_modified);
    } else if (data.status === "document.approved" && event === "Recovery") {
        statusSheet.getRange(row, columns.dateApproved).setValue(data.date_modified);
    } else {
        statusSheet.getRange(row, columns.dateApproved).setValue(data.date_status_changed);
    }
};
const handleDocumentWaitingPay = (data, row, columns) => {
    setStatusValue(row, columns, "Waiting For Payment", "document.waiting_pay");
};
const handleDocumentPaid = (data, row, columns) => {
    setStatusValue(row, columns, "Paid", "document.paid");
};
const handleDocumentVoid = (data, row, columns) => {
    setStatusValue(row, columns, "Voided", "document.voided");
};
const handleDocumentDeclined = (data, row, columns) => {
    setStatusValue(row, columns, "Declined", "document.declined");
};
const handleDocumentExternalReview = (data, row, columns) => {
    setStatusValue(row, columns, "External Review", "document.external_review");
};

//A map of of the different statuses and their corresponding handling function
const statusHandlers = {
    "document.draft": handleDocumentDraft,
    "document.sent": handleDocumentSent,
    "document.completed": handleDocumentCompleted,
    "document.viewed": handleDocumentViewed,
    "document.waiting_approval": handleDocumentWaitingApproval,
    "document.rejected": handleDocumentRejected,
    "document.approved": handleDocumentApproved,
    "document.waiting_pay": handleDocumentWaitingPay,
    "document.paid": handleDocumentPaid,
    "document.voided": handleDocumentVoid,
    "document.declined": handleDocumentDeclined,
    "document.external_review": handleDocumentExternalReview
};

//For each status, I patch the first 4 columns with the document's ID, Workspace, Name, Date of Creation. Not called during Recovery when doc already exists.
const basicInfo = (row, data, workspaceName) => {
    statusSheet.getRange(row, 1, 1, 4).setValues([
        [data.id, workspaceName, data.name, data.date_created]
    ]);
};

//Based on doc status & the event calls the relevant handler to write doc details to row
const documentStatus = async (data, row, workspaceName, event, retries = 0) => {
    try {
        const headers = statusSheet.getRange(1, 1, 1, statusSheet.getLastColumn()).getValues();
        const columns = columnHeaders(headers);
        const values = statusSheet.getRange(row, 1, 1, statusSheet.getLastColumn()).getValues();
        let status = data.status;

        if (event !== "Recovery") {
            basicInfo(row, data, workspaceName);
            const statusMap = ['document.draft', 'document.sent', 'document.completed', '', '', 'document.viewed', 'document.waiting_approval', 'document.approved', 'document.rejected', 'document.waiting_pay', 'document.paid', 'document.voided', 'document.declined', 'document.external_review']
            if (typeof status === 'number') status = statusMap[data.status];
        }

        const handler = statusHandlers[status];
        if (handler) {
            handler(data, row, columns, values, event);
        }

    } catch (error) {
        if (retries > 2) {
            logError(error);
        }
        console.log(error);
        console.log(`Received error, retrying in 3 seconds... (attempt ${retries + 1} of 3)`);
        Utilities.sleep(3000);
        return await documentStatus(data, row, workspaceName, event, retries + 1);
    }
};

//Search first column by document ID
const searchId = (id) => {
    try {
        const lastRow = statusSheet.getLastRow();
        const columnValues = statusSheet.getRange(2, 1, lastRow).getValues();
        const searchResult = columnValues.findIndex(row => row[0] === id);
        const rowIndex = searchResult !== -1 ? searchResult + 2 : lastRow + 1;
        return rowIndex
    } catch (error) {
        logError(error);
    }
}
//Search: return row index if doc ID exists
Array.prototype.findIndex = function (search) {
    for (let i = 0; i < this.length; i++)
        if (this[i] == search) return i;

    return -1;
}

//Calculation: time between two events
const timeTo = (timeFirst, timeSecond) => {
    const earlier = new Date(timeFirst);
    const later = new Date(timeSecond);
    const diffInSeconds = Math.floor((later - earlier) / 1000);
    const hms = (seconds) => {
        return [3600, 60]
            .reduceRight(
                (p, b) => r => [Math.floor(r / b)].concat(p(r % b)),
                r => [r]
            )(seconds)
            .map(a => a.toString().padStart(2, '0'))
            .join(':');
    }
    const duration = hms(diffInSeconds);
    return duration
}

//---RECOVERY PROCESS---
//Index Function
const wsName = async (name) => {
    try {
        const key = workspaceProperties(name);
        while (true) {
            const {
                length,
                docs
            } = await listDocuments(`API-Key ${key}`);
            if (length == 0) break;
            await eachDoc(docs, name, `API-Key ${key}`);
        }
        finishScriptRun();
    } catch (error) {
        logError(error);
    }
}

//Recovery Process: Determine relevant API Key
const workspaceProperties = (name) => {
    for (const property in properties) {
        if (!property.startsWith("name")) continue;
        if (properties[property] === name) {
            return getWorkspaceProperty(14, "authorization", property);
        }
    }
    throw new Error("Cannot find Workspace from script properties");
};

//Recovery Process: List Document Endpoint 
const listDocuments = async (key) => {
    let fromDate = scriptProperties.getProperty("listDocDate");
    const createOptions = {
        'method': 'get',
        'headers': {
            'Authorization': key,
            'Content-Type': 'application/json;charset=UTF-8'
        }
    };
    const response = UrlFetchApp.fetch(`https://api.pandadoc.com/public/v1/documents?page=${page}&count=100&order_by=date_created&created_from=${fromDate}`, createOptions);
    const responseJson = JSON.parse(response);
    page++
    return {
        length: responseJson.results.length,
        docs: responseJson.results,
    }
}

//Recovery Process: For each document check if it exists in the spreadsheet, if so call check status function, if not append to the bottom of the sheet
const eachDoc = async (docs, wsname, key) => {
    for (const doc of docs) {
        addLastRow();
        if (doc.version !== "2") continue;
        if (doc.name.startsWith("[DEV]")) continue;
        if (key.startsWith("Bearer")) {
            const document = await getDocDetails(doc.id, key);
            documentStatus(document, statusSheet.getLastRow() + 1, wsname, "Setup");
            continue;
        }

        const rowIndex = searchId(doc.id)
        if (rowIndex > statusSheet.getLastRow() + 1) {
            documentStatus(doc, rowIndex, wsname, "Recovery");
        } else {
            await checkRowStatus(doc, rowIndex, key, wsname)
        }
    }
};

//Recovery Process: check if existing doc has the correct status, if not update the row in the spreadsheet
const checkRowStatus = async (doc, row, key, wsname) => {
    try {
        const headers = statusSheet.getRange(1, 1, 1, statusSheet.getLastColumn()).getValues();
        const statusIndex = headers[0].indexOf("Status Unformatted") + 1;
        const statusUnformat = statusSheet.getRange(row, statusIndex).getValues();

        if (doc.status !== statusUnformat[0][0]) {
            const createOptions = {
                'method': 'get',
                'headers': {
                    'Authorization': `API-Key ${key}`,
                    'Content-Type': 'application/json;charset=UTF-8'
                }
            };
            const response = UrlFetchApp.fetch(`https://api.pandadoc.com/public/v1/documents/${doc.id}/details`, createOptions);
            const responseJson = JSON.parse(response);
            documentStatus(responseJson, row, wsname, "document_state_changed");
        }
        return
    } catch (error) {
        logError(error);
    }
};

//---SETUP PROCESS---
//Index for the setup script
const setup = async () => {
    try {
        createTrigger();
        const lastRow = statusSheet.getLastRow();
        const errorValue = errorsSheet.getRange(1, 1).getValues();
        if (lastRow > 1 && errorValue[0][0].startsWith("Maximum Script Execution")) {
            setupAlert('WARNING: Setup can only occur with no data in the Document_status Sheet.');
            return;
        }
        for (const key of propertiesKeys) {
            if (!key.startsWith("token")) continue;
            const workspaceName = getWorkspaceProperty(5, "name", key);
            while (true) {
                const {
                    length,
                    docs
                } = await listDocuments(`Bearer ${properties[key]}`);
                if (length === 0) break;
                await eachDoc(docs, workspaceName, `Bearer ${properties[key]}`);
            }
        }
        finishScriptRun();
    } catch (error) {
        if (error.message.startsWith("Maximum Script Execution Time Approaching")) {
            logError("Maximum Script Execution Time Approaching, please restart the Initial Script");
            return;
        }
        logError(error);
    }
};

//Get document details calling Private API for full details
const getDocDetails = async (id, key) => {
    const createOptions = {
        'method': 'get',
        'headers': {
            'Authorization': key,
            'Content-Type': 'application/json;charset=UTF-8'
        }
    };
    const response = UrlFetchApp.fetch(`https://api.pandadoc.com/documents/${id}`, createOptions);
    const responseJson = JSON.parse(response);
    return responseJson
};

//With each new document a new row is programmatically added
const addLastRow = () => {
    try {
        let lastBlankRowStatusSheet = statusSheet.getLastRow() + 99;
        statusSheet.insertRowAfter(lastBlankRowStatusSheet);
        let lastBlankRowLogSheet = logSheet.getLastRow() + 99;
        logSheet.insertRowAfter(lastBlankRowLogSheet);
    } catch (error) {
        logError(error);
    }
};

//GAS only allows 30 mins of execution this 
const createTrigger = () => {
    ScriptApp.newTrigger("maximumScriptTime")
        .timeBased()
        .after(1750000)
        .create();
};

const maximumScriptTime = () => {
    const lastRow = statusSheet.getLastRow();
    const createDate = statusSheet.getRange(lastRow, 4).getValues();
    const createDateParse = new Date(createDate[0][0]);
    const createDatePlusOneSec = new Date(createDateParse.getFullYear(), createDateParse.getMonth(), createDateParse.getDate(), createDateParse.getHours(), createDateParse.getMinutes(), createDateParse.getSeconds() + 1);
    const createDateUTC = utcString(createDatePlusOneSec)
    scriptProperties.setProperty("listDocDate", createDateUTC);
    ScriptApp.deleteTrigger("maximumScriptTime");
    throw new Error("Maximum Script Execution Time Approaching, please restart the Initial Script");
};

const finishScriptRun = () => {
    const now = new Date();
    const oneYearAgo = new Date(now.getFullYear() - 1, now.getMonth(), now.getDate(), now.getHours(), now.getMinutes(), now.getSeconds());
    const oneYearAgoUTC = utcString(oneYearAgo);
    scriptProperties.setProperty("listDocDate", oneYearAgoUTC);
    errorsSheet.deleteColumn(1);
};

const utcString = (time) => {
    let year = time.getUTCFullYear();
    let month = ("0" + (time.getUTCMonth() + 1)).slice(-2);
    let day = ("0" + time.getUTCDate()).slice(-2);
    let hours = ("0" + time.getUTCHours()).slice(-2);
    let minutes = ("0" + time.getUTCMinutes()).slice(-2);
    let seconds = ("0" + time.getUTCSeconds()).slice(-2);
    let milliseconds = ("00000" + time.getUTCMilliseconds())

    let dateString = year + "-" + month + "-" + day + "T" + hours + ":" + minutes + ":" + seconds + "." + milliseconds + "Z";
    return dateString;
}