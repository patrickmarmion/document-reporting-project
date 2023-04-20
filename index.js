//To be done:
//Initial Script: catch for timeout, mark script properties with latest doc created date
//Forms?
//Auto Set Recovery Script: https://developers.google.com/apps-script/guides/triggers/installable#time_driven_triggers
//Document Versions...?
//Build test scripts 
//Split Workspaces into different Sheets
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

//Catch Hook
const doPost = (e) => {
    try {
        const workspaceName = getWorkspaceName(e);
        if (!workspaceName) throw new Error("Webhook signature does not match. Key changed or payload has been modified!");

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
}

// Error Handling
const logError = (error) => {
    errorsSheet.appendRow([error]);
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
}

//Based on doc status write details to row
const documentStatus = async (data, row, workspaceName, event, retries = 0) => {
    try {
        const headers = statusSheet.getRange(1, 1, 1, statusSheet.getLastColumn()).getValues();
        const columns = columnHeaders(headers)
        const values = statusSheet.getRange(row, 1, 1, statusSheet.getLastColumn()).getValues();
        basicInfo(row, data, columns, workspaceName);
        let createToSent;
        let sentToViewed;
        switch (true) {
            case data.status === "document.draft" || data.status === 0:
                statusSheet.getRange(row, columns.status).setValue("Draft");
                statusSheet.getRange(row, columns.statusUnformat).setValue("document.draft");
                break;

            case data.status === "document.sent" || data.status === 1:
                statusSheet.getRange(row, columns.status).setValue("Sent");
                statusSheet.getRange(row, columns.statusUnformat).setValue("document.sent");
                if (data.status === 1) {
                    createToSent = timeTo(data.date_created, data.date_sent);
                    statusSheet.getRange(row, columns.dateSent).setValue(data.date_sent);
                } else {
                    createToSent = timeTo(data.date_created, data.date_modified);
                    statusSheet.getRange(row, columns.dateSent).setValue(data.date_modified);
                }
                statusSheet.getRange(row, columns.timeCreatedToSent).setValue(createToSent);
                break;

            case event === "recipient_completed" && data.status === "document.viewed":
                const col = statusSheet.getLastColumn() + 1;
                const recipCompletedTime = timeTo(data.date_created, data.date_modified);
                statusSheet.getRange(row, col, 1, 3).setValues([
                    [data.action_by.email, data.action_date, recipCompletedTime]
                ]);
                statusSheet.getRange(1, col, 1, 3).setValues([
                    ["Recipient Email", "Recipient Complete Date", "Recipient Time to Complete (HH:MM:SS)"]
                ]);
                break;

            case data.status === "document.viewed" || data.status === 5:
                statusSheet.getRange(row, columns.status).setValue("Viewed");
                statusSheet.getRange(row, columns.statusUnformat).setValue("document.viewed");
                if (data.status === 5) {
                    sentToViewed = timeTo(data.date_sent, data.date_status_changed);
                    statusSheet.getRange(row, columns.dateViewed).setValue(data.date_status_changed);
                    statusSheet.getRange(row, columns.timeSentToViewed).setValue(sentToViewed);
                } else if (values[0][4]) {
                    sentToViewed = timeTo(values[0][4], data.date_modified);
                    statusSheet.getRange(row, columns.timeSentToViewed).setValue(sentToViewed);
                    statusSheet.getRange(row, columns.dateViewed).setValue(data.date_modified);
                } else {
                    statusSheet.getRange(row, columns.dateViewed).setValue(data.date_modified);
                }
                break;

            case data.status === "document.waiting_approval" || data.status === 6:
                statusSheet.getRange(row, columns.status).setValue("Waiting For Approval");
                statusSheet.getRange(row, columns.statusUnformat).setValue("document.waiting_approval");
                if (data.status !== 6) statusSheet.getRange(row, dateSentForApproval).setValue(data.date_modified);
                break;

            case data.status === "document.rejected" || data.status === 8:
                statusSheet.getRange(row, columns.status).setValue("Rejected");
                statusSheet.getRange(row, columns.statusUnformat).setValue("document.rejected");
                break;

            case data.status === "document.approved" || data.status === 7:
                statusSheet.getRange(row, columns.status).setValue("Approved");
                statusSheet.getRange(row, columns.statusUnformat).setValue("document.approved");
                if (data.status === "document.approved" && event !== "Recovery") {
                    const timeToApprove = timeTo(values[0][7], data.date_modified);
                    statusSheet.getRange(row, columns.timeToApproveDoc).setValue(timeToApprove);
                    statusSheet.getRange(row, columns.dateApproved).setValue(data.date_modified);
                } else if (data.status === "document.approved" && event === "Recovery") {
                    statusSheet.getRange(row, columns.dateApproved).setValue(data.date_modified);
                } else {
                    statusSheet.getRange(row, columns.dateApproved).setValue(data.date_status_changed);
                }
                break;

            case data.status === "document.waiting_pay" || data.status === 9:
                statusSheet.getRange(row, columns.status).setValue("Waiting For Payment");
                statusSheet.getRange(row, columns.statusUnformat).setValue("document.waiting_pay");
                break;
            case data.status === "document.paid" || data.status === 10:
                statusSheet.getRange(row, columns.status).setValue("Paid");
                statusSheet.getRange(row, columns.statusUnformat).setValue("document.paid");
                break;

            case data.status === "document.completed" || data.status === 2:
                statusSheet.getRange(row, columns.status).setValue("Completed");
                statusSheet.getRange(row, columns.statusUnformat).setValue("document.completed");

                if (event === "Setup") {
                    statusSheet.getRange(row, columns.dateSent).setValue(data.date_sent);
                    statusSheet.getRange(row, columns.dateCompleted).setValue(data.date_completed);
                    let createToSent = timeTo(data.date_created, data.date_sent);
                    let sentToComplete = timeTo(data.date_sent, data.date_completed);
                    let createToComplete = timeTo(data.date_created, data.date_completed);
                    statusSheet.getRange(row, columns.timeSentToCompleted).setValue(sentToComplete);
                    statusSheet.getRange(row, columns.timeCreatedToCompleted).setValue(createToComplete);
                    statusSheet.getRange(row, columns.timeCreatedToSent).setValue(createToSent);
                } else {
                    if (values[0][4]) {
                        let sentToComplete = timeTo(values[0][4], data.date_modified);
                        statusSheet.getRange(row, columns.timeSentToCompleted).setValue(sentToComplete);
                    }
                    if (values[0][5]) {
                        let viewToComplete = timeTo(values[0][5], data.date_modified);
                        statusSheet.getRange(row, columns.timeViewedToCompleted).setValue(viewToComplete);
                    }

                    let createToComplete = timeTo(data.date_created, data.date_modified);
                    statusSheet.getRange(row, columns.dateCompleted).setValue(data.date_modified);
                    statusSheet.getRange(row, columns.timeCreatedToCompleted).setValue(createToComplete);
                }
                break;

            case data.status === "document.voided" || data.status === 11:
                statusSheet.getRange(row, columns.status).setValue("Voided");
                statusSheet.getRange(row, columns.statusUnformat).setValue("document.voided");
                break;
            case data.status === "document.declined" || data.status === 12:
                statusSheet.getRange(row, columns.status).setValue("Declined");
                statusSheet.getRange(row, columns.statusUnformat).setValue("document.declined");
                break;
            case data.status === "external_review" || data.status === 13:
                statusSheet.getRange(row, columns.status).setValue("External Review");
                statusSheet.getRange(row, columns.statusUnformat).setValue("document.external_review");
                break;
        }
    } catch (error) {
        if (retries > 2) {
            logError(error);
        }
        console.log(error.response.data)
        console.log(`Received error, retrying in 3 seconds... (attempt ${retries + 1} of 3)`);
        await new Promise(resolve => setTimeout(resolve, 3000));
        return await documentStatus(data, row, workspaceName, event, retries + 1);
    }
}
const basicInfo = (row, data, columns, workspaceName) => {
    try {
        statusSheet.getRange(row, columns.id).setValue(data.id);
        statusSheet.getRange(row, columns.workspace).setValue(workspaceName);
        statusSheet.getRange(row, columns.name).setValue(data.name);
        statusSheet.getRange(row, columns.dateCreated).setValue(data.date_created);
    } catch (error) {
        logError(error);
    }
}

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
    for (var i = 0; i < this.length; i++)
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
}

//Recovery Process: Index Function
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
}

const setup = async () => {
    try {
        const lastRow = statusSheet.getLastRow();
        const errorValue = errorsSheet.getRange(1, 1).getValues();
        if (lastRow > 1 && !errorValue[0][0].startsWith("Maximum Script Execution")) {
            setupAlert();
            return
        }
        createTrigger()
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
    } catch (error) {
        logError(error);
    }
}

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

const setupAlert = () => {
    let ui = SpreadsheetApp.getUi();
    ui.alert('WARNING: Setup can only occur with no data in the Document_status Sheet.');
}

const addLastRow = () => {
    try {
        let lastBlankRowStatusSheet = statusSheet.getLastRow() + 99;
        statusSheet.insertRowAfter(lastBlankRowStatusSheet);

        let lastBlankRowLogSheet = logSheet.getLastRow() + 99;
        logSheet.insertRowAfter(lastBlankRowLogSheet);
    } catch (error) {
        logError(error);
    }
}

const createTrigger = () => {
    ScriptApp.newTrigger("maximumScriptTime")
    .timeBased()
    .after(1755000)
    .create();
}
const maximumScriptTime = () => {
    const lastRow = statusSheet.getLastRow();
    const createDate = statusSheet.getRange(lastRow, 4).getValues();
    scriptProperties.setProperty("listDocDate", createDate[0][0]);
    throw new Error("Maximum Script Execution Time Approaching, please restart the Initial Script");
    //End Script Run
}