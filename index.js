//To be done:
// Each Recipient Completed Time
//Document Versions...?
//Handle multiple rows, start of recovery script
//Build test scripts 

// Constants
let page = 1;
const ERRORS_SHEET_NAME = "Errors";
const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs");
const statusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Document_status");
const properties = PropertiesService.getScriptProperties().getProperties();
const propertiesKeys = Object.keys(properties);

//App: Dropdown menu
const onOpen = () => {
    let ui = SpreadsheetApp.getUi();
    ui.createMenu('PandaDoc')
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
    const errorsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ERRORS_SHEET_NAME);
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
            const workspaceProperty = getWorkspaceProperty(3, "name", key);
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
const documentStatus = (data, row, workspaceName, event) => {
    const headers = statusSheet.getRange(1, 1, 1, statusSheet.getLastColumn()).getValues();
    const columns = columnHeaders(headers)
    const values = statusSheet.getRange(row, 1, 1, statusSheet.getLastColumn()).getValues();
    const status = values[columns.status - 1];

    switch (status) {
        case "document.draft":
            basicInfo(row, data, columns, workspaceName);
            statusSheet.getRange(row, status).setValue("Draft");
            break;

        case "document.sent":
            basicInfo(row, data, columns, workspaceName);
            const createToSent = timeTo(data.date_created, data.date_modified);
            statusSheet.getRange(row, columns.status).setValue("Sent");
            statusSheet.getRange(row, columns.dateSent).setValue(data.date_modified);
            statusSheet.getRange(row, columns.timeCreatedToSent).setValue(createToSent);
            break;

        case "document.viewed":
            if (event === "document_state_changed") {
                basicInfo(row, data, columns, workspaceName);
                const sentToViewed = timeTo(values[0][4], data.date_modified);
                statusSheet.getRange(row, columns.status).setValue("Viewed");
                statusSheet.getRange(row, columns.dateViewed).setValue(data.date_modified);
                statusSheet.getRange(row, columns.timeSentToViewed).setValue(sentToViewed);
            } else if (event === 'recipient_completed') {
                const col = statusSheet.getLastColumn() + 1;
                const recipCompletedTime = timeTo(data.date_created, data.date_modified);
                statusSheet.getRange(row, col, 1, 2).setValues([
                    [data.action_by.email, data.action_date]
                ]);
                statusSheet.getRange(1, col, 1, 2).setValues([
                    ["Recipient Email", "Recipient Complete Date"]
                ]);
                //statusSheet.getRange(row, timeSentToViewed).setValue(sentToViewed);
            }
            break;

        case "document.waiting_approval":
            basicInfo(row, data, columns, workspaceName);
            statusSheet.getRange(row, columns.status).setValue("Waiting For Approval");
            statusSheet.getRange(row, dateSentForApproval).setValue(data.date_modified);
            break;

        case "document.rejected":
            basicInfo(row, data, columns, workspaceName);
            statusSheet.getRange(row, columns.status).setValue("Rejected");
            break;

        case "document.approved":
            basicInfo(row, data, columns, workspaceName);

            if (values[0][7]) {
                const timeToApprove = timeTo(values[0][7], data.date_modified);
                statusSheet.getRange(row, columns.timeToApproveDoc).setValue(timeToApprove);
            }

            statusSheet.getRange(row, columns.status).setValue("Approved");
            statusSheet.getRange(row, columns.dateApproved).setValue(data.date_modified);
            break;

        case "document.waiting_pay":
            basicInfo(row, data, columns, workspaceName);
            statusSheet.getRange(row, columns.status).setValue("Waiting For Payment");
            break;

        case "document.paid":
            basicInfo(row, data, columns, workspaceName);
            statusSheet.getRange(row, columns.status).setValue("Paid");
            break;

        case "document.completed":
            basicInfo(row, data, columns, workspaceName);

            if (values[0][4]) {
                const sentToComplete = timeTo(values[0][4], data.date_modified);
                statusSheet.getRange(row, columns.timeSentToCompleted).setValue(sentToComplete);
            }
            if (values[0][5]) {
                const viewToComplete = timeTo(values[0][5], data.date_modified);
                statusSheet.getRange(row, columns.timeViewedToCompleted).setValue(viewToComplete);
            }

            const createToComplete = timeTo(data.date_created, data.date_modified);
            statusSheet.getRange(row, columns.status).setValue("Completed");
            statusSheet.getRange(row, columns.dateCompleted).setValue(data.date_modified);
            statusSheet.getRange(row, columns.timeCreatedToCompleted).setValue(createToComplete);
            break;

        case "document.voided":
            basicInfo(row, data, columns, workspaceName);
            statusSheet.getRange(row, columns.status).setValue("Voided");
            break;
        case "document.declined":
            basicInfo(row, data, columns, workspaceName);
            statusSheet.getRange(row, columns.status).setValue("Declined");
            break;
    }
}
const basicInfo = (row, data, columns, workspaceName) => {
    statusSheet.getRange(row, columns.id).setValue(data.id);
    statusSheet.getRange(row, columns.workspace).setValue(workspaceName);
    statusSheet.getRange(row, columns.name).setValue(data.name);
    statusSheet.getRange(row, columns.dateCreated).setValue(data.date_created);
    statusSheet.getRange(row, columns.statusUnformat).setValue(data.status);
}

//Search first column by document ID
const searchId = (id) => {
    const lastRow = statusSheet.getLastRow();
    const columnValues = statusSheet.getRange(2, 1, lastRow).getValues();
    const searchResult = columnValues.findIndex(row => row[0] === id);
    const rowIndex = searchResult !== -1 ? searchResult + 2 : lastRow + 1;
    return rowIndex
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
const columnHeaders = () => {
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
    const key = workspaceProperties(name);
    while (true) {
        const {
            length,
            docs
        } = await listDocuments(key);
        if (length == 0) break;
        await eachDoc(docs, name, key);
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
    const createOptions = {
        'method': 'get',
        'headers': {
            'Authorization': `API-Key ${key}`,
            'Content-Type': 'application/json;charset=UTF-8'
        }
    };
    const response = UrlFetchApp.fetch(`https://api.pandadoc.com/public/v1/documents?page=${page}&count=100&order_by=date_created`, createOptions);
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
        if (doc.name.startsWith("[DEV]")) continue;
        const rowIndex = searchId(doc.id)
        if (rowIndex < statusSheet.getLastRow() + 1) {
            await checkRowStatus(doc, rowIndex, key, wsname)
        } else {
            documentStatus(doc, rowIndex, wsname, "document_state_changed");
        }
    }
};

//Recovery Process: check if existing doc has the correct status, if not update the row in the spreadsheet
const checkRowStatus = async (doc, row, key, wsname) => {
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
}
