const indexLoopThroughWorkspaces = () => {
    try {
        Logger.log("1: loop through workspaces")
        let pauseForTime = false;
        for (const key of propertiesKeys) {
            if (!key.startsWith("apiKey")) continue;
            const hasKeyBeenIterated = property.getValueFromScriptProperties(6, "hasKeyBeenIterated", key);
            if (hasKeyBeenIterated !== "false") continue;
    
            //check if script run time is up
            pauseForTime = triggers.terminateExecution("Recovery", "");
            if (pauseForTime) return;
    
            const workspaceName = property.getValueFromScriptProperties(5, "name", key);
            console.log(workspaceName);
            const modifiedDate = setModifiedDate();
            const filteredDocs = listDocsRecovery(modifiedDate, `API-Key ${properties[key]}`);
            const noIdsInSheet = getColumns(filteredDocs, workspaceName, properties[key]);
            console.log("Number of rows to add: " + noIdsInSheet.length);
    
            //check if script run time is up
            pauseForTime = triggers.terminateExecution("Recovery");
            if (pauseForTime) return;
    
            if (noIdsInSheet.length) {
                pdIndex.processListDocResultPublicDetails(noIdsInSheet, workspaceName, `API-Key ${properties[key]}`);
                //Insert corresponding blank rows
                statusSheet.insertRows(statusSheet.getLastRow() + 1, noIdsInSheet.length);
            };
            const slice = key.slice(6);
            console.log("hasKeyBeenIterated"+slice);
            //scriptProperties.setProperty("hasKeyBeenIterated"+slice, "true");
        }
        console.log("finished loop")
    } catch (error) {
        console.log(error)
    }

};

const setModifiedDate = () => {
    Logger.log("2: Set Modified date")
    let now = new Date();
    let threeMonthsAgo = new Date(now.getFullYear(), now.getMonth() - 3, now.getDate(), now.getHours(), now.getMinutes(), now.getSeconds(), now.getMilliseconds());
    let isoString = threeMonthsAgo.toISOString();
    let createDate = isoString.slice(0, 23) + "000Z";
    return createDate;
};

const listDocsRecovery = (modifiedDate, apiKey) => {
    Logger.log("3: list docs");
    const createOptions = {
        'method': 'get',
        'headers': {
            'Authorization': apiKey,
            'Content-Type': 'application/json;charset=UTF-8'
        }
    };
    let page = 1;
    let filteredDocsArr = [];

    while (true) {
        const response = UrlFetchApp.fetch(`https://api.pandadoc.com/public/v1/documents?page=${page}&count=100&order_by=date_created&modified_from=${modifiedDate}`, createOptions);
        const responseJson = JSON.parse(response);
        if (responseJson.results.length < 1) break;
        page++;

        const docsFiltered = responseJson.results.filter(doc => !doc.name.startsWith("[DEV]") && doc.version === "2");
        filteredDocsArr = filteredDocsArr.concat(docsFiltered);
    }
    return filteredDocsArr;
};

const getColumns = (docs, workspaceName, apiKey) => {
    Logger.log("4: sort columns");
    let noIdInSheet = [];
    const col = statusSheet.getRange(2, 1, statusSheet.getLastRow(), 6).getValues();

    docs.forEach((doc) => {
        let idExists = col.some((rowItem) => rowItem[0] === doc.id);
        if (!idExists) {
            noIdInSheet.push(doc);
        } else {
            let matchStatus = col.find((rowItem) => rowItem[0] === doc.id)[5];
            if (matchStatus !== doc.status) {
                let docArr = [];
                //const rowIndex = searchId(doc.id);
                docArr.push(doc);
                pdIndex.processListDocResultPublicDetails(docArr, workspaceName, `API-Key ${apiKey}`, "RecoveryUpdateStatus");
            }
        }
    });

    return noIdInSheet;
};
/*
//Search first column by document ID
const searchId = (id) => {
    const lastRow = statusSheet.getLastRow();
    const columnValues = statusSheet.getRange(2, 1, lastRow).getValues();
    const searchResult = columnValues.findIndex(row => row[0] === id);
    const rowIndex = searchResult !== -1 ? searchResult + 2 : lastRow + 1;
    return rowIndex
};
//Search: return row index if doc ID exists
Array.prototype.findIndex = function (search) {
    for (let i = 0; i < this.length; i++)
        if (this[i] == search) return i;

    return -1;
};*/

const recovery = {
    recoveryIndex: indexLoopThroughWorkspaces
}