const indexLoopThroughWorkspaces = (retries = 0) => {
    try {
        Logger.log("Loop through workspaces");
        for (const key of propertiesKeys) {
            if (!key.startsWith("apiKey")) continue;

            const hasKeyBeenIterated = property.getValueFromScriptProperties(6, "hasKeyBeenIterated", key);
            if (hasKeyBeenIterated !== "false") continue;

            // check if script run time is up
            const pauseForTime = triggers.terminateExecution("Recovery", "");
            if (pauseForTime) return;

            const workspaceName = property.getValueFromScriptProperties(6, "name", key);
            const modifiedDate = setModifiedDate();
            const filteredDocs = listDocsRecovery(modifiedDate, `API-Key ${properties[key]}`);
            const noIdsInSheet = getColumns(filteredDocs, workspaceName, properties[key]);
            console.log("Number of rows to add: " + noIdsInSheet.length);

            if (noIdsInSheet.length) {
                pdIndex.processListDocResult(noIdsInSheet, `API-Key ${properties[key]}`, workspaceName, "RecoveryAddRow");
                statusSheet.insertRows(statusSheet.getLastRow() + 1, noIdsInSheet.length);
            }

            const slice = key.slice(6);
            scriptProperties.setProperty("hasKeyBeenIterated" + slice, "true");
        }
        formatSheet.sortRowsByCreateDate();
        console.log("Finished loop");
        finalRefresh(retries)
    } catch (error) {
        console.log(error);
    }
};

const setModifiedDate = () => {
    let now = new Date();
    let oneMonthAgo = new Date(now.getFullYear(), now.getMonth() - 1, now.getDate(), now.getHours(), now.getMinutes(), now.getSeconds(), now.getMilliseconds());
    let isoString = oneMonthAgo.toISOString();
    let createDate = isoString.slice(0, 23) + "000Z";
    return createDate;
};

const listDocsRecovery = (modifiedDate, apiKey) => {
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
    let noIdInSheet = [];
    const col = statusSheet.getRange(2, 1, statusSheet.getLastRow(), 6).getValues();

    docs.forEach((doc) => {
        let idExists = col.some((rowItem) => rowItem[0] === doc.id);
        if (!idExists) {
            noIdInSheet.push(doc);
        } else {
            if (doc.status === "document.external_review") return;
            let matchStatus = col.find((rowItem) => rowItem[0] === doc.id)[5];
            if (matchStatus !== doc.status) {
                console.log("Status to be updated!");
                let docArr = [];
                docArr.push(doc);
                pdIndex.processListDocResult(docArr, `API-Key ${apiKey}`, workspaceName, "RecoveryUpdateStatus");
            }
        }
    });
    return noIdInSheet;
};

const finalRefresh = (retries) => {
    let arr = [];

    propertiesKeys.forEach(item => {
        if (item.startsWith("hasKeyBeenIterated") && scriptProperties.getProperty(item) === "false") {
            arr.push(item);
        }
    });
    if (arr.length && retries < 2) {
        console.log("Refreshing data")
        Utilities.sleep(3000);
        indexLoopThroughWorkspaces(retries + 1);
    }
};


const recovery = {
    recoveryIndex: indexLoopThroughWorkspaces
}