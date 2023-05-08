//This file calls various PandaDoc API endpoints 

const listDocs = (key, date, page) => {
    Logger.log("3. list docs");
    const createOptions = {
        'method': 'get',
        'headers': {
            'Authorization': key,
            'Content-Type': 'application/json;charset=UTF-8'
        }
    };

    const response = UrlFetchApp.fetch(`https://api.pandadoc.com/public/v1/documents?page=${page}&count=100&order_by=date_created&created_from=${date}`, createOptions);
    const responseJson = JSON.parse(response);
    return {
        length: responseJson.results.length,
        docs: responseJson.results,
    }
};

const fetchAllListDocResult = (docs, workspaceName, key, retries = 0) => {
    Logger.log("4. Each doc");
    try {
        if (scriptProperties.getProperty('stopFlag') === 'true') {
            scriptProperties.setProperty('createDate', docs[0].date_created);
            return false;
        };
        const docsFiltered = docs.filter(doc => !doc.name.startsWith("[DEV]") && doc.version === "2")
        const docsMap = docsFiltered.map(doc => `https://api.pandadoc.com/documents/${doc.id}`);

        const requests = docsMap.map(url => {
            return {
                url: url,
                method: "GET",
                headers: {
                    'Authorization': key,
                    'Content-Type': 'application/json;charset=UTF-8'
                }
            };
        });

        const responses = UrlFetchApp.fetchAll(requests);
        const jsonResponses = responses.map(response => JSON.parse(response.getContentText()));
        handleStatus.documentStatus(jsonResponses, statusSheet.getLastRow() + 1, workspaceName, "Setup");
        return docsFiltered;
    } catch (error) {
        if (retries > 1) {
            errorHandler.logAPIError(error);
            return false;
        }
        console.log(error);
        console.log(`Received error, retrying in 2 seconds... (attempt ${retries + 1} of 2)`);
        Utilities.sleep(2000);
        return fetchAllListDocResult(docs, workspaceName, key, retries + 1);
    }
};

const fetchAllListDocResultForms = (docs, key, retries = 0) => {
    Logger.log("6. Form?");
    try {
        const docsMap = docs.map(doc => `https://api.pandadoc.com/public/v1/documents/${doc.id}/details`);
        const requests = docsMap.map(url => {
            return {
                url: url,
                method: "GET",
                headers: {
                    'Authorization': key,
                    'Content-Type': 'application/json;charset=UTF-8'
                }
            };
        });

        const responses = UrlFetchApp.fetchAll(requests);
        const jsonResponses = responses.map(response => JSON.parse(response.getContentText()));
        form.handleForm(jsonResponses);

        if (scriptProperties.getProperty('stopFlag') === 'true') {
            const createDate = statusSheet.getRange(statusSheet.getLastRow(), 4).getValues();
            scriptProperties.setProperty('createDate', createDate[0][0]);
            return false;
        };
        return true;
    } catch (error) {
        if (retries > 1) {
            errorHandler.logAPIError(error);
            return false;
        }
        console.log(error);
        console.log(`Received error, retrying in 2 seconds... (attempt ${retries + 1} of 2)`);
        Utilities.sleep(2000);
        return fetchAllListDocResultForms(docs, key, retries + 1);
    }
};

const pdIndex = {
    listDocuments: listDocs,
    processListDocResult: fetchAllListDocResult,
    checkIfForm: fetchAllListDocResultForms
};