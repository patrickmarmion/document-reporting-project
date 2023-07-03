//This file calls various PandaDoc API endpoints 

const listDocs = (key, date, page) => {
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

const getDocDetailsFromListDocResult = (docs, key, workspaceName, eventRec, retries = 0) => {
    Logger.log("Get Doc Details from API");
    try {
        const filteredDocs = docs.filter(doc => !doc.name.startsWith("[DEV]") && doc.version === "2");

        const publicAPIURLs = filteredDocs.map(doc => `https://api.pandadoc.com/public/v1/documents/${doc.id}/details`);
        const publicAPIRequests = publicAPIURLs.map(url => {
            return {
                url: url,
                method: "GET",
                headers: {
                    'Authorization': key,
                    'Content-Type': 'application/json;charset=UTF-8'
                }
            };
        });

        const publicAPIResponses = UrlFetchApp.fetchAll(publicAPIRequests);
        const publicAPIJsonResponses = publicAPIResponses.map(response => JSON.parse(response.getContentText()));

        if (eventRec === "RecoveryUpdateStatus") {
            handleDocDetailsResponse.wrongStatus(publicAPIJsonResponses, workspaceName);
            return;
        };
        if (eventRec === "RecoveryAddRow") {
            handleDocDetailsResponse.updateRowFromPubAPIResponse(publicAPIJsonResponses, workspaceName)
            return;
        }

        const privateAPIURLs = filteredDocs.map(doc => `https://api.pandadoc.com/documents/${doc.id}`);
        const privateAPIRequests = privateAPIURLs.map(url => {
            return {
                url: url,
                method: "GET",
                headers: {
                    'Authorization': key,
                    'Content-Type': 'application/json;charset=UTF-8'
                }
            };
        });

        const privateAPIResponses = UrlFetchApp.fetchAll(privateAPIRequests);
        const privateAPIJsonResponses = privateAPIResponses.map(response => JSON.parse(response.getContentText()));
        const privateAPIMap = handleDocDetailsResponse.privAPIResponseMap(privateAPIJsonResponses);
        handleDocDetailsResponse.updateRowFromPubAPIResponse(publicAPIJsonResponses, workspaceName, privateAPIMap);

    } catch (error) {
        if (retries > 2) {
            errorHandler.logAPIError(error);
            const lastRow = statusSheet.getLastRow();
            const values = statusSheet.getRange(`D1:D${lastRow}`).getValues().reverse();
            const lastCreateDate = values.find(([value]) => value !== '')?.[0] || "2021-01-01T01:01:01.000000Z";
            scriptProperties.setProperty('createDate', lastCreateDate);

            throw new Error("Script terminated: Maximum number of errors exceeded");
        }

        console.log("API " + error);
        console.log(`Received error, retrying in 3 seconds... (attempt ${retries + 1} of 3)`);
        Utilities.sleep(3000);
        return getDocDetailsFromListDocResult(docs, key, workspaceName, eventRec, retries + 1);
    }
};

const pdIndex = {
    listDocuments: listDocs,
    processListDocResult: getDocDetailsFromListDocResult
};