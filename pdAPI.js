//This file calls various PandaDoc API endpoints 

const listDocs = (key, date, page) => {
    Logger.log("4. list docs");
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

const getDocDetailsFromListDocResultPrivateAPI = (docs, key, retries = 0) => {
    Logger.log("5. Each doc");
    try {
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
        handleDocDetailsResponse.addRowFromPrivAPIResponse(jsonResponses, statusSheet.getLastRow() + 1);
        return docsFiltered;
    } catch (error) {
        if (retries > 2) {
            errorHandler.logAPIError(error);

            //Set script propery createDate
            const column = statusSheet.getSheetValues(statusSheet.getLastRow() - 19, headers[0].indexOf("Date Created") + 1, 20, 1);
            const filteredData = column.filter(arr => arr.some(val => val !== ''));
            const lastCreateDate = filteredData.length ? filteredData[filteredData.length - 1][0] : "2021-01-01T01:01:01.000000Z"
            scriptProperties.setProperty('createDate', lastCreateDate);

            throw new Error("Script terminated: Maximum number of errors exceeded");
        }
        console.log("Private API " + error);
        console.log(`Received error, retrying in 3 seconds... (attempt ${retries + 1} of 3)`);
        Utilities.sleep(3000);
        return getDocDetailsFromListDocResultPrivateAPI(docs, key, retries + 1);
    }
};

const getDocDetailsFromListDocResultPublicAPI = (docs, workspaceName, key, retries = 0) => {
    Logger.log("7. Form?");
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
        handleDocDetailsResponse.updateRowFromPubAPIResponse(jsonResponses, workspaceName);
    } catch (error) {
        if (retries > 2) {
            errorHandler.logAPIError(error);

            //Set script propery createDate
            const column = statusSheet.getSheetValues(statusSheet.getLastRow() - 19, headers[0].indexOf("Date Created") + 1, 20, 1);
            const filteredData = column.filter(arr => arr.some(val => val !== ''));
            const lastCreateDate = filteredData.length ? filteredData[filteredData.length - 1][0] : "2021-01-01T01:01:01.000000Z"
            scriptProperties.setProperty('createDate', lastCreateDate);

            throw new Error("Script terminated: Maximum number of errors exceeded");
        }
        console.log("Public API " + error);
        console.log(`Received error, retrying in 3 seconds... (attempt ${retries + 1} of 3)`);
        Utilities.sleep(3000);
        return getDocDetailsFromListDocResultPublicAPI(docs, workspaceName, key, retries + 1);
    }
};

const pdIndex = {
    listDocuments: listDocs,
    processListDocResult: getDocDetailsFromListDocResultPrivateAPI,
    processListDocResultPublicDetails: getDocDetailsFromListDocResultPublicAPI
};