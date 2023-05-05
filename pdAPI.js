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
    //Change the line URL below after testing.
    const response = UrlFetchApp.fetch(`https://api.pandadoc.com/public/v1/documents?page=${page}&count=100&order_by=date_created&created_from=${date}&status=0`, createOptions);
    const responseJson = JSON.parse(response);
    return {
        length: responseJson.results.length,
        docs: responseJson.results,
    }
};

const loopThroughListDocResult = (docs, workspaceName, key) => {
    Logger.log("4. each docs");
    for (const doc of docs) {
        //addLastRow(); Need to add this
        if (scriptProperties.getProperty('stopFlag') === 'true') {
            scriptProperties.setProperty('createDate', doc.date_created);
            return false;
        }
        if (doc.version !== "2") continue;
        if (doc.name.startsWith("[DEV]")) continue;
        if (key.startsWith("Bearer")) {
            const documentDetailsPriv = getDocDetails(key, `https://api.pandadoc.com/documents/${doc.id}`);
            handleStatus.documentStatus(documentDetailsPriv, statusSheet.getLastRow() + 1, workspaceName, "Setup");
            const documentDetailsPub = getDocDetails(key, `https://api.pandadoc.com/public/v1/documents/${doc.id}/details`);
            form.handleForm(documentDetailsPub, statusSheet.getLastRow());
            continue;
        }
    }
    return true;
};

const getDocDetails = (key, url, retries = 0) => {
    Logger.log("5. get Docs. Also should come after step 10 and after step 11");
    try {
        const createOptions = {
            'method': 'get',
            'headers': {
                'Authorization': key,
                'Content-Type': 'application/json;charset=UTF-8'
            }
        };
        const response = UrlFetchApp.fetch(url, createOptions);
        const responseJson = JSON.parse(response);
        return responseJson
    } catch (error) {
        if (retries > 1) {
            //logError(error); => need to handle errors
            console.log(error);
        }
        console.log(error);
        console.log(`Received error, retrying in 3 seconds... (attempt ${retries + 1} of 2)`);
        Utilities.sleep(3000);
        return getDocDetails(key, url, retries + 1);
    }
};

const pdIndex = {
    listDocuments: listDocs,
    processListDocResult: loopThroughListDocResult
};