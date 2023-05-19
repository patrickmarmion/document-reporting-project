const processWorkspaces = (date) => {
    Logger.log('2. loop through workspaces');

    for (const propertyKey of propertiesKeys) {
        if (!propertyKey.startsWith("token")) continue;

        const workspaceName = property.getValueFromScriptProperties(5, "name", propertyKey);
        let page = 1;

        while (true) {
            const {
                shouldPause,
                documentsFetched
            } = fetchAndProcessDocuments(properties[propertyKey], date, page, workspaceName, propertyKey);
            if (shouldPause) return;

            if (documentsFetched) break;

            page++;
        }
    }
    deleteOperations();
};

const fetchAndProcessDocuments = (token, date, page, workspaceName, propertyKey) => {
    let pauseForTime = false;
    const {
        length,
        docs
    } = pdIndex.listDocuments(`Bearer ${token}`, date, page);
    if (length === 0) {
        scriptProperties.deleteProperty(propertyKey)

        //Return createDate back to 2021. I need this for multiple workspaces
        scriptProperties.setProperty('createDate', "2021-01-01T01:01:01.000000Z");
        return {
            shouldPause: false,
            documentsFetched: true
        }
    };

    pauseForTime = triggers.terminateExecution("SetupPrivate", docs);
    if (pauseForTime) {
        return {
            shouldPause: true,
            documentsFetched: false
        }
    };

    //Insert 100 blank rows
    statusSheet.insertRows(statusSheet.getLastRow() + 1, 100);

    //temporary fix for throttling error
    Utilities.sleep(8000);

    pdIndex.processListDocResult(docs, `Bearer ${token}`, workspaceName, "");
    pauseForTime = triggers.terminateExecution("SetupPublic");
    if (pauseForTime) {
        return {
            shouldPause: true,
            documentsFetched: false
        }
    };

    return {
        shouldPause: false,
        documentsFetched: false
    }
};

const deleteOperations = () => {
    triggers.deleteTriggers();
    formatSheet.deleteDuplicateRows();
    formatSheet.sortByCreateDate();
    //Items older than 1 year deleted? Once I have backed them up to a database...    
};

const setup = {
    setupIndex: processWorkspaces
};