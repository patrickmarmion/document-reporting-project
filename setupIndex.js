const loopThroughWorkspaces = (date) => {
    Logger.log('2. loop through workspaces');
    let pauseForTime = false;
    for (const key of propertiesKeys) {
        if (!key.startsWith("token")) continue;
        const workspaceName = property.getValueFromScriptProperties(5, "name", key);
        let page = 1;

        while (true) {
            const {
                shouldPause,
                documentsFetched
            } = fetchAndProcessDocuments(properties[key], date, page, workspaceName, key);
            if (shouldPause) {
                pauseForTime = true;
                break;
            }
            if (documentsFetched) break;

            page++;
        }
        if (pauseForTime) break;
    }
    if (!pauseForTime) {
        triggers.deleteTriggers();
        deleteDuplicateRowsById();

        //Items older than 1 year deleted? Once I have back them up to a database...    
    }
};

const fetchAndProcessDocuments = (key, date, page, workspaceName, propertyKey) => {
    Logger.log('3 Fetch and process docs');
    let pauseForTime = false;
    const {
        length,
        docs
    } = pdIndex.listDocuments(`Bearer ${key}`, date, page);
    console.log(workspaceName + ": " + length);
    if (length === 0) {
        //Delete token in script properties.
        scriptProperties.deleteProperty(propertyKey)

        //Return createDate back to 2021.
        scriptProperties.setProperty('createDate', "2021-01-01T01:01:01.000000Z");
        return {
            shouldPause: false,
            documentsFetched: true
        }
    };

    pauseForTime = triggers.terminateExecution("SetupPrivate", docs);
    if (pauseForTime)  {
        return {
            shouldPause: true,
            documentsFetched: false
        }
    };

    const docsFiltered = pdIndex.processListDocResult(docs, `Bearer ${key}`);

    //Insert 100 blank rows
    statusSheet.insertRows(statusSheet.getLastRow() + 1, 100);

    //temporary fix for throttling error
    Utilities.sleep(2000);

    pdIndex.processListDocResultPublicDetails(docsFiltered, workspaceName, `Bearer ${key}`);
    pauseForTime = triggers.terminateExecution("SetupPublic");
    if (pauseForTime)  {
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

const deleteDuplicateRowsById = () => {
    let data = statusSheet.getDataRange().getValues();
    let newData = [];
    let seen = {};

    for (let i = 0; i < data.length; i++) {
        let row = data[i];
        let value = row[0];
        if (value && !seen[value]) {
            newData.push(row);
            seen[value] = true;
        } else {
            statusSheet.deleteRow(i + 1);
        }
    }
    statusSheet.getDataRange().clearContent();
    statusSheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
};

const setup = {
    setupIndex: loopThroughWorkspaces
};