/**
 * Based on the Script Properties this function loops through each workspace's Bearer Token.
 * @param {string} createDate - The document create date parameter, the earliest this can be is 01-01-2021. If the script has partially run but was stopped for time, the createDate will be the last added row.
 * @returns {void}
 */
const processWorkspaces = (createDate) => {
    const properties = scriptProperties.getProperties();
    for (const propertyKey of propertiesKeys) {
        if (!propertyKey.startsWith("token")) continue;

        const workspaceName = property.getValueFromScriptProperties(5, "name", propertyKey, properties);
        Logger.log(`Checking workspace ${workspaceName}`);

        let page = 1;

        while (true) {
            const {
                shouldPause,
                documentsFetched
            } = fetchAndProcessDocuments(properties[propertyKey], createDate, page, workspaceName, propertyKey);
            if (shouldPause) return;
            if (documentsFetched) break;

            page++;
        }
    }
    deleteOperations();
};

/**
 * Calls the List Doc endpoint, up to 100 documents returned.
 * If nothing is returned, loop is finished: deletes the token, returns the createDate to 01-01-2021, ready for the next workspace.
 * If docs are returned, adds rows to sheet and passes the listed docs array to get their details.
 * Twice checks if the script is approaching its maximum execution time.
 * @param {string} token - The Bearer token, with access to the Private API.
 * @param {string} createDate - The createDate parameter, used in the List Docs endpoint.
 * @param {number} page - The page parameter, incremented in the processWorkspace function and used in the List Docs endpoint.
 * @param {string} workspaceName - The workspace name parameter.
 * @param {string} propertyKey - The property key is from the Script Properties. It is the key corresponding to the token value, I pass this so I can delete the token to prevent looping through it mulitple times. 
 * @returns {Object} - Return two parameters, shouldPause & documentsFetched. shouldPause kills the script from running, documentsFetched increments the workspace loop.
 */
const fetchAndProcessDocuments = (token, createDate, page, workspaceName, propertyKey) => {
   let pauseForTime = false;
    const {
        length,
        docs
    } = pdIndex.listDocuments(`Bearer ${token}`, createDate, page);
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

    statusSheet.insertRows(statusSheet.getLastRow() + 1, 100);
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
    formatSheet.sortRowsByCreateDate();
    //Items older than 1 year deleted? Once I have backed them up to a database...    
};

const setup = {
    processWorkspaces: processWorkspaces
};