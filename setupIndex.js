const loopThroughWorkspaces = (date) => {
    Logger.log('2. loop through workspaces');
    for (const key of propertiesKeys) {
        if (!key.startsWith("token")) continue;
        const workspaceName = property.getValueFromScriptProperties(5, "name", key);
        let page = 1;

        while (fetchAndProcessDocuments(properties[key], date, page, workspaceName)) {
            page++;
        }
    }
};

const fetchAndProcessDocuments = (key, date, page, workspaceName) => {
    const {
        length,
        docs
    } = pdIndex.listDocuments(`Bearer ${key}`, date, page);

    if (length === 0) {
        triggers.deleteTriggers();
        return false;
    }

    const docsFiltered = pdIndex.processListDocResult(docs, workspaceName, `Bearer ${key}`);
    if (!docsFiltered) {
        return false;
    }

    const forms = pdIndex.checkIfForm(docsFiltered, `Bearer ${key}`);
    return forms;
};


const setup = {
    setupIndex: loopThroughWorkspaces
};