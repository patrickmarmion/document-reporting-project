const loopThroughDocuments = (date) => {
    Logger.log('2. loop through doc');
    let page = 1;
    let foo = true;
    for (const key of propertiesKeys) {
        if (!key.startsWith("token")) continue;
        const workspaceName = property.getValueFromScriptProperties(5, "name", key);
        while (foo) {
            const {
                length,
                docs
            } = pdIndex.listDocuments(`Bearer ${properties[key]}`, date, page);
            page++
            if (length === 0) {
                triggers.deleteTriggers();
                break;
            }
            foo = pdIndex.processListDocResult(docs, workspaceName, `Bearer ${properties[key]}`);
        }
    }
};

const setup = {
    setupIndex: loopThroughDocuments
};