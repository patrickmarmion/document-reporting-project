//Handle documents' linked objects

const logForm = (data) => {
    //Logger.log("11. Check if doc was a form")

    const dataArray = data.map(obj => [
        obj.linked_objects.length > 0 ? formatProvider(obj.linked_objects[0].provider) : ''
    ]);

    statusSheet.getRange(statusSheet.getLastRow() - dataArray.length, 7, dataArray.length, 1).setValues(dataArray);
};

const formatProvider = (provider) => {
    switch (provider) {
        case "salesforce-oauth2":
            return "Salesforce";
        case "hubspot":
            return "HubSpot";
        case "pandadoc-eform":
            return "PandaDoc Form";
        default:
            return provider;
    }
}

const form = {
    handleForm: logForm
};