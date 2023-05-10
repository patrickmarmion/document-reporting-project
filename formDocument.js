//Handle documents' linked objects

const logPublicAPIData = (data) => {

    const dataArray = data.map(obj => [
        obj.linked_objects.length > 0 ? formatProvider(obj.linked_objects[0].provider) : "",
        obj.grand_total.currency,
        obj.grand_total.amount
    ]);
    const index = headers[0].indexOf("Linked Object") + 1;
    statusSheet.getRange(statusSheet.getLastRow() - dataArray.length, index, dataArray.length, dataArray[0].length).setValues(dataArray);
};

const formatProvider = (provider) => {
    switch (provider) {
        case "salesforce-oauth2":
            return "Salesforce";
        case "hubspot":
            return "HubSpot";
        case "pandadoc-eform":
            return "PandaDoc Form";
        case "pipedrive":
            return "Pipedrive";
        case "salesforce-oauth2-sandbox":
            return "Salesforce Sandbox";
        default:
            return provider;
    }
}

const form = {
    handleForm: logPublicAPIData
};