//Verification and determine workspace from Script Properties
const verifyWebhookSignature = (e) => {
    const {
        signature
    } = e.parameter;
    const input = e.postData.contents;

    for (const key of propertiesKeys) {
        if (!key.startsWith("sharedKey")) continue;

        const generatedSignature = getGeneratedSignature(input, properties[key]);

        if (signature === generatedSignature) {
            const workspaceProperty = property.getValueFromScriptProperties(9, "name", key);
            return workspaceProperty;
        }
    }
};

//Webhook Verification
const getGeneratedSignature = (input, secret) => {
    try {
        const signatureBytes = Utilities.computeHmacSha256Signature(input, secret);
        const generatedSignature = signatureBytes.reduce((str, byte) => {
            byte = (byte < 0 ? byte + 256 : byte).toString(16);
            return str + (byte.length === 1 ? "0" : "") + byte;
        }, "");
        return generatedSignature;
    } catch (error) {
        errorHandler.logAPIError("Error: Couldn't verify shared key");
    }
};

//Delete row when doc gets deleted
const docDeleted = (id, rowIndex) => {
    if (rowIndex === statusSheet.getLastRow() + 1) {
        errorHandler.logAPIError(`Unable to find the ID of a deleted document ID of: ${id}`);
        return;
    };
    statusSheet.deleteRow(rowIndex);
};

const webhookAddRow = (data, workspaceName, row) => {
    const docDetailsArray = handleDocDetailsResponse.documentMap(data, workspaceName, row);
    const docTimings = additionalDocInfoMap(data);

    const values = docDetailsArray[0].concat(docTimings[0]);
    statusSheet.getRange(row, 1, 1, values.length).setValues([values]);
};

const webhookUpdateRow = (data, row) => {
    const docDetailsArray = handleDocDetailsResponse.documentMapUpdate(data, row);
    statusSheet.getRange(row, 5, docDetailsArray.length, docDetailsArray[0].length).setValues(docDetailsArray);
};

const webhookRecipientCompleted = (data, row) => {
    if (data[0].status === "document.completed") {
        const docDetailsArray = handleDocDetailsResponse.documentMapUpdate(data, row);
        statusSheet.getRange(row, 5, docDetailsArray.length, docDetailsArray[0].length).setValues(docDetailsArray);
    };

    const rowValues = statusSheet.getRange(row, 1, 1, statusSheet.getLastColumn()).getValues()[0];
    for (let i = 24; i < rowValues.length; i++) {
        if (rowValues[i] === '') {
            rowValues.splice(i, 1);
            i--; 
        }
    };
    const recipientDetails = data.map((obj) => {
        return [
            obj.action_by.email,
            obj.action_date
        ]
    });
    
    
    const values = rowValues.concat(recipientDetails[0]);
    statusSheet.getRange(row, 1, 1, values.length).setValues([values]);
};

const handleWebook = {
    verifyWebhookSignature: verifyWebhookSignature,
    documentDeleted: docDeleted,
    webhookAddRow: webhookAddRow,
    webhookUpdateRow: webhookUpdateRow,
    webhookRecipientCompleted: webhookRecipientCompleted
};