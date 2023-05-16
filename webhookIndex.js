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
            errorHandler.logAPIError(workspaceProperty);
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

const handleWebook = {
    verifyWebhookSignature: verifyWebhookSignature,
    documentDeleted: docDeleted
};