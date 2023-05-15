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
    throw new Error("Webhook signature does not match. Key changed or payload has been modified!");
};

//Webhook Verification
const getGeneratedSignature = (input, secret) => {
    const signatureBytes = Utilities.computeHmacSha256Signature(input, secret);
    const generatedSignature = signatureBytes.reduce((str, byte) => {
        byte = (byte < 0 ? byte + 256 : byte).toString(16);
        return str + (byte.length === 1 ? "0" : "") + byte;
    }, "");
    return generatedSignature;
};

//Delete row when doc gets deleted
const docDeleted = (id) => {
    const rowIndex = handleDocDetailsResponse.findRowIndex(id) + 1;
    if (rowIndex === statusSheet.getLastRow() + 1) return;
    if (rowIndex < 1) {
        errorHandler.logAPIError(`Unable to find the ID of a deleted document ID of: ${id}`);
        return;
    };
    statusSheet.deleteRow(rowIndex);
};

const webhookIndex = {
    verifyWebhookSignature: verifyWebhookSignature,
    documentDeleted: docDeleted
}