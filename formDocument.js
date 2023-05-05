//Handle documents that were created from a form

const logForm = (data, row) => {
    Logger.log("11. Check if doc was a form")
    if (data.linked_objects.length > 0) {
        const firstObject = data.linked_objects[0];
        if (firstObject.provider === "pandadoc-eform") {
            const columns = columnHeaders(headers);
            statusSheet.getRange(row, columns.form).setValue("True");
        }
    }
    return
};


const form = {
    handleForm: logForm 
};