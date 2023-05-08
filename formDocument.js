//Handle documents that were created from a form

const logForm = (data) => {
    //Logger.log("11. Check if doc was a form")

    const dataArray = data.map(obj => [
        obj.linked_objects.length > 0 ? obj.linked_objects[0].provider : ''
    ]);

    statusSheet.getRange(statusSheet.getLastRow() - dataArray.length, 7, dataArray.length, 1).setValues(dataArray);
};


const form = {
    handleForm: logForm
};