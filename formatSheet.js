const sortSheetByCreateDate = () => {
    const range = statusSheet.getDataRange();
    const data = range.getValues();
    const sortColumn = 4;

    // Convert strings to dates
    for (let i = 1; i < data.length; i++) {
        const dateString = data[i][sortColumn - 1];
        data[i][sortColumn - 1] = new Date(dateString);
    }

    // Sort the data by the specified column
    data.sort(function (a, b) {
        return a[sortColumn - 1] - b[sortColumn - 1];
    });

    // Write the sorted data back to the sheet
    range.setValues(data);
};

const deleteDuplicateRowsById = () => {
    let data = statusSheet.getDataRange().getValues();
    let newData = [];
    let seen = {};

    for (let i = 0; i < data.length; i++) {
        let row = data[i];
        let value = row[0];
        if (value && !seen[value]) {
            newData.push(row);
            seen[value] = true;
        } else {
            statusSheet.deleteRow(i + 1);
        }
    }
    statusSheet.getDataRange().clearContent();
    statusSheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
};

const formatSheet = {
    sortRowsByCreateDate: sortSheetByCreateDate,
    deleteDuplicateRows: deleteDuplicateRowsById
}