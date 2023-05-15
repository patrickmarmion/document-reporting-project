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

const formatSheet = {
    sortByCreateDate: sortSheetByCreateDate
}