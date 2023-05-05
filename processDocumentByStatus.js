//Based on doc status & the event calls the relevant handler to write doc details to row

const documentStatus = (data, row, workspaceName, event, retries = 0) => {
    Logger.log("6. In Doc status")
    try {
        let rowValues;
        const columns = columnIndex.returnIndexOfHeader();
        if (event !== "Setup") {
            Logger.log("Not setup")
            rowValues = statusSheet.getRange(row, 1, 1, statusSheet.getLastColumn()).getValues(); //Do not need this in Setup...
        }
        let status = data.status;

        if (event !== "Recovery") {
            basicInfo(row, data, workspaceName);
            const statusMap = ['document.draft', 'document.sent', 'document.completed', '', '', 'document.viewed', 'document.waiting_approval', 'document.approved', 'document.rejected', 'document.waiting_pay', 'document.paid', 'document.voided', 'document.declined', 'document.external_review']
            if (typeof status === 'number') status = statusMap[data.status];
        }

        const handler = statusHandlers[status];
        handler(data, row, columns, rowValues, event);

    } catch (error) {
        console.log(error);
        /*
        if (retries > 2) {
            logError(error);
        }
        console.log(`Received error, retrying in 3 seconds... (attempt ${retries + 1} of 3)`);
        Utilities.sleep(3000);
        return documentStatus(data, row, workspaceName, event, retries + 1);
        */
    }
};

//For each status, I patch the first 4 columns with the document's ID, Workspace, Name, Date of Creation. Not called during Recovery when doc already exists.
const basicInfo = (row, data, workspaceName) => {
    Logger.log("8. Basic Info")
    statusSheet.getRange(row, 1, 1, 4).setValues([
        [data.id, workspaceName, data.name, data.date_created]
    ]);
};

//Sets the two status columns with both the formatted and unformatted status
const setStatusValue = (row, column, status, unformattedStatus) => {
    Logger.log("10. In set status - Final action in Draft Doc Process");
    statusSheet.getRange(row, column.status).setValue(status);
    statusSheet.getRange(row, column.statusUnformat).setValue(unformattedStatus);
};

//A map of of the different statuses and their corresponding handling function
const statusHandlers = {
    "document.draft": draft.draftStatus,
    /*
    "document.sent": handleDocumentSent,
    "document.completed": handleDocumentCompleted,
    "document.viewed": handleDocumentViewed,
    "document.waiting_approval": handleDocumentWaitingApproval,
    "document.rejected": handleDocumentRejected,
    "document.approved": handleDocumentApproved,
    "document.waiting_pay": handleDocumentWaitingPay,
    "document.paid": handleDocumentPaid,
    "document.voided": handleDocumentVoid,
    "document.declined": handleDocumentDeclined,
    "document.external_review": handleDocumentExternalReview
    */
};

const handleStatus = {
    documentStatus: documentStatus,
    setStatus: setStatusValue
};