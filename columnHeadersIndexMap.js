const columnHeaders = () => {
    Logger.log("7. Column Headers")
    return {
        status: Math.floor(headers[0].indexOf("Status") + 1),
        dateSent: Math.floor(headers[0].indexOf("Date Sent") + 1),
        dateViewed: Math.floor(headers[0].indexOf("Date Viewed") + 1),
        dateCompleted: Math.floor(headers[0].indexOf("Date Completed") + 1),
        dateSentForApproval: Math.floor(headers[0].indexOf("Date Sent For Approval") + 1),
        dateApproved: Math.floor(headers[0].indexOf("Date Approved") + 1),
        timeCreatedToCompleted: Math.floor(headers[0].indexOf("Time Created to Completed (HH:MM:SS)") + 1),
        timeSentToCompleted: Math.floor(headers[0].indexOf("Time Sent to Completed (HH:MM:SS)") + 1),
        timeViewedToCompleted: Math.floor(headers[0].indexOf("Time Viewed to Completed (HH:MM:SS)") + 1),
        timeCreatedToSent: Math.floor(headers[0].indexOf("Time Created to Sent (HH:MM:SS)") + 1),
        timeToApproveDoc: Math.floor(headers[0].indexOf("Total Time to Approve (HH:MM:SS)") + 1),
        timeSentToViewed: Math.floor(headers[0].indexOf("Time Sent to First View (HH:MM:SS)") + 1),
        id: Math.floor(headers[0].indexOf("ID") + 1),
        workspace: Math.floor(headers[0].indexOf("Workspace Name") + 1),
        name: Math.floor(headers[0].indexOf("Document Name") + 1),
        dateCreated: Math.floor(headers[0].indexOf("Date Created") + 1),
        statusUnformat: Math.floor(headers[0].indexOf("Status Unformatted") + 1),
        form: Math.floor(headers[0].indexOf("Form") + 1)
    }
};

const columnIndex = {
    returnIndexOfHeader: columnHeaders
}