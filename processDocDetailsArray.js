
const addNewRow = (data, row) => {
    Logger.log("6. In Doc status");
    try {
        const dataArray = data.map((obj) => {
            return [
                obj.date_sent,
                "", //Date Viewed
                "", //Date Sent For Approval
                "", //Date Approved
                timeTo(obj.date_created, obj.date_completed), //Time Created to Completed
                timeTo(obj.date_sent, obj.date_completed), //Time Sent to Completed
                "", //Time Viewed to Completed
                timeTo(obj.date_created, obj.date_sent), //Time Created to Sent
                "", //Total time to approve
                "", //Time sent to first View
                obj.renewal ? obj.renewal.renewal_date : "",
                obj.date_expiration
            ]
        });

        const index = headers[0].indexOf("Date Sent") + 1;
        statusSheet.getRange(row, index, dataArray.length, dataArray[0].length).setValues(dataArray);
    } catch (error) {
        console.log(error);
        throw new Error("Script terminated: Error Adding New Row");
    }
};

const updateRowWithPublicAPIResponse = (data, workspaceName) => {
    try {
        const dataArray = data.map((obj) => {
            return [
                obj.id,
                workspaceName ? workspaceName : "",
                obj.name,
                obj.date_created,
                getStatusFormattedText(obj.status),
                getStatusText(obj.status),
                obj.template && obj.template.id ? obj.template.id : "", //Template ID
                obj.created_by && obj.created_by.email ? obj.created_by.email : "", //Owner Email
                obj.linked_objects && obj.linked_objects.length > 0 ? formatProvider(obj.linked_objects[0].provider) : "",
                obj.grand_total ? obj.grand_total.currency : "", 
                obj.grand_total ? obj.grand_total.amount : "", 
                obj.date_completed ? obj.date_completed : "" 
            ];
        });
        const lastRow = statusSheet.getLastRow();
        let index;
        for (let i = lastRow +1; i >= 1; i--) {
            const value = statusSheet.getRange(`A${i}`).getValue();
            if (value) {
                index = i + 1;
                break
              }
        }
        statusSheet.getRange(index, 1, dataArray.length, dataArray[0].length).setValues(dataArray);
    } catch (error) {
        console.log(error);
        throw new Error("Script terminated: Error Adding details from Public API");
    }
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

const getStatusFormattedText = (status) => {
    switch (status) {
        case 0:
            return "Draft";
        case 1:
            return "Sent";
        case 2:
            return "Completed";
        case 5:
            return "Viewed";
        case 6:
            return "Waiting For Approval";
        case 7:
            return "Approved";
        case 8:
            return "Rejected";
        case 9:
            return "Waiting For Payment";
        case 10:
            return "Paid";
        case 11:
            return "Voided";
        case 12:
            return "Declined";
        case 13:
            return "External Review";
    }
};

const getStatusText = (status) => {
    switch (status) {
        case 0:
            return "document.draft";
        case 1:
            return "document.sent";
        case 2:
            return "document.completed";
        case 5:
            return "document.viewed";
        case 6:
            return "document.waiting_approval";
        case 7:
            return "document.approved";
        case 8:
            return "document.rejected";
        case 9:
            return "document.waiting_pay";
        case 10:
            return "document.paid";
        case 11:
            return "document.voided";
        case 12:
            return "document.declined";
        case 13:
            return "document.external_review";
    }
};

//Calculation: time between two events
const timeTo = (timeFirst, timeSecond) => {
    if (timeFirst && timeSecond) {
        const earlier = new Date(timeFirst);
        const later = new Date(timeSecond);
        const diffInSeconds = Math.floor((later - earlier) / 1000);
        const hms = (seconds) => {
            return [3600, 60]
                .reduceRight(
                    (p, b) => r => [Math.floor(r / b)].concat(p(r % b)),
                    r => [r]
                )(seconds)
                .map(a => a.toString().padStart(2, '0'))
                .join(':');
        }
        const duration = hms(diffInSeconds);
        return duration
    }
    return ""
}

const handleDocDetailsResponse = {
    addRowFromPrivAPIResponse: addNewRow,
    updateRowFromPubAPIResponse: updateRowWithPublicAPIResponse
};