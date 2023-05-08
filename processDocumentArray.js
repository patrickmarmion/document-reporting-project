//Title
const documentStatus = (data, row, workspaceName, event) => {
    Logger.log("6. In Doc status");
    try {
        const dataArray = data.map(obj => [
            obj.id,
            workspaceName,
            obj.name,
            obj.date_created,
            getStatusText(obj.status),
            getStatusFormattedText(obj.status),
            "", //Provider
            obj.date_sent,
            "", //Date Viewed
            obj.date_completed,
            "", //Date Sent For Approval
            "", //Date Approved
            timeTo(obj.date_created, obj.date_completed), //Time Created to Completed
            timeTo(obj.date_sent, obj.date_completed), //Time Sent to Completed
            "", //Time Viewed to Completed
            timeTo(obj.date_created, obj.date_sent) //Time Created to Sent
        ]);

        statusSheet.getRange(row, 1, dataArray.length, dataArray[0].length).setValues(dataArray);
    } catch (error) {
        console.log(error);
    }
};

const getStatusText = (status) => {
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

//NEED to complete
const getStatusFormattedText = (status) => {
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

const handleStatus = {
    documentStatus: documentStatus
};