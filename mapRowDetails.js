const additionalDocInfoMap = (data) => {
    try {
        const dataArray = data.map((obj) => {
            return [
                obj.status === "document.sent" ? obj.date_modified : (obj.date_sent ? obj.date_sent : ""),
                timeTo(obj.date_created, obj.date_completed ? obj.date_completed : ""), //Time Created to Completed
                timeTo(obj.status === "document.sent" ? obj.date_modified : (obj.date_sent ? obj.date_sent : ""), obj.date_completed ? obj.date_completed : ""), //Time Sent to Completed
                "", //Time Viewed to Completed
                timeTo(obj.date_created, obj.status === "document.sent" ? obj.date_modified : (obj.date_sent ? obj.date_sent : "")), //Time Created to Sent
                "", //Total time to approve
                timeTo(obj.status === "document.sent" ? obj.date_modified : (obj.date_sent ? obj.date_sent : ""), obj.status === 5 ? obj.date_status_changed : ""), //Time sent to first View
                obj.renewal ? obj.renewal.renewal_date : "",
                obj.date_expiration ? obj.date_expiration : ""
            ]
        });
        return dataArray
    } catch (error) {
        console.log(error);
        throw new Error("Script terminated: Error Adding New Row");
    }
};

const addDataToSheet = (data, workspaceName, privateAPIDetails) => {
    try {
        const dataArray = documentMap(data, workspaceName);
        let rowValues = null;

        if (privateAPIDetails) {
            rowValues = dataArray.map((innerArr, index) => {
                return innerArr.concat(privateAPIDetails[index]);
            });
        }

        const values = rowValues ? rowValues : dataArray;
        statusSheet.getRange(statusSheet.getLastRow() + 1, 1, values.length,values[0].length).setValues(values);
    } catch (error) {
        console.log(error);
        throw new Error("Script terminated: Error Adding details from Public API");
    }
};

const updateRowWhenStatusIsWrong = (data, workspaceName) => {
    try {
        const lastRow = statusSheet.getLastRow();
        const values = statusSheet.getRange(`A1:A${lastRow}`).getValues();
        let index = values.findIndex(data[0].id) + 1;
        console.log("Row Index: " + index);

        const rowValues = statusSheet.getRange(`A${index}:X${index}`).getValues()[0];
        console.log(rowValues);
        const dataArray = data.map((obj) => {
            return [
                obj.id,
                workspaceName ? workspaceName : "",
                obj.name,
                obj.date_created,
                getStatusFormattedText(obj.status),
                obj.status,
                obj.template && obj.template.id ? obj.template.id : "", 
                obj.created_by && obj.created_by.email ? obj.created_by.email : "", //Owner Email
                obj.linked_objects && obj.linked_objects.length > 0 ? formatProvider(obj.linked_objects[0].provider) : "",
                obj.grand_total ? obj.grand_total.currency : "",
                obj.grand_total ? obj.grand_total.amount : "",
                obj.date_completed ? obj.date_completed : "",
                obj.status === "document.viewed" ? obj.date_modified : "",
                obj.status === "document.waiting_approval" ? obj.date_modified : "",
                obj.status === "document.approved" ? obj.date_modified : "",
                rowValues[15].length > 0 ? rowValues[15] : "", //Date Sent
                obj.status === "document.completed" ? timeTo(obj.date_created, obj.date_modified) : "", //Time created to completed
                obj.status === "document.completed" && rowValues[15].length > 0 ? timeTo(rowValues[15], obj.date_modified): "", //Time Sent to completed
                obj.status === "document.completed" && rowValues[12].length > 0 ? timeTo(rowValues[12], obj.date_modified): "", //Time First View to completed
                rowValues[19].length > 0 ? rowValues[19] : (obj.status === "document.sent" ? timeTo(obj.date_created, rowValues[15]) : "") //Time Created to Sent
            ];
        });

        statusSheet.getRange(index, 1, 1, dataArray[0].length).setValues([dataArray[0]]);
    } catch (error) {
        console.log(error);
    }
};

const documentMap = (data, workspaceName) => {
    const dataArray = data.map((obj) => {
        return [
            obj.id,
            workspaceName ? workspaceName : "",
            obj.name,
            obj.date_created,
            getStatusFormattedText(obj.status),
            obj.status,
            obj.template && obj.template.id ? obj.template.id : "", //Template ID
            obj.created_by && obj.created_by.email ? obj.created_by.email : "", //Owner Email
            obj.linked_objects && obj.linked_objects.length > 0 ? formatProvider(obj.linked_objects[0].provider) : "",
            obj.grand_total ? obj.grand_total.currency : "",
            obj.grand_total ? obj.grand_total.amount : "",
            obj.date_completed ? obj.date_completed : "",
            obj.status === "document.viewed" ? obj.date_modified : "",
            obj.status === "document.waiting_approval" ? obj.date_modified : "",
            obj.status === "document.approved" ? obj.date_modified : ""
        ];
    });
    return dataArray
};

const documentMapUpdate = (data, row) => {
    const dataArray = data.map((obj) => {
        const rowValues = statusSheet.getRange(row, 1, 1, statusSheet.getLastColumn()).getValues()[0];
        return [
            getStatusFormattedText(obj.status),
            obj.status,
            rowValues[6], //Template ID
            rowValues[7], //Owner Email
            obj.linked_objects && obj.linked_objects.length > 0 ? formatProvider(obj.linked_objects[0].provider) : "",
            obj.grand_total ? obj.grand_total.currency : "",
            obj.grand_total ? obj.grand_total.amount : "",
            obj.date_completed ? obj.date_completed : rowValues[11],
            obj.status === "document.viewed" ? obj.date_modified : rowValues[12],
            obj.status === "document.waiting_approval" ? obj.date_modified : rowValues[13],
            obj.status === "document.approved" ? obj.date_modified : rowValues[14],
            obj.status === "document.sent" ? obj.date_modified : rowValues[15],
            timeTo(obj.date_created, obj.status === "document.completed" ? obj.date_modified : ""), //Time Created to Completed
            timeTo(rowValues[15], obj.status === "document.completed" ? obj.date_modified : rowValues[11]), //Time Sent to Completed
            timeTo(rowValues[12], obj.status === "document.completed" ? obj.date_modified : rowValues[11]), //Time viewed to Complete
            timeTo(obj.date_created, obj.status === "document.sent" ? obj.date_modified : rowValues[15]), //Time Created to Sent
            timeTo(rowValues[13], obj.status === "document.approved" ? obj.date_modified : rowValues[14]), //Total time to approve
            timeTo(rowValues[15], obj.status === "document.viewed" ? obj.date_modified : rowValues[12]), //Sent to first view
            rowValues[22], //Renewal Date
            obj.date_expiration ? obj.date_expiration : rowValues[23]
        ];
    });
    return dataArray
};

Array.prototype.findIndex = function (search) {
    for (let i = 0; i < this.length; i++) {
        if (this[i][0] === search) {
            return i;
        }
    }
    return -1;
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
};

const getStatusFormattedText = (status) => {
    switch (status) {
        case "document.draft":
            return "Draft";
        case "document.sent":
            return "Sent";
        case "document.completed":
            return "Completed";
        case "document.viewed":
            return "Viewed";
        case "document.waiting_approval":
            return "Waiting For Approval";
        case "document.approved":
            return "Approved";
        case "document.rejected":
            return "Rejected";
        case "document.waiting_pay":
            return "Waiting For Payment";
        case "document.paid":
            return "Paid";
        case "document.voided":
            return "Voided";
        case "document.declined":
            return "Declined";
        case "document.external_review":
            return "External Review";
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
};

const handleDocDetailsResponse = {
    privAPIResponseMap: additionalDocInfoMap,
    updateRowFromPubAPIResponse: addDataToSheet,
    wrongStatus: updateRowWhenStatusIsWrong,
    findRowIndex: Array.prototype.findIndex,
    documentMap: documentMap,
    documentMapUpdate: documentMapUpdate
};