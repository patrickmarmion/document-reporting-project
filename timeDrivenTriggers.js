const createTimeTriggers = () => {
    let incrementTrigger = ScriptApp.newTrigger('incrementCreateDate')
        .timeBased()
        .everyMinutes(10)
        .create();

    let continueTrigger = ScriptApp.newTrigger('continueFunction')
        .timeBased()
        .everyMinutes(15)
        .create();

    const incrementTriggerId = incrementTrigger.getUniqueId();
    const continueTriggerId = continueTrigger.getUniqueId();
    scriptProperties.setProperty("incrementTriggerID", incrementTriggerId);
    scriptProperties.setProperty("continueTriggerID", continueTriggerId);
    Logger.log('Triggers created successfully.');

    //Below need to create recovery trigger to run weekly.
};

const incrementCreateDate = () => {
    scriptProperties.setProperty('stopFlag', 'true');
    Utilities.sleep(12000);

    scriptProperties.setProperty('stopFlag', 'false');
};

const continueFunction = () => {
    Logger.log('1. Continue Function');
    const createDate = scriptProperties.getProperty('createDate');
    setup.setupIndex(createDate);
};

const deleteSetupTriggers = () => {

    const projectTriggers = ScriptApp.getProjectTriggers();
    const incrementTrigg = scriptProperties.getProperty('incrementTriggerID');
    const continueTrigg = scriptProperties.getProperty('continueTriggerID');

    // Iterate over the project triggers
    for (let i = 0; i < projectTriggers.length; i++) {
        const trigger = projectTriggers[i];
        const triggerId = trigger.getUniqueId();

        if (incrementTrigg.includes(triggerId) || continueTrigg.includes(triggerId)) {
            ScriptApp.deleteTrigger(trigger);
            Logger.log('Trigger deleted successfully.');
        }
    }
};

const shouldPause = (event, docs) => {
    try {
        const stopFlag = scriptProperties.getProperty('stopFlag');
        if (stopFlag === 'true') {
            Logger.log("Paused for time");
            switch (event) {
                case "Recovery":
                    return true;
                case "SetupPrivate":
                    scriptProperties.setProperty('createDate', docs[0].date_created);
                    return true;
                case "SetupPublic":
                    const lastRow = statusSheet.getLastRow();
                    const createDate = statusSheet.getRange(`D1:D${lastRow}`).getValues().reverse().find(([value]) => value !== '' && value !== "Date Created")?.[0];
                    scriptProperties.setProperty('createDate', createDate);
                    return true;
                default:
                    break;
            }
        }
        return false;
    } catch (error) {
        console.log(error)
    }

};

const triggers = {
    createTriggers: createTimeTriggers,
    deleteTriggers: deleteSetupTriggers,
    terminateExecution: shouldPause
};