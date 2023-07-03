/**
 * Creates time-based triggers. 
 * The stopExecutionTrigger prevents the script from running too long and receiving a time-out error.
 * Sets unique trigger IDs as script properties.
 * @returns {void}
 */
const createTimeTriggers = () => {
    let stopExecutionTrigger = ScriptApp.newTrigger('stopExecution')
        .timeBased()
        .everyMinutes(5)
        .create();

    let runSetupTrigger = ScriptApp.newTrigger('runSetup')
        .timeBased()
        .everyMinutes(10)
        .create();

    let recoveryCheck = ScriptApp.newTrigger('runRecovery')
        .timeBased()
        .onWeekDay(ScriptApp.WeekDay.MONDAY)
        .atHour(6)
        .create();

    const stopExecutionTriggerId = stopExecutionTrigger.getUniqueId();
    const runSetupTriggerId = runSetupTrigger.getUniqueId();
    scriptProperties.setProperty("stopExecutionTriggerID", stopExecutionTriggerId);
    scriptProperties.setProperty("runSetupTriggerID", runSetupTriggerId);
    Logger.log('Triggers created successfully.');
};

/**
 * Sets the "stopFlag" property before and after sleeping.
 * @returns {void}
 */
const stopExecution = () => {
    scriptProperties.setProperty('stopFlag', 'true');
    Utilities.sleep(14000);
    scriptProperties.setProperty('stopFlag', 'false');
};

/**
 * Calls the "setupIndex" function from the "setup" module, passing the "createDate" property as an argument.
 * @returns {void}
 */
const runSetup = () => {
    const createDate = scriptProperties.getProperty('createDate');
    setup.processWorkspaces(createDate);
};

const runRecovery = () => {
    recovery.recoveryIndex();
}

const deleteSetupTriggers = () => {

    const projectTriggers = ScriptApp.getProjectTriggers();
    const stopExecutionTrig = scriptProperties.getProperty('stopExecutionTriggerID');
    const runSetupTrig = scriptProperties.getProperty('runSetupTriggerID');

    // Iterate over the project triggers
    for (let i = 0; i < projectTriggers.length; i++) {
        const trigger = projectTriggers[i];
        const triggerId = trigger.getUniqueId();

        if (stopExecutionTrig.includes(triggerId) || runSetupTrig.includes(triggerId)) {
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