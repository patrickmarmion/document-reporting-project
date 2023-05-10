const createTimeTriggers = () => {
    let incrementTrigger = ScriptApp.newTrigger('incrementCreateDate')
        .timeBased()
        .everyMinutes(1)
        .create();

    let continueTrigger = ScriptApp.newTrigger('continueFunction')
        .timeBased()
        .everyMinutes(5)
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
    Utilities.sleep(5000);
    const increment = Number(scriptProperties.getProperty('increment')) || 1;
    scriptProperties.setProperty('increment', increment + 1);

    statusSheet.appendRow(["Increment Index Executed HERE", scriptProperties.getProperty('createDate')]);
    scriptProperties.setProperty('stopFlag', 'false');
};


const continueFunction = () => {
    Logger.log('1. Continue Function');
    const increment = scriptProperties.getProperty('increment') || 0;
    if (increment >= 2) {
        let createDate = scriptProperties.getProperty('createDate') || "2021-01-01T01:01:01.000000Z";
        setup.setupIndex(createDate);
    }
};

const deleteSetupTriggers = () => {
    //Frist Delete the increment property, in case you have to run the script again
    scriptProperties.deleteProperty('increment');

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

const triggers = {
    createTriggers: createTimeTriggers,
    deleteTriggers: deleteSetupTriggers
};