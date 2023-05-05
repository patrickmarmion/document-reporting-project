//Return Value from Script Properties


const getValueFromScriptProperties = (num, prop, property) => {
    const workspaceName = property.slice(num);
    const workspaceNameValue = propertiesKeys
        .filter((key) => key.includes(workspaceName) && key.includes(prop))
        .reduce((cur, key) => {
            return Object.assign(cur, {
                [key]: properties[key]
            });
        }, {});
    const valueArr = Object.values(workspaceNameValue);
    return valueArr[0];
};

const property = {
    getValueFromScriptProperties: getValueFromScriptProperties
}