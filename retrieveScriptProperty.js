//Return Value from Script Properties

const getValueFromScriptProperties = (num, prop, property) => {
    const strSliced = property.slice(num);
    const value = propertiesKeys
        .filter((key) => key.includes(strSliced) && key.includes(prop))
        .reduce((cur, key) => {
            return Object.assign(cur, {
                [key]: properties[key]
            });
        }, {});
    const valueArr = Object.values(value);
    return valueArr[0];
};

const property = {
    getValueFromScriptProperties: getValueFromScriptProperties
}