const runAddOn = () => {
    Logger.log("start runAddOn");
    // Get active sheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const range = sheet.getDataRange();

    const mergedArr = convertToArray(range, 1);

    Logger.log("mergedArr");
    Logger.log(mergedArr);
}

const convertToArray = (range, matchIndex) => {
    Logger.log("start convertToArray");

    const origValuesArr = range.getValues();

    Logger.log("origValuesArr");
    Logger.log(origValuesArr);

    const cache = {};

    origValuesArr.forEach(row => {
        const matchValue = row[matchIndex];
        const mergeValue = row[0];
        if (cache.hasOwnProperty(matchValue)) {
            Logger.log("match found");
            Logger.log(cache[matchValue]);
            cache[matchValue] = cache[matchValue].concat("\n", mergeValue);
            Logger.log("merging");
            Logger.log(cache[matchValue]);
        } else {
            cache[matchValue] = `${mergeValue}`;
        }
    });

    const mergedValuesArr = [];

    for (const property in cache) {
        const row = [cache[property], property];
        mergedValuesArr.push(row);
    }

    return mergedValuesArr;
}
