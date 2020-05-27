const utils = require('./utils');

const mapCollectErrorsSheetWithCreatorAndProperties =                (collectors, sheet, creator, ...properties) => mapCollectErrorsSheetByStartWithCreatorAndProperties(collectors, sheet, 0, creator, ...properties);
const mapCollectErrorsSheetByStartWithCreatorAndProperties =         (collectors, sheet, startIndex, creator, ...properties) => mapCollectErrorsSheetByRangeWithCreatorAndProperties(collectors, sheet, startIndex, sheet.rows + 1 - startIndex, creator, ...properties);
const mapCollectErrorsSheetByRangeWithCreatorAndProperties =         (collectors, sheet, startIndex, count, creator, ...properties) => {
    let result = [];
    for (let r = startIndex; r < startIndex + count; r++) {
        let row = sheet.getRow(r);
        if (row != null && row.columns > 0) {
            let rowError = 0;
            let obj = creator();
            for (let property of properties) {
                try {
                    property.set(obj, row, r);
                } catch (e) {
                    collectors.push(e);
                    rowError++;
                }
            }
            if (!rowError) {
                result.push(obj);
            }
        }
    }
    return result;
}

const mapCollectErrorsExcelWithCreatorAndProperties =                (collectors, sheetIndex, excel, creator, ...properties) => mapCollectErrorsSheetWithCreatorAndProperties(collectors, utils.getSheet(excel, sheetIndex), creator, ...properties);
const mapCollectErrorsExcelByStartWithCreatorAndProperties =         (collectors, sheetIndex, excel, startIndex, creator, ...properties) => mapCollectErrorsSheetByStartWithCreatorAndProperties(collectors, utils.getSheet(excel, sheetIndex), startIndex, creator, ...properties);
const mapCollectErrorsExcelByRangeWithCreatorAndProperties =         (collectors, sheetIndex, excel, startIndex, count, creator, ...properties) => mapCollectErrorsSheetByRangeWithCreatorAndProperties(collectors, utils.getSheet(excel, sheetIndex), startIndex, count, creator, ...properties);
const mapCollectErrorsFirstSheetWithCreatorAndProperties =           (collectors, excel, creator, ...properties) => mapCollectErrorsExcelWithCreatorAndProperties(collectors, 0, excel, creator, ...properties);
const mapCollectErrorsFirstSheetByStartWithCreatorAndProperties =    (collectors, excel, startIndex, creator, ...properties) => mapCollectErrorsExcelByStartWithCreatorAndProperties(collectors, 0, excel, startIndex, creator, ...properties);
const mapCollectErrorsFirstSheetByRangeWithCreatorAndProperties =    (collectors, excel, startIndex, count, creator, ...properties) => mapCollectErrorsExcelByRangeWithCreatorAndProperties(collectors, 0, excel, startIndex, count, creator, ...properties);

const mapIgnoreErrorsSheetWithCreatorAndProperties =                (sheet, creator, ...properties) => mapCollectErrorsSheetByStartWithCreatorAndProperties([], sheet, 0, creator, ...properties);
const mapIgnoreErrorsSheetByStartWithCreatorAndProperties =         (sheet, startIndex, creator, ...properties) => mapCollectErrorsSheetByRangeWithCreatorAndProperties([], sheet, startIndex, sheet.rows + 1 - startIndex, creator, ...properties);
const mapIgnoreErrorsSheetByRangeWithCreatorAndProperties =         (sheet, startIndex, count, creator, ...properties) => mapCollectErrorsSheetByRangeWithCreatorAndProperties([], sheet, startIndex, count, creator, ...properties);
const mapIgnoreErrorsExcelWithCreatorAndProperties =                (sheetIndex, excel, creator, ...properties) => mapCollectErrorsSheetWithCreatorAndProperties([], utils.getSheet(excel, sheetIndex), creator, ...properties);
const mapIgnoreErrorsExcelByStartWithCreatorAndProperties =         (sheetIndex, excel, startIndex, creator, ...properties) => mapCollectErrorsSheetByStartWithCreatorAndProperties([], utils.getSheet(excel, sheetIndex), startIndex, creator, ...properties);
const mapIgnoreErrorsExcelByRangeWithCreatorAndProperties =         (sheetIndex, excel, startIndex, count, creator, ...properties) => mapCollectErrorsSheetByRangeWithCreatorAndProperties([], utils.getSheet(excel, sheetIndex), startIndex, count, creator, ...properties);
const mapIgnoreErrorsFirstSheetWithCreatorAndProperties =           (excel, creator, ...properties) => mapCollectErrorsExcelWithCreatorAndProperties([], 0, excel, creator, ...properties);
const mapIgnoreErrorsFirstSheetByStartWithCreatorAndProperties =    (excel, startIndex, creator, ...properties) => mapCollectErrorsExcelByStartWithCreatorAndProperties([], 0, excel, startIndex, creator, ...properties);
const mapIgnoreErrorsFirstSheetByRangeWithCreatorAndProperties =    (excel, startIndex, count, creator, ...properties) => mapCollectErrorsExcelByRangeWithCreatorAndProperties([], 0, excel, startIndex, count, creator, ...properties);


const mapIgnoreErrors = (...args) => {
    if (args[0] instanceof Array) {
        if (args[1] instanceof utils.Sheet) {
            let lastFunctionIndex = args.slice(2).findIndex(arg => utils.isFunction(arg)) + 2;
            if (lastFunctionIndex === 2) {
                return mapCollectErrorsSheetWithCreatorAndProperties(...args);
            } else if (lastFunctionIndex === 3) {
                return mapCollectErrorsSheetByStartWithCreatorAndProperties(...args);
            } else if (lastFunctionIndex === 4) {
                return mapCollectErrorsSheetByRangeWithCreatorAndProperties(...args);
            }
        } else if (args[1] instanceof utils.Excel || utils.isString(args[1]) || utils.isArrayBuffer(args[1])) {
            let lastFunctionIndex = args.slice(2).findIndex(arg => utils.isFunction(arg)) + 2;
            if (lastFunctionIndex === 2) {
                return mapCollectErrorsFirstSheetWithCreatorAndProperties(args[0], utils.getExcel(args[1]), ...args.slice(2));
            } else if (lastFunctionIndex === 3) {
                return mapCollectErrorsFirstSheetByStartWithCreatorAndProperties(args[0], utils.getExcel(args[1]), ...args.slice(2));
            } else if (lastFunctionIndex === 4) {
                return mapCollectErrorsFirstSheetByRangeWithCreatorAndProperties(args[0], utils.getExcel(args[1]), ...args.slice(2));
            }
        } else {
            let lastFunctionIndex = args.slice(3).findIndex(arg => utils.isFunction(arg)) + 3;
            if (lastFunctionIndex === 3) {
                return mapCollectErrorsExcelWithCreatorAndProperties(args[0], args[1], utils.getExcel(args[2]), ...args.slice(3));
            } else if (lastFunctionIndex === 4) {
                return mapCollectErrorsExcelByStartWithCreatorAndProperties(args[0], args[1], utils.getExcel(args[2]), ...args.slice(3));
            } else if (lastFunctionIndex === 5) {
                return mapCollectErrorsExcelByRangeWithCreatorAndProperties(args[0], args[1], utils.getExcel(args[2]), ...args.slice(3));
            }
        }
    } else if (args[0] instanceof utils.Sheet) {
        let lastFunctionIndex = args.slice(1).findIndex(arg => utils.isFunction(arg)) + 1;
        if (lastFunctionIndex === 1) {
            return mapIgnoreErrorsSheetWithCreatorAndProperties(...args);
        } else if (lastFunctionIndex === 2) {
            return mapIgnoreErrorsSheetByStartWithCreatorAndProperties(...args);
        } else if (lastFunctionIndex === 3) {
            return mapIgnoreErrorsSheetByRangeWithCreatorAndProperties(...args);
        }
    } else if (args[0] instanceof utils.Excel || utils.isString(args[0]) || utils.isArrayBuffer(args[0])) {
        let lastFunctionIndex = args.slice(1).findIndex(arg => utils.isFunction(arg)) + 1;
        if (lastFunctionIndex === 1) {
            return mapIgnoreErrorsFirstSheetWithCreatorAndProperties(utils.getExcel(args[0]), ...args.slice(1));
        } else if (lastFunctionIndex === 2) {
            return mapIgnoreErrorsFirstSheetByStartWithCreatorAndProperties(utils.getExcel(args[0]), ...args.slice(1));
        } else if (lastFunctionIndex === 3) {
            return mapIgnoreErrorsFirstSheetByRangeWithCreatorAndProperties(utils.getExcel(args[0]), ...args.slice(1));
        }
    } else {
        let lastFunctionIndex = args.slice(2).findIndex(arg => utils.isFunction(arg)) + 2;
        if (lastFunctionIndex === 2) {
            return mapIgnoreErrorsExcelWithCreatorAndProperties(args[0], utils.getExcel(args[1]), ...args.slice(2));
        } else if (lastFunctionIndex === 3) {
            return mapIgnoreErrorsExcelByStartWithCreatorAndProperties(args[0], utils.getExcel(args[1]), ...args.slice(2));
        } else if (lastFunctionIndex === 4) {
            return mapIgnoreErrorsExcelByRangeWithCreatorAndProperties(args[0], utils.getExcel(args[1]), ...args.slice(2));
        }
    }
    throw new utils.ExcelError('arguments error');
}

module.exports = mapIgnoreErrors;
