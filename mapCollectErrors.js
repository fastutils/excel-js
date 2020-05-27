const utils = require('./utils');

const mapCollectErrorsSheetWithCreatorAndProperties =                (sheet, creator, ...properties) => mapCollectErrorsSheetByStartWithCreatorAndProperties(sheet, 0, creator, ...properties);
const mapCollectErrorsSheetByStartWithCreatorAndProperties =         (sheet, startIndex, creator, ...properties) => mapCollectErrorsSheetByRangeWithCreatorAndProperties(sheet, startIndex, sheet.rows + 1 - startIndex, creator, ...properties);
const mapCollectErrorsSheetByRangeWithCreatorAndProperties =         (sheet, startIndex, count, creator, ...properties) => {
    let result = [];
    let errors = [];
    for (let r = startIndex; r < startIndex + count; r++) {
        let row = sheet.getRow(r);
        if (row != null && row.columns > 0) {
            let obj = creator();
            for (let property of properties) {
                try {
                    property.set(obj, row, r);
                } catch (e) {
                    errors.push(e);
                }
            }
            result.push(obj);
        }
    }
    if (errors.length) {
        throw new utils.ExcelMultiValueFormatError(errors);
    } else {
        return result;
    }
}

const mapCollectErrorsExcelWithCreatorAndProperties =                (sheetIndex, excel, creator, ...properties) => mapCollectErrorsSheetWithCreatorAndProperties(utils.getSheet(excel, sheetIndex), creator, ...properties);
const mapCollectErrorsExcelByStartWithCreatorAndProperties =         (sheetIndex, excel, startIndex, creator, ...properties) => mapCollectErrorsSheetByStartWithCreatorAndProperties(utils.getSheet(excel, sheetIndex), startIndex, creator, ...properties);
const mapCollectErrorsExcelByRangeWithCreatorAndProperties =         (sheetIndex, excel, startIndex, count, creator, ...properties) => mapCollectErrorsSheetByRangeWithCreatorAndProperties(utils.getSheet(excel, sheetIndex), startIndex, count, creator, ...properties);
const mapCollectErrorsFirstSheetWithCreatorAndProperties =           (excel, creator, ...properties) => mapCollectErrorsExcelWithCreatorAndProperties(0, excel, creator, ...properties);
const mapCollectErrorsFirstSheetByStartWithCreatorAndProperties =    (excel, startIndex, creator, ...properties) => mapCollectErrorsExcelByStartWithCreatorAndProperties(0, excel, startIndex, creator, ...properties);
const mapCollectErrorsFirstSheetByRangeWithCreatorAndProperties =    (excel, startIndex, count, creator, ...properties) => mapCollectErrorsExcelByRangeWithCreatorAndProperties(0, excel, startIndex, count, creator, ...properties);


const mapCollectErrors = (...args) => {
    if (args[0] instanceof utils.Sheet) {
        let lastFunctionIndex = args.slice(1).findIndex(arg => utils.isFunction(arg)) + 1;
        if (lastFunctionIndex === 1) {
            return mapCollectErrorsSheetWithCreatorAndProperties(...args);
        } else if (lastFunctionIndex === 2) {
            return mapCollectErrorsSheetByStartWithCreatorAndProperties(...args);
        } else if (lastFunctionIndex === 3) {
            return mapCollectErrorsSheetByRangeWithCreatorAndProperties(...args);
        }
    } else if (args[0] instanceof utils.Excel || utils.isString(args[0]) || utils.isArrayBuffer(args[0])) {
        let lastFunctionIndex = args.slice(1).findIndex(arg => utils.isFunction(arg)) + 1;
        if (lastFunctionIndex === 1) {
            return mapCollectErrorsFirstSheetWithCreatorAndProperties(utils.getExcel(args[0]), ...args.slice(1));
        } else if (lastFunctionIndex === 2) {
            return mapCollectErrorsFirstSheetByStartWithCreatorAndProperties(utils.getExcel(args[0]), ...args.slice(1));
        } else if (lastFunctionIndex === 3) {
            return mapCollectErrorsFirstSheetByRangeWithCreatorAndProperties(utils.getExcel(args[0]), ...args.slice(1));
        }
    } else {
        let lastFunctionIndex = args.slice(2).findIndex(arg => utils.isFunction(arg)) + 2;
        if (lastFunctionIndex === 2) {
            return mapCollectErrorsExcelWithCreatorAndProperties(args[0], utils.getExcel(args[1]), ...args.slice(2));
        } else if (lastFunctionIndex === 3) {
            return mapCollectErrorsExcelByStartWithCreatorAndProperties(args[0], utils.getExcel(args[1]), ...args.slice(2));
        } else if (lastFunctionIndex === 4) {
            return mapCollectErrorsExcelByRangeWithCreatorAndProperties(args[0], utils.getExcel(args[1]), ...args.slice(2));
        }
    }
    throw new utils.ExcelError('arguments error');
}

module.exports = mapCollectErrors;
