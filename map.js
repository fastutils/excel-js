const utils = require('./utils');


const mapSheetWithCreatorAndProperties =                (sheet, creator, ...properties) => mapSheetByStartWithCreatorAndProperties(sheet, 0, creator, ...properties);
const mapSheetByStartWithCreatorAndProperties =         (sheet, startIndex, creator, ...properties) => mapSheetByRangeWithCreatorAndProperties(sheet, startIndex, sheet.rows + 1 - startIndex, creator, ...properties);
const mapSheetByRangeWithCreatorAndProperties =         (sheet, startIndex, count, creator, ...properties) => {
    let result = [];
    for (let r = startIndex; r < startIndex + count; r++) {
        let row = sheet.getRow(r);
        if (row != null && row.columns > 0) {
            let obj = creator();
            for (let property of properties) {
                property.set(obj, row, r);
            }
            result.push(obj);
        }
    }
    return result;
}

const mapExcelWithCreatorAndProperties =                (sheetIndex, excel, creator, ...properties) => mapSheetWithCreatorAndProperties(utils.getSheet(excel, sheetIndex), creator, ...properties);
const mapExcelByStartWithCreatorAndProperties =         (sheetIndex, excel, startIndex, creator, ...properties) => mapSheetByStartWithCreatorAndProperties(utils.getSheet(excel, sheetIndex), startIndex, creator, ...properties);
const mapExcelByRangeWithCreatorAndProperties =         (sheetIndex, excel, startIndex, count, creator, ...properties) => mapSheetByRangeWithCreatorAndProperties(utils.getSheet(excel, sheetIndex), startIndex, count, creator, ...properties);
const mapFirstSheetWithCreatorAndProperties =           (excel, creator, ...properties) => mapExcelWithCreatorAndProperties(0, excel, creator, ...properties);
const mapFirstSheetByStartWithCreatorAndProperties =    (excel, startIndex, creator, ...properties) => mapExcelByStartWithCreatorAndProperties(0, excel, startIndex, creator, ...properties);
const mapFirstSheetByRangeWithCreatorAndProperties =    (excel, startIndex, count, creator, ...properties) => mapExcelByRangeWithCreatorAndProperties(0, excel, startIndex, count, creator, ...properties);


const map = (...args) => {
    if (args[0] instanceof utils.Sheet) {
        let lastFunctionIndex = args.slice(1).findIndex(arg => utils.isFunction(arg)) + 1;
        if (lastFunctionIndex === 1) {
            return mapSheetWithCreatorAndProperties(...args);
        } else if (lastFunctionIndex === 2) {
            return mapSheetByStartWithCreatorAndProperties(...args);
        } else if (lastFunctionIndex === 3) {
            return mapSheetByRangeWithCreatorAndProperties(...args);
        }
    } else if (args[0] instanceof utils.Excel || utils.isString(args[0])) {
        let lastFunctionIndex = args.slice(1).findIndex(arg => utils.isFunction(arg)) + 1;
        if (lastFunctionIndex === 1) {
            return mapFirstSheetWithCreatorAndProperties(utils.getExcel(args[0]), ...args.slice(1));
        } else if (lastFunctionIndex === 2) {
            return mapFirstSheetByStartWithCreatorAndProperties(utils.getExcel(args[0]), ...args.slice(1));
        } else if (lastFunctionIndex === 3) {
            return mapFirstSheetByRangeWithCreatorAndProperties(utils.getExcel(args[0]), ...args.slice(1));
        }
    } else {
        let lastFunctionIndex = args.slice(2).findIndex(arg => utils.isFunction(arg)) + 2;
        if (lastFunctionIndex === 2) {
            return mapExcelWithCreatorAndProperties(args[0], utils.getExcel(args[1]), ...args.slice(2));
        } else if (lastFunctionIndex === 3) {
            return mapExcelByStartWithCreatorAndProperties(args[0], utils.getExcel(args[1]), ...args.slice(2));
        } else if (lastFunctionIndex === 4) {
            return mapExcelByRangeWithCreatorAndProperties(args[0], utils.getExcel(args[1]), ...args.slice(2));
        }
    }
    throw new utils.ExcelError('arguments error');
}

module.exports = map;
