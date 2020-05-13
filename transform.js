const utils = require('./utils');

const transformFromSheetByRange =       (sheet, startIndex, count, parser)              => {
    let result = [];
    console.log(sheet, startIndex, count, parser);
    for (let r = startIndex; r < startIndex + count; r++) {
        result.push(parser(sheet.getRow(r)));
    }
    return result;
};
const transformFromSheetByStart =         (sheet, startIndex, parser)                     => transformFromSheetByRange(sheet, startIndex, sheet.rows + 1 - startIndex, parser);
const transformFromSheet =                (sheet, parser)                                 => transformFromSheetByStart(sheet, 0, parser);
const transformFromFirstSheet =           (excel, parser)                                 => transformFromExcel(0, excel, parser);
const transformFromFirstSheetByStart =    (excel, startIndex, parser)                     => transformFromExcelByStart(0, excel, startIndex, parser);
const transformFromFirstSheetByRange =    (excel, startIndex, count, parser)              => transformFromExcelByRange(0, excel, startIndex, count, parser);
const transformFromExcel =                (sheetIndex, excel, parser)                     => transformFromSheet(utils.getSheet(excel, sheetIndex), parser);
const transformFromExcelByStart =         (sheetIndex, excel, startIndex, parser)         => transformFromSheetByStart(utils.getSheet(excel, sheetIndex), startIndex, parser);
const transformFromExcelByRange =         (sheetIndex, excel, startIndex, count, parser)  => transformFromSheetByRange(utils.getSheet(excel, sheetIndex), startIndex, count, parser);

const transformFromSheetByRangeWithCreatorAndSetters =  (sheet, startIndex, count, creator, ...setters) => {
    let result = [];
    for (let r = startIndex; r < startIndex + count; r++) {
        let row = sheet.getRow(r);
        let obj = creator();
        for (let i = 0; i < row.cols; i++) {
            if (setters[i]) {
                setters[i](obj, row.cells[i]);
            }
        }
        result.push(obj);
    }
    return result;
};
const transformFromSheetWithCreatorAndSetters =         (sheet, creator, ...setters) => transformFromSheetByStartWithCreatorAndSetters(sheet, 0, creator, ...setters);
const transformFromSheetByStartWithCreatorAndSetters =  (sheet, startIndex, creator, ...setters) => transformFromSheetByRangeWithCreatorAndSetters(sheet, 0, sheet.rows + 1 - startIndex, creator, ...setters);

const transformFromFirstSheetWithCreatorAndSetters =         (excel, creator, ...setters) => transformFromExcelWithCreatorAndSetters(0, excel, creator, ...setters);
const transformFromFirstSheetByStartWithCreatorAndSetters =  (excel, startIndex, creator, ...setters) => transformFromExcelByStartWithCreatorAndSetters(0, excel, creator, ...setters);
const transformFromFirstSheetByRangeWithCreatorAndSetters =  (excel, startIndex, count, creator, ...setters) => transformFromExcelByRangeWithCreatorAndSetters(0, excel, startIndex, count, creator, ...setters);

const transformFromExcelWithCreatorAndSetters =         (sheetIndex, excel, creator, ...setters) => transformFromSheetWithCreatorAndSetters(utils.getSheet(excel, sheetIndex), creator, ...setters);
const transformFromExcelByStartWithCreatorAndSetters =  (sheetIndex, excel, startIndex, creator, ...setters) => transformFromSheetByStartWithCreatorAndSetters(utils.getSheet(excel, sheetIndex), startIndex, creator, ...setters);
const transformFromExcelByRangeWithCreatorAndSetters =  (sheetIndex, excel, startIndex, count, creator, ...setters) => transformFromSheetByRangeWithCreatorAndSetters(utils.getSheet(excel, sheetIndex), startIndex, count, creator, ...setters);

const transform = (...args) => {
    console.log(args);

    if (args[0] instanceof utils.Sheet) {
        let lastFunctionIndex = args.slice(1).findIndex(arg => utils.isFunction(arg)) + 1;
        if (args.length - lastFunctionIndex > 1) {
            if (lastFunctionIndex === 1) {
                return transformFromSheetWithCreatorAndSetters(...args);
            } else if (lastFunctionIndex === 2) {
                return transformFromSheetByStartWithCreatorAndSetters(...args);
            } else if (lastFunctionIndex === 3) {
                return transformFromSheetByRangeWithCreatorAndSetters(...args);
            }
        } else {
            if (lastFunctionIndex === 1) {
                return transformFromSheet(...args);
            } else if (lastFunctionIndex === 2) {
                return transformFromSheetByStart(...args);
            } else if (lastFunctionIndex === 3) {
                return transformFromSheetByRange(...args);
            }
        }
    } else if (args[0] instanceof utils.Excel || utils.isString(args[0])) {
        let lastFunctionIndex = args.slice(1).findIndex(arg => utils.isFunction(arg)) + 1;
        if (args.length - lastFunctionIndex > 1) {
            if (lastFunctionIndex === 1) {
                return transformFromFirstSheetWithCreatorAndSetters(utils.getExcel(args[0]), ...args.slice(1));
            } else if (lastFunctionIndex === 2) {
                return transformFromFirstSheetByStartWithCreatorAndSetters(utils.getExcel(args[0]), ...args.slice(1));
            } else if (lastFunctionIndex === 3) {
                return transformFromFirstSheetByRangeWithCreatorAndSetters(utils.getExcel(args[0]), ...args.slice(1));
            }
        } else {
            if (lastFunctionIndex === 1) {
                return transformFromFirstSheet(utils.getExcel(args[0]), ...args.slice(1));
            } else if (lastFunctionIndex === 2) {
                return transformFromFirstSheetByStart(utils.getExcel(args[0]), ...args.slice(1));
            } else if (lastFunctionIndex === 3) {
                return transformFromFirstSheetByRange(utils.getExcel(args[0]), ...args.slice(1));
            }
        }
    } else {
        let lastFunctionIndex = args.slice(2).findIndex(arg => utils.isFunction(arg)) + 2;
        if (args.length - lastFunctionIndex > 1) {
            if (lastFunctionIndex === 2) {
                return transformFromExcelWithCreatorAndSetters(args[0], utils.getExcel(args[1]), ...args.slice(2));
            } else if (lastFunctionIndex === 3) {
                return transformFromExcelByStartWithCreatorAndSetters(args[0], utils.getExcel(args[1]), ...args.slice(2));
            } else if (lastFunctionIndex === 4) {
                return transformFromExcelByRangeWithCreatorAndSetters(args[0], utils.getExcel(args[1]), ...args.slice(2));
            }
        } else {
            if (lastFunctionIndex === 2) {
                return transformFromExcel(args[0], utils.getExcel(args[1]), ...args.slice(2));
            } else if (lastFunctionIndex === 3) {
                return transformFromExcelByStart(args[0], utils.getExcel(args[1]), ...args.slice(2));
            } else if (lastFunctionIndex === 4) {
                return transformFromExcelByRange(args[0], utils.getExcel(args[1]), ...args.slice(2));
            }
        }
    }
    throw new utils.ExcelError('arguments error');
}

module.exports = transform;
