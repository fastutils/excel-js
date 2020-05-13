const xlsx = require('xlsx');
const fs = require('fs');

const words = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

const isString = value => typeof value === 'string' || value instanceof String;
const isFunction = value => typeof value === 'function' || value instanceof Function;
const isBoolean = value => typeof value === 'boolean' || value instanceof Boolean;
const parseColumnIndex = value => {
    value = value.split('');
    let result = 0;
    for (let i = 0; i < value.length; i++) {
        let v = words.indexOf(value[i]);
        for (let j = i + 1; j < value.length; j++) {
            v *= words.length;
        }
        result += v;
    }
    return result;
}
const parseRowIndex = value => parseInt(value);
const parseCellIndex = value => ({
    row: parseRowIndex(value.replace(/^[A-Z]+/ig, '')),
    column: parseColumnIndex(value.replace(/[0-9]+$/ig, ''))
});
const parseRangeIndex = value => {
    value = value.split(':');
    return {
        start: parseCellIndex(value[0]),
        end: parseCellIndex(value[1])
    };
};
const formatColumnIndex = index => {
    let result = '';
    do {
        result = words[index % words.length] + result;
        index = Math.floor(index / words.length);
    } while(index !== 0);
    return result;
};
const formatRowIndex = index => index;
const formatCellIndex = (rowIndex, cellIndex) => formatColumnIndex(cellIndex) + formatRowIndex(rowIndex);

class ExcelError extends Error {
    constructor(message) {
        super(message);
    }
}
class ExcelValueFormatError extends ExcelError {
    constructor(rowIndex, columnIndex, message) {
        super(`单元格格式错误[${formatCellIndex(rowIndex, columnIndex)}]:${message instanceof Error ? message.message : message}`);
        this.rowIndex = rowIndex;
        this.columnIndex = columnIndex;
    }
}

class Excel {
    constructor(delegate) {
        this.delegate = delegate;
    }
    getSheet(sheetIndex) {
        return new Sheet(this.delegate.Sheets[this.delegate.SheetNames[sheetIndex]]);
    }
}
class Sheet {
    constructor(delegate) {
        this.delegate = delegate;
    }
    get rows() {
        let range = parseRangeIndex(this.delegate['!ref']);
        return range.end.row;
    }
    get columns() {
        let range = parseRangeIndex(this.delegate['!ref']);
        return range.end.column + 1;
    }
    getCell(rowIndex, columnIndex) {
        return new Cell(this.delegate[formatCellIndex(rowIndex + 1, columnIndex)], rowIndex, columnIndex, this);
    }
    getRow(rowIndex) {
        return new Row(new Array(this.columns).fill(false).map((v, i) => this.getCell(rowIndex, i)), rowIndex, this);
    }
    getColumn(columnIndex) {
        return new Column(new Array(this.rows).fill(false).map((v, i) => this.getCell(i, columnIndex)), columnIndex, this);
    }
}
class Row {
    constructor(cells, index, sheet) {
        this.cells = cells;
        this.index = index;
        this.sheet = sheet;
    }
    getCell(columnIndex) {
        return this.cells[columnIndex];
    }
    get columns() {
        return this.sheet.columns;
    }
}
class Column {
    constructor(cells, index, sheet) {
        this.cells = cells;
        this.index = index;
        this.sheet = sheet;
    }
    getCell(rowIndex) {
        return this.cells[rowIndex];
    }
    get rows() {
        return this.sheet.rows;
    }
}
class Cell {
    constructor(delegate, rowIndex, columnIndex, sheet) {
        this.delegate = delegate;
        this.rowIndex = rowIndex;
        this.columnIndex = columnIndex;
        this.sheet = sheet;
    }
    get value() {
        return this.delegate ? this.delegate.w : null;
    }
    get dateValue() {
        return this.delegate ? this.delegate.d : null;
    }
    get row() {
        return this.sheet.getRow(this.rowIndex);
    }
    get column() {
        return this.sheet.getColumn(this.columnIndex);
    }
}

const getSheet = (excel, sheetIndex) => {
    return getExcel(excel).getSheet(sheetIndex);
}
const getExcel = (excel) => {
    if (excel instanceof Excel) {
        return excel;
    } else if (isString(excel)) {
        return new Excel(xlsx.readFile(excel, {
            cellDates: true,
        }));
    } else {
        return new Excel(excel);
    }
}

const defaultCreator = t => t;
const defaultSetter = (t, p) => { };
const defaultPropertySetter = setter => (t, p, v, r, c) => setter(p, v, r, c);
const defaultValidator = required => (t, p, v, r, c) => required && v == null ? new ExcelValueFormatError(r, c, "Can not be null") : null;


module.exports = {
    isString,
    isFunction,
    isBoolean,

    getSheet,
    getExcel,

    Excel,
    Sheet,

    ExcelError,
    ExcelValueFormatError,

    defaultCreator,
    defaultSetter,
    defaultValidator,
    defaultPropertySetter,

    DEFAULT_REQUIRED: true
}
