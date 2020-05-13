const xlsx = require('xlsx');
const fs = require('fs');

const words = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

const isString = value => typeof value === 'string' || value instanceof String;
const isFunction = value => typeof value === 'function' || value instanceof Function;
const parseColIndex = value => {
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
    col: parseColIndex(value.replace(/[0-9]+$/ig, ''))
});
const parseRangeIndex = value => {
    value = value.split(':');
    return {
        start: parseCellIndex(value[0]),
        end: parseCellIndex(value[1])
    };
};
const formatColIndex = index => {
    let result = '';
    do {
        result = words[index % words.length] + result;
        index = Math.floor(index / words.length);
    } while(index !== 0);
    return result;
};
const formatRowIndex = index => index;
const formatCellIndex = (rowIndex, cellIndex) => formatColIndex(cellIndex) + formatRowIndex(rowIndex);

class ExcelError extends Error {
    constructor(message) {
        super(message);
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
        console.log(range);
        return range.end.row;
    }
    get cols() {
        let range = parseRangeIndex(this.delegate['!ref']);
        return range.end.col + 1;
    }
    getCell(rowIndex, colIndex) {
        return new Cell(this.delegate[formatCellIndex(rowIndex + 1, colIndex)], rowIndex, colIndex, this);
    }
    getRow(rowIndex) {
        return new Row(new Array(this.cols).fill(false).map((v, i) => this.getCell(rowIndex, i)), rowIndex, this);
    }
    getCol(colIndex) {
        return new Col(new Array(this.rows).fill(false).map((v, i) => this.getCell(i, colIndex)), colIndex, this);
    }
}
class Row {
    constructor(cells, index, sheet) {
        this.cells = cells;
        this.index = index;
        this.sheet = sheet;
    }
    getCell(colIndex) {
        return this.cells[colIndex];
    }
    get cols() {
        return this.sheet.cols;
    }
}
class Col {
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
    constructor(delegate, rowIndex, colIndex, sheet) {
        this.delegate = delegate;
        this.rowIndex = rowIndex;
        this.colIndex = colIndex;
        this.sheet = sheet;
    }
    get value() {
        return this.delegate ? this.delegate.w : null;
    }
    get row() {
        return this.sheet.getRow(this.rowIndex);
    }
    get col() {
        return this.sheet.getCol(this.colIndex);
    }
}

const getSheet = (excel, sheetIndex) => {
    return getExcel(excel).getSheet(sheetIndex);
}
const getExcel = (excel) => {
    if (excel instanceof Excel) {
        return excel;
    } else if (isString(excel)) {
        return new Excel(xlsx.readFile(excel, { }));
    } else {
        return new Excel(excel);
    }
}

module.exports = {
    isString,
    isFunction,

    getSheet,
    getExcel,

    Excel,
    Sheet,

    ExcelError
}
