const utils = require('./utils');

class PropertyInfo {
    constructor(minColumn, maxColumn, parser, creator, reader, setter, validator) {
        this.minColumn = minColumn;
        this.maxColumn = maxColumn;
        this.parser = parser;
        this.creator = creator;
        this.reader = reader;
        this.setter = setter;
        this.validator = validator;
    }

    set(t, row, rowIndex) {
        let minColumn = this.minColumn == null ? 0 : this.minColumn < 0 ? row.cols + this.minColumn : this.minColumn;
        let maxColumn = this.maxColumn == null ? row.cols - 1 : this.maxColumn < 0 ? row.cols + this.maxColumn - 1 : this.maxColumn;
        let property = this.creator(t);
        for (let columnIndex = minColumn; columnIndex <= maxColumn; columnIndex++) {
            let cell = row.getCell(columnIndex);
            let value = cell == null ? null : this.parser(cell);
            let exception = this.validator(t, property, value, rowIndex, columnIndex);
            if (exception == null) {
                this.reader(t, property, value, rowIndex, columnIndex);
            } else {
                throw exception;
            }
        }
        this.setter(t, property);
    }
}

module.exports = PropertyInfo;
