const utils = require('./utils');
const PropertyInfo = require('./property');

const forNumberByIndexAndSetter =                                 (column, setter) => forNumberByIndexAndSetterAndRequired(column, setter, utils.DEFAULT_REQUIRED);
const forNumberByIndexAndSetterAndRequired =                      (column, setter, required) => forNumberByIndexAndSetterAndValidator(column, setter, utils.defaultPropertyRequiredValidator(required));
const forNumberByIndexAndSetterAndValidator =                     (column, setter, validator) => forNumberByRangeAndCreatorAndReaderAndSetterAndValidator(column, column, utils.defaultCreator, utils.defaultPropertySetter(setter), utils.defaultSetter, utils.defaultPropertyValidator(validator));

const forNumberByRangeAndSetter =                                 (minColumn, maxColumn, setter) => forNumberByRangeAndSetterAndRequired(minColumn, maxColumn, setter, utils.DEFAULT_REQUIRED);
const forNumberByRangeAndSetterAndRequired =                      (minColumn, maxColumn, setter, required) => forNumberByRangeAndSetterAndValidator(minColumn, maxColumn, setter, utils.defaultPropertyRequiredValidator(required));
const forNumberByRangeAndSetterAndValidator =                     (minColumn, maxColumn, setter, validator) => forNumberByRangeAndCreatorAndReaderAndSetterAndValidator(minColumn, maxColumn, utils.defaultCreator, utils.defaultPropertySetter(setter), utils.defaultSetter, utils.defaultPropertyValidator(validator));

const forNumberByIndexAndCreatorAndReaderAndSetter =              (column, creator, reader, setter) => forNumberByIndexAndCreatorAndReaderAndSetterAndRequired(column, creator, reader, setter, utils.DEFAULT_REQUIRED);
const forNumberByIndexAndCreatorAndReaderAndSetterAndRequired =   (column, creator, reader, setter, required) => forNumberByIndexAndCreatorAndReaderAndSetterAndValidator(column, creator, reader, setter, utils.defaultRequiredValidator(required));
const forNumberByIndexAndCreatorAndReaderAndSetterAndValidator =  (column, creator, reader, setter, validator) => forNumberByRangeAndCreatorAndReaderAndSetterAndValidator(column, column, creator, reader, setter, validator);

const forNumberByRangeAndCreatorAndReaderAndSetter =              (minColumn, maxColumn, creator, reader, setter) => forNumberByRangeAndCreatorAndReaderAndSetterAndRequired(minColumn, maxColumn, creator, reader, setter, utils.DEFAULT_REQUIRED);
const forNumberByRangeAndCreatorAndReaderAndSetterAndRequired =   (minColumn, maxColumn, creator, reader, setter, required) => forNumberByRangeAndCreatorAndReaderAndSetterAndValidator(minColumn, maxColumn, creator, reader, setter, utils.defaultRequiredValidator(required));
const forNumberByRangeAndCreatorAndReaderAndSetterAndValidator =  (minColumn, maxColumn, creator, reader, setter, validator) => {
    return new PropertyInfo(minColumn, maxColumn, cell => {
        if (cell.value === 'NaN') {
            return NaN;
        } else if (isNaN(cell.value)) {
            throw new utils.ExcelValueFormatError(cell.rowIndex, cell.columnIndex, cell.value, 'value is not a number:' + cell.value);
        } else {
            return parseFloat(cell.value);
        }
    }, creator, reader, setter, validator);
};

const forNumber = (...args) => {
    let lastBoolean = utils.isBoolean(args[args.length - 1]);
    if (utils.isFunction(args[1])) {
        if (args.length === 2) {
            return forNumberByIndexAndSetter(...args);
        } else if (args.length === 3) {
            if (lastBoolean) {
                return forNumberByIndexAndSetterAndRequired(...args);
            } else {
                return forNumberByIndexAndSetterAndValidator(...args);
            }
        } else if (args.length === 4) {
            return forNumberByIndexAndCreatorAndReaderAndSetter(...args);
        } else if (args.length === 5) {
            if (lastBoolean) {
                return forNumberByIndexAndCreatorAndReaderAndSetterAndRequired(...args);
            } else {
                return forNumberByIndexAndCreatorAndReaderAndSetterAndValidator(...args);
            }
        }
    } else {
        if (args.length === 3) {
            return forNumberByRangeAndSetter(...args);
        } else if (args.length === 4) {
            if (lastBoolean) {
                return forNumberByRangeAndSetterAndRequired(...args);
            } else {
                return forNumberByRangeAndSetterAndValidator(...args);
            }
        } else if (args.length === 5) {
            return forNumberByRangeAndCreatorAndReaderAndSetter(...args);
        } else if (args.length === 6) {
            if (lastBoolean) {
                return forNumberByRangeAndCreatorAndReaderAndSetterAndRequired(...args);
            } else {
                return forNumberByRangeAndCreatorAndReaderAndSetterAndValidator(...args);
            }
        }
    }
}

module.exports = forNumber;
