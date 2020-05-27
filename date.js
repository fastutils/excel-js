const utils = require('./utils');
const PropertyInfo = require('./property');
const moment = require('moment');

const forDateByIndexAndSetter =                                 (column, setter) => forDateByIndexAndSetterAndRequired(column, setter, utils.DEFAULT_REQUIRED);
const forDateByIndexAndSetterAndRequired =                      (column, setter, required) => forDateByIndexAndSetterAndValidator(column, setter, utils.defaultPropertyRequiredValidator(required));
const forDateByIndexAndSetterAndValidator =                     (column, setter, validator) => forDateByRangeAndCreatorAndReaderAndSetterAndValidator(column, column, utils.defaultCreator, utils.defaultPropertySetter(setter), utils.defaultSetter, utils.defaultPropertyValidator(validator));

const forDateByRangeAndSetter =                                 (minColumn, maxColumn, setter) => forDateByRangeAndSetterAndRequired(minColumn, maxColumn, setter, utils.DEFAULT_REQUIRED);
const forDateByRangeAndSetterAndRequired =                      (minColumn, maxColumn, setter, required) => forDateByRangeAndSetterAndValidator(minColumn, maxColumn, setter, utils.defaultPropertyRequiredValidator(required));
const forDateByRangeAndSetterAndValidator =                     (minColumn, maxColumn, setter, validator) => forDateByRangeAndCreatorAndReaderAndSetterAndValidator(minColumn, maxColumn, utils.defaultCreator, utils.defaultPropertySetter(setter), utils.defaultSetter, utils.defaultPropertyValidator(validator));

const forDateByIndexAndCreatorAndReaderAndSetter =              (column, creator, reader, setter) => forDateByIndexAndCreatorAndReaderAndSetterAndRequired(column, creator, reader, setter, utils.DEFAULT_REQUIRED);
const forDateByIndexAndCreatorAndReaderAndSetterAndRequired =   (column, creator, reader, setter, required) => forDateByIndexAndCreatorAndReaderAndSetterAndValidator(column, creator, reader, setter, utils.defaultRequiredValidator(required));
const forDateByIndexAndCreatorAndReaderAndSetterAndValidator =  (column, creator, reader, setter, validator) => forDateByRangeAndCreatorAndReaderAndSetterAndValidator(column, column, creator, reader, setter, validator);

const forDateByRangeAndCreatorAndReaderAndSetter =              (minColumn, maxColumn, creator, reader, setter) => forDateByRangeAndCreatorAndReaderAndSetterAndRequired(minColumn, maxColumn, creator, reader, setter, utils.DEFAULT_REQUIRED);
const forDateByRangeAndCreatorAndReaderAndSetterAndRequired =   (minColumn, maxColumn, creator, reader, setter, required) => forDateByRangeAndCreatorAndReaderAndSetterAndValidator(minColumn, maxColumn, creator, reader, setter, utils.defaultRequiredValidator(required));
const forDateByRangeAndCreatorAndReaderAndSetterAndValidator =  (minColumn, maxColumn, creator, reader, setter, validator) => {
    return new PropertyInfo(minColumn, maxColumn, cell => {
        if (cell.dateValue) {
            return cell.dateValue;
        } else {
            let dateValue = moment(cell.value, ['YYYY-MM-DD HH:mm:ss', 'YYYY/MM/DD HH:mm:ss', 'YYYY-MM-DD', 'YYYY/MM/DD']);
            if (dateValue.isValid()) {
                return dateValue.toDate();
            } else {
                throw new utils.ExcelValueFormatError(cell.row, cell.column, cell.value, 'value is a invalid date or time:' + cell.value);
            }
        }
    }, creator, reader, setter, validator);
};

const forDate = (...args) => {
    let lastBoolean = utils.isBoolean(args[args.length - 1]);
    if (utils.isFunction(args[1])) {
        if (args.length === 2) {
            return forDateByIndexAndSetter(...args);
        } else if (args.length === 3) {
            if (lastBoolean) {
                return forDateByIndexAndSetterAndRequired(...args);
            } else {
                return forDateByIndexAndSetterAndValidator(...args);
            }
        } else if (args.length === 4) {
            return forDateByIndexAndCreatorAndReaderAndSetter(...args);
        } else if (args.length === 5) {
            if (lastBoolean) {
                return forDateByIndexAndCreatorAndReaderAndSetterAndRequired(...args);
            } else {
                return forDateByIndexAndCreatorAndReaderAndSetterAndValidator(...args);
            }
        }
    } else {
        if (args.length === 3) {
            return forDateByRangeAndSetter(...args);
        } else if (args.length === 4) {
            if (lastBoolean) {
                return forDateByRangeAndSetterAndRequired(...args);
            } else {
                return forDateByRangeAndSetterAndValidator(...args);
            }
        } else if (args.length === 5) {
            return forDateByRangeAndCreatorAndReaderAndSetter(...args);
        } else if (args.length === 6) {
            if (lastBoolean) {
                return forDateByRangeAndCreatorAndReaderAndSetterAndRequired(...args);
            } else {
                return forDateByRangeAndCreatorAndReaderAndSetterAndValidator(...args);
            }
        }
    }
}

module.exports = forDate;
