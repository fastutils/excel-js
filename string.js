const utils = require('./utils');
const PropertyInfo = require('./property');

const forStringByIndexAndSetter =                                 (column, setter) => forStringByIndexAndSetterAndRequired(column, setter, utils.DEFAULT_REQUIRED);
const forStringByIndexAndSetterAndRequired =                      (column, setter, required) => forStringByIndexAndSetterAndValidator(column, setter, utils.defaultPropertyRequiredValidator(required));
const forStringByIndexAndSetterAndValidator =                     (column, setter, validator) => forStringByRangeAndSetterAndValidator(column, column, setter, utils.defaultPropertyValidator(validator));

const forStringByRangeAndSetter =                                 (minColumn, maxColumn, setter) => forStringByRangeAndSetterAndRequired(minColumn, maxColumn, setter, utils.DEFAULT_REQUIRED);
const forStringByRangeAndSetterAndRequired =                      (minColumn, maxColumn, setter, required) => forStringByRangeAndSetterAndValidator(minColumn, maxColumn, setter, utils.defaultPropertyRequiredValidator(required));
const forStringByRangeAndSetterAndValidator =                     (minColumn, maxColumn, setter, validator) => forStringByRangeAndCreatorAndReaderAndSetterAndValidator(minColumn, maxColumn, utils.defaultCreator, utils.defaultPropertySetter(setter), utils.defaultSetter, utils.defaultPropertyValidator(validator));

const forStringByIndexAndCreatorAndReaderAndSetter =              (column, creator, reader, setter) => forStringByIndexAndCreatorAndReaderAndSetterAndRequired(column, creator, reader, setter, utils.DEFAULT_REQUIRED);
const forStringByIndexAndCreatorAndReaderAndSetterAndRequired =   (column, creator, reader, setter, required) => forStringByIndexAndCreatorAndReaderAndSetterAndValidator(column, creator, reader, setter, utils.defaultRequiredValidator(required));
const forStringByIndexAndCreatorAndReaderAndSetterAndValidator =  (column, creator, reader, setter, validator) => forStringByRangeAndCreatorAndReaderAndSetterAndValidator(column, column, creator, reader, setter, validator);

const forStringByRangeAndCreatorAndReaderAndSetter =              (minColumn, maxColumn, creator, reader, setter) => forStringByRangeAndCreatorAndReaderAndSetterAndRequired(minColumn, maxColumn, creator, reader, setter, utils.DEFAULT_REQUIRED);
const forStringByRangeAndCreatorAndReaderAndSetterAndRequired =   (minColumn, maxColumn, creator, reader, setter, required) => forStringByRangeAndCreatorAndReaderAndSetterAndValidator(minColumn, maxColumn, creator, reader, setter, utils.defaultRequiredValidator(required));
const forStringByRangeAndCreatorAndReaderAndSetterAndValidator =  (minColumn, maxColumn, creator, reader, setter, validator) => {
    return new PropertyInfo(minColumn, maxColumn, cell => {
        return cell.value;
    }, creator, reader, setter, validator);
};

const forString = (...args) => {
    let lastBoolean = utils.isBoolean(args[args.length - 1]);
    if (utils.isFunction(args[1])) {
        if (args.length === 2) {
            return forStringByIndexAndSetter(...args);
        } else if (args.length === 3) {
            if (lastBoolean) {
                return forStringByIndexAndSetterAndRequired(...args);
            } else {
                return forStringByIndexAndSetterAndValidator(...args);
            }
        } else if (args.length === 4) {
            return forStringByIndexAndCreatorAndReaderAndSetter(...args);
        } else if (args.length === 5) {
            if (lastBoolean) {
                return forStringByIndexAndCreatorAndReaderAndSetterAndRequired(...args);
            } else {
                return forStringByIndexAndCreatorAndReaderAndSetterAndValidator(...args);
            }
        }
    } else {
        if (args.length === 3) {
            return forStringByRangeAndSetter(...args);
        } else if (args.length === 4) {
            if (lastBoolean) {
                return forStringByRangeAndSetterAndRequired(...args);
            } else {
                return forStringByRangeAndSetterAndValidator(...args);
            }
        } else if (args.length === 5) {
            return forStringByRangeAndCreatorAndReaderAndSetter(...args);
        } else if (args.length === 6) {
            if (lastBoolean) {
                return forStringByRangeAndCreatorAndReaderAndSetterAndRequired(...args);
            } else {
                return forStringByRangeAndCreatorAndReaderAndSetterAndValidator(...args);
            }
        }
    }
}

module.exports = forString;
