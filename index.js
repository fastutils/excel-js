const transform = require('./transform');
const map = require('./map');
const mapCollectErrors = require('./mapCollectErrors');
const create = require('./create');

const forString = require('./string');
const forNumber = require('./number');
const forDate = require('./date');

module.exports = {
    transform,
    map,
    mapCollectErrors,
    create,

    forString,
    forNumber,
    forDate,
}
