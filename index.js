const transform = require('./transform');
const map = require('./map');
const mapCollectErrors = require('./mapCollectErrors');
const mapIgnoreErrors = require('./mapIgnoreErrors');
const create = require('./create');

const forString = require('./string');
const forNumber = require('./number');
const forDate = require('./date');

module.exports = {
    transform,
    map,
    mapCollectErrors,
    mapIgnoreErrors,
    create,

    forString,
    forNumber,
    forDate,
}
