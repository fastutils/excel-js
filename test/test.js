const excel = require('../index');
const utils = require('../utils');
const fs = require('fs');
const path = require('path');

let result = excel.transform(path.join(__dirname, 'test.xlsx'), row => row.cells.map(cell => cell.value).join('|'));
console.log(result);
let result2 = excel.transform(path.join(__dirname, 'test.xlsx'), () => ({}),
    (obj, cell) => obj.A = cell.value,
    (obj, cell) => obj.B = cell.value,
    (obj, cell) => obj.C = cell.value,
);
console.log(result2);
let result3 = excel.map(path.join(__dirname, 'test.xlsx'), 1, () => ({}),
    excel.forNumber(0, (o, v) => o.id = v),
    excel.forString(1, (o, v) => o.name = v),
    excel.forNumber(3, (o, v) => o.scole = v),
    excel.forNumber(4, (o, v) => o.english = v),
    excel.forNumber(5, (o, v) => o.chinese = v),
    excel.forNumber(6, (o, v) => o.math = v),
);
console.log(result3);
try {
    excel.mapCollectErrors(path.join(__dirname, 'test.xlsx'), () => ({}),
        excel.forNumber(0, (o, v) => o.id = v),
        excel.forString(1, (o, v) => o.name = v),
        excel.forNumber(3, (o, v) => o.scole = v),
        excel.forNumber(4, (o, v) => o.english = v),
        excel.forNumber(5, (o, v) => o.chinese = v),
        excel.forNumber(6, (o, v) => o.math = v),
    );
} catch (e) {
    console.log(e.message);
}
let result4 = excel.mapIgnoreErrors(path.join(__dirname, 'test.xlsx'), () => ({}),
    excel.forNumber(0, (o, v) => o.id = v),
    excel.forString(1, (o, v) => o.name = v),
    excel.forNumber(3, (o, v) => o.scole = v),
    excel.forNumber(4, (o, v) => o.english = v),
    excel.forNumber(5, (o, v) => o.chinese = v),
    excel.forNumber(6, (o, v) => o.math = v),
);
console.log(result4);
let result5 = [];
excel.mapIgnoreErrors(result5, path.join(__dirname, 'test.xlsx'), () => ({}),
    excel.forNumber(0, (o, v) => o.id = v),
    excel.forString(1, (o, v) => o.name = v),
    excel.forNumber(3, (o, v) => o.scole = v),
    excel.forNumber(4, (o, v) => o.english = v),
    excel.forNumber(5, (o, v) => o.chinese = v),
    excel.forNumber(6, (o, v) => o.math = v),
);
console.log(result5.map(e => e.message));
