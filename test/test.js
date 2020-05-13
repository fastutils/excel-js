const excel = require('../index');
const fs = require('fs');
const path = require('path');

let result = excel.transform(path.join(__dirname, 'test.xlsx'), row => row.cells.map(cell => cell.value).join('|'));
console.log(result);
let result2 = excel.transform(path.join(__dirname, 'test.xlsx'), _ => ({}),
    (obj, cell) => obj.A = cell.value,
    (obj, cell) => obj.B = cell.value,
    (obj, cell) => obj.C = cell.value
    )
console.log(result2);
