const fs = require('fs');
const XLSX = require('xlsx');
const data = fs.readFileSync('renwu-04.csv', 'binary');
const workbook = XLSX.read(data, {type: 'binary'});
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const result = XLSX.utils.sheet_to_json(sheet, {header: 1});
console.log('Rows:', result.length);
if (result.length > 0) {
  console.log('Row 0 length:', result[0].length);
  console.log('Row 1 length:', result[1] ? result[1].length : 'none');
}
