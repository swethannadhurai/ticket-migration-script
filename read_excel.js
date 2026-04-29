const xlsx = require('xlsx');

const workbook = xlsx.readFile('sheet/Service Request.xlsx');
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

console.log("Excel columns:", data[0]);
