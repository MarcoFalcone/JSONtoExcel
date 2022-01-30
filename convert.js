const XLSX = require('xlsx');
const json = require('./en/en-us.json'); // import json file

delete json.creationDate; // if creation date present

const flattenObject = (obj, prefix = '') => Object.keys(obj).reduce((acc, k) => { // if json object is nested
  const pre = prefix.length ? `${prefix}.` : '';
  if (
    typeof obj[k] === 'object'
      && obj[k] !== null
      && Object.keys(obj[k]).length > 0
  ) { Object.assign(acc, flattenObject(obj[k], pre + k)); } else acc[pre + k] = obj[k];
  return acc;
}, {});

const flattenedJson = flattenObject(json);

const wscols = [ // columns width
  { wch: 40 },
  { wch: 40 },
];

function ec(r, c) {
  return XLSX.utils.encode_cell({ r, c });
}
function deleteRow(ws, rowIndex) { // delete rows if needed
  const variable = XLSX.utils.decode_range(ws['!ref']);
  for (let R = rowIndex; R < variable.e.r; ++R) {
    for (let C = variable.s.c; C <= variable.e.c; ++C) {
      ws[ec(R, C)] = ws[ec(R + 1, C)];
    }
  }
  variable.e.r--;
  ws['!ref'] = XLSX.utils.encode_range(variable.s, variable.e);
}

const convertJsonToExcel = () => {
  const workSheet = XLSX.utils.json_to_sheet(Object.entries(flattenedJson));
  workSheet['!cols'] = wscols;
  const workBook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(workBook, workSheet, 'translationsToSend');

  deleteRow(workSheet, 0);

  XLSX.writeFile(workBook, 'translationsToSend.xlsx');
};

convertJsonToExcel();
