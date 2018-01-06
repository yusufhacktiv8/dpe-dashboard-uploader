const XLSX = require('xlsx');

const BAD_SHEET_POSITION = 2;
const MAXIMUM_ROW = 150;

const readCell = (worksheet, rowNum, colNum) => {
  const cellIdentity = XLSX.utils.encode_cell({ r: rowNum, c: colNum });
  return worksheet[cellIdentity] ? worksheet[cellIdentity].v : '';
};

const isEndOfData = (worksheet, rowNum) => {
  const cellData = readCell(worksheet, rowNum, 0);
  if (cellData === 'TOTAL') {
    return true;
  }
  return false;
};

exports.readBad = function (fileName, callback) {
  var workbook = XLSX.readFile(fileName);

  const workbook = XLSX.readFile(fileName);
  const worksheet = workbook.Sheets[workbook.SheetNames[BAD_SHEET_POSITION]];

  // const yearCellValue = worksheet.F6.v;
  const YEAR = 2017;
  const MONTH = 1;
  const bads = [];
  const startRow = 27;
  for (let i = startRow; i < (startRow + MAXIMUM_ROW); i += 1) {
    if (isEndOfData(worksheet, i)) {
      break;
    }
    const projectCode = readCell(worksheet, i, 6);
    const piutangUsaha = readCell(worksheet, i, 1);
    const tagihanBruto = readCell(worksheet, i, 2);
    const piutangRetensi = readCell(worksheet, i, 3);
    const pdp = readCell(worksheet, i, 4);
    const bad = readCell(worksheet, i, 5);
    bads.push({
      projectCode,
      piutangUsaha,
      tagihanBruto,
      piutangRetensi,
      pdp,
      bad,
      year,
    });
  }

  callback({ year: YEAR, month: MONTH, payload: bads });
};
