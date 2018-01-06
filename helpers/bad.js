const XLSX = require('xlsx');

const BAD_SHEET_POSITION = 2;
const MAXIMUM_ROW = 150;

const readCell = (worksheet, rowNum, colNum) => {
  const cellIdentity = XLSX.utils.encode_cell({ r: rowNum, c: colNum });
  return worksheet[cellIdentity] ? worksheet[cellIdentity].v : '';
};

const isEndOfData = (worksheet, rowNum) => {
  const cellData = readCell(worksheet, rowNum, 2);
  if (cellData === 'TOTAL') {
    return true;
  }
  return false;
};

exports.readBad = function (fileName, callback) {
  const workbook = XLSX.readFile(fileName);
  const worksheet = workbook.Sheets[workbook.SheetNames[BAD_SHEET_POSITION]];

  // const yearCellValue = worksheet.F6.v;
  const year = 2017;
  const month = 1;
  const bads = [];
  const startRow = 27;
  for (let i = startRow; i < (startRow + MAXIMUM_ROW); i += 1) {
    if (isEndOfData(worksheet, i)) {
      break;
    }
    const projectCode = readCell(worksheet, i, 0);
    if (projectCode) {
      const piutangUsaha = readCell(worksheet, i, 2);
      const tagihanBruto = readCell(worksheet, i, 3);
      const piutangRetensi = readCell(worksheet, i, 4);
      const pdp = readCell(worksheet, i, 5);
      const bad = readCell(worksheet, i, 6);
      bads.push({
        projectCode,
        piutangUsaha,
        tagihanBruto,
        piutangRetensi,
        pdp,
        bad,
        year,
        month
      });
    }
  }

  callback({ payload: bads });
};
