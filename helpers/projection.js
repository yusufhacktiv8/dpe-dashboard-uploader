const XLSX = require('xlsx');

const PROJECTION_SHEET_POSITION = 0;

const readCell = (worksheet, rowNum, colNum) => {
  const cellIdentity = XLSX.utils.encode_cell({ r: rowNum, c: colNum });
  return worksheet[cellIdentity] ? worksheet[cellIdentity].v : '';
};

const readProjectionsDataInAYear = (worksheet, year, startRow) => {
  const projectionDataInAYear = [];
  const initialColumnPosition = 2;
  for (let month = 1; month <= 12; month += 1) {
    const columnPosition = initialColumnPosition + (month - 1);
    const projectionData = {
      pdp: readCell(worksheet, startRow, columnPosition) || 0,
      tagihanBruto: readCell(worksheet, startRow + 1, columnPosition) || 0,
      piutangUsaha: readCell(worksheet, startRow + 2, columnPosition) || 0,
      piutangRetensi: readCell(worksheet, startRow + 3, columnPosition) || 0,
      month,
      year,
    };
    projectionDataInAYear.push(projectionData);
  }
  return projectionDataInAYear;
};

exports.read = function (fileName, year, callback) {
  const workbook = XLSX.readFile(fileName);
  const worksheet = workbook.Sheets[workbook.SheetNames[PROJECTION_SHEET_POSITION]];

  const projectionsDataInAYear = readProjectionsDataInAYear(worksheet, year, 4);

  callback({ year, payload: projectionsDataInAYear });
};
