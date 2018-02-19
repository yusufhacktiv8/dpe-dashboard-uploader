const XLSX = require('xlsx');

const PIUTANG_SHEET_POSITION = 0;

const readCell = (worksheet, rowNum, colNum) => {
  const cellIdentity = XLSX.utils.encode_cell({ r: rowNum, c: colNum });
  return worksheet[cellIdentity] ? worksheet[cellIdentity].v : '';
};

const readProjectDataInAYear = (worksheet, year, startRow) => {
  const projectDataInAYear = [];
  const initialColumnPosition = 8;
  for (let month = 1; month <= 12; month += 1) {
    const columnPosition = initialColumnPosition + ((month - 1) * 6);
    const projectData = {
      owner: readCell(worksheet, startRow, 3) || 0,
      projectCode: readCell(worksheet, startRow, 6),
      pdp1: readCell(worksheet, startRow, columnPosition) || 0,
      tagihanBruto1: readCell(worksheet, startRow + 1, columnPosition) || 0,
      piutangUsaha1: readCell(worksheet, startRow + 2, columnPosition) || 0,
      piutangRetensi1: readCell(worksheet, startRow + 3, columnPosition) || 0,
      pdp2: readCell(worksheet, startRow, columnPosition + 1) || 0,
      tagihanBruto2: readCell(worksheet, startRow + 1, columnPosition + 1) || 0,
      piutangUsaha2: readCell(worksheet, startRow + 2, columnPosition + 1) || 0,
      piutangRetensi2: readCell(worksheet, startRow + 3, columnPosition + 1) || 0,
      pdp3: readCell(worksheet, startRow, columnPosition + 2) || 0,
      tagihanBruto3: readCell(worksheet, startRow + 1, columnPosition + 2) || 0,
      piutangUsaha3: readCell(worksheet, startRow + 2, columnPosition + 2) || 0,
      piutangRetensi3: readCell(worksheet, startRow + 3, columnPosition + 2) || 0,
      pdp4: readCell(worksheet, startRow, columnPosition + 3) || 0,
      tagihanBruto4: readCell(worksheet, startRow + 1, columnPosition + 3) || 0,
      piutangUsaha4: readCell(worksheet, startRow + 2, columnPosition + 3) || 0,
      piutangRetensi4: readCell(worksheet, startRow + 3, columnPosition + 3) || 0,
      pdp5: readCell(worksheet, startRow, columnPosition + 4) || 0,
      tagihanBruto5: readCell(worksheet, startRow + 1, columnPosition + 4) || 0,
      piutangUsaha5: readCell(worksheet, startRow + 2, columnPosition + 4) || 0,
      piutangRetensi5: readCell(worksheet, startRow + 3, columnPosition + 4) || 0,
      month,
      year,
    };
    projectDataInAYear.push(projectData);
  }
  return projectDataInAYear;
};

const readAllProjectsDataInAYear = (worksheet, year) => {
  const allProjectsDataInAYear = [];
  for (let row = 4; row <= 29; row += 5) {
    if (readCell(worksheet, row, 0)) {
      allProjectsDataInAYear.push(...readProjectDataInAYear(worksheet, year, row));
    }
  }
  return allProjectsDataInAYear;
};

exports.read = function (fileName, year, callback) {
  const workbook = XLSX.readFile(fileName);
  const worksheet = workbook.Sheets[workbook.SheetNames[PIUTANG_SHEET_POSITION]];
  const allProjectsDataInAYear = readAllProjectsDataInAYear(worksheet, year);

  callback({ year, payload: allProjectsDataInAYear });
}
