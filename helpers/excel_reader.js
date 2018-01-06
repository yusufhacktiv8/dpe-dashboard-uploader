const XLSX = require('xlsx');

const INPUTAN_SHEET_POSITION = 0;
const RINCIAN_SHEET_POSITION = 1;

exports.readProjectProgress = function (fileName, callback) {
  const workbook = XLSX.readFile(fileName);
  const worksheet = workbook.Sheets[workbook.SheetNames[INPUTAN_SHEET_POSITION]];

  const yearCellValue = worksheet.F6.v;
  const YEAR = 2017; // parseInt(yearCellValue.match(/[0-9]+/)[0], 10);

  const projectProgresses = [];

  const colNames = ['I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
    'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR'
  ];

  const getProjectProgress = (projectCode, row, month, year) => {
    let cellMonthPositionInit = (month - 1) * 3;
    let cellName = colNames[cellMonthPositionInit] + row;
    const rkapOk = worksheet[cellName] ? worksheet[cellName].v : 0;
    cellName = colNames[cellMonthPositionInit] + (row + 1);
    const prognosaOk = worksheet[cellName] ? worksheet[cellName].v : 0;
    cellName = colNames[cellMonthPositionInit] + (row + 2);
    const realisasiOk = worksheet[cellName] ? worksheet[cellName].v : 0;

    cellMonthPositionInit += 1;
    cellName = colNames[cellMonthPositionInit] + row;
    const rkapOp = worksheet[cellName] ? worksheet[cellName].v : 0;
    cellName = colNames[cellMonthPositionInit] + (row + 1);
    const prognosaOp = worksheet[cellName] ? worksheet[cellName].v : 0;
    cellName = colNames[cellMonthPositionInit] + (row + 2);
    const realisasiOp = worksheet[cellName] ? worksheet[cellName].v : 0;

    cellMonthPositionInit += 1;
    cellName = colNames[cellMonthPositionInit] + row;
    const rkapLk = worksheet[cellName] ? worksheet[cellName].v : 0;
    cellName = colNames[cellMonthPositionInit] + (row + 1);
    const prognosaLk = worksheet[cellName] ? worksheet[cellName].v : 0;
    cellName = colNames[cellMonthPositionInit] + (row + 2);
    const realisasiLk = worksheet[cellName] ? worksheet[cellName].v : 0;

    const projectProgress = {
      month,
      year,
      projectCode,
      rkapOk,
      rkapOp,
      rkapLk,
      prognosaOk,
      prognosaOp,
      prognosaLk,
      realisasiOk,
      realisasiOp,
      realisasiLk,
    };

    return projectProgress;
  };

  for (let row = 11; row < 500; row += 1) {
    let projectCode = worksheet['C' + row] ? worksheet['C' + row].v : '';

    if (projectCode != '') {

      if (projectCode == 'END') {
        break;
      }

      projectCode = projectCode.trim();

      for (let month = 1; month <= 12; month += 1) {
        const tmpProjectProgress = getProjectProgress(projectCode, row, month, YEAR, worksheet);
        projectProgresses.push(tmpProjectProgress);
      }
      row += 2;
    }
  }

  callback(projectProgresses);
};
