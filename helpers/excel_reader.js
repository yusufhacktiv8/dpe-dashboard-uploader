const XLSX = require('xlsx');
const _ = require('lodash');

const INPUTAN_SHEET_POSITION = 0;
const RINCIAN_SHEET_POSITION = 1;

exports.readProjectProgress = function (fileName, year, callback) {
  const workbook = XLSX.readFile(fileName);
  const worksheet = workbook.Sheets[workbook.SheetNames[INPUTAN_SHEET_POSITION]];

  const yearCellValue = worksheet.F6.v;
  const YEAR = year;

  const projectProgresses = [];

  const colNames = ['I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
    'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR'
  ];

  const PROJECT_TYPES = [
    {
      id: 1,
      title: 'PROYEK LAMA NON JO/NON KSO',
    },
    {
      id: 2,
      title: 'PROYEK LAMA JO/KSO',
    },
    {
      id: 3,
      title: 'PROYEK LAMA INTERN',
    },
    {
      id: 4,
      title: 'PROYEK BARU SUDAH DIPEROLEH NON JO/NON KSO',
    },
    {
      id: 5,
      title: 'PROYEK BARU SUDAH DIPEROLEH JO/KSO',
    },
    {
      id: 6,
      title: 'PROYEK BARU SUDAH DIPEROLEH INTERN',
    },
    {
      id: 7,
      title: 'PROYEK BARU DALAM PENGUSAHAAN NON JO/NON KSO',
    },
    {
      id: 8,
      title: 'PROYEK BARU DALAM PENGUSAHAAN JO/KSO',
    },
    {
      id: 9,
      title: 'PROYEK BARU DALAM PENGUSAHAAN INTERN',
    },
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

  const fillAndOrderProjectTypesRow = (worksheet) => {
    const result = [];
    for (let row = 10; row < 500; row += 1) {
      let projectCode = worksheet[`C${row}`] ? worksheet[`C${row}`].v : '';
      let details = worksheet[`D${row}`] ? worksheet[`D${row}`].v : '';
      if (projectCode === 'END') {
        break;
      }
      const projectType = _.find(PROJECT_TYPES, { title: details });
      if (projectType) {
        result.push({ id: projectType.id, row});
      }
    }
    return _.orderBy(result, ['row'], ['desc']);
  }

  const getProjectTypeIdByRow = (projectTypes, row) => {
    for (let i = 0; i < projectTypes.length; i += 1) {
      const projectType = projectTypes[i];
      if (row > projectType.row) {
        return projectType.id;
      }
    }
    return null;
  }

  const orderedProjectTypes = fillAndOrderProjectTypesRow(worksheet);
  // console.log('orderedProjectTypes: ', orderedProjectTypes);

  for (let row = 11; row < 500; row += 1) {
    let projectCode = worksheet['C' + row] ? worksheet['C' + row].v : '';

    if (projectCode != '') {

      if (projectCode == 'END') {
        break;
      }

      projectCode = projectCode.trim();

      for (let month = 1; month <= 12; month += 1) {
        const tmpProjectProgress = getProjectProgress(projectCode, row, month, YEAR, worksheet);
        tmpProjectProgress.projectType = getProjectTypeIdByRow(orderedProjectTypes, row);
        projectProgresses.push(tmpProjectProgress);
      }
      row += 2;
    }
  }

  console.log('projectProgresses: =======> ', projectProgresses);

  callback({ year: YEAR, projectProgresses });
};

exports.readLsp = function (fileName, year, callback) {
  var workbook = XLSX.readFile(fileName);

  var firstSheetName = workbook.SheetNames[RINCIAN_SHEET_POSITION];
  var worksheet = workbook.Sheets[firstSheetName];

  const YEAR = year;

  var labaSetelahPajak = [];

  var getData = function(cellName, ws){

    return ws[cellName]? ws[cellName].v : 0;
  }

  for (let row = 8; row < 150; row += 1) {
    const name = worksheet['C' + row] ? worksheet['C' + row].v : '';

    if(name == 'Laba Setelah Pajak'){
      labaSetelahPajak.push({
        month: 1,
        year: YEAR,
        lsp_rkap: getData(('J' + row), worksheet),
        lsp_prognosa: getData(('M' + row), worksheet),
        lsp_realisasi: getData(('P' + row), worksheet),
      });

      labaSetelahPajak.push({
        month: 2,
        year: YEAR,
        lsp_rkap: getData(('S' + row), worksheet),
        lsp_prognosa: getData(('V' + row), worksheet),
        lsp_realisasi: getData(('Y' + row), worksheet),
      });

      labaSetelahPajak.push({
        month: 3,
        year: YEAR,
        lsp_rkap: getData(('AB' + row), worksheet),
        lsp_prognosa: getData(('AE' + row), worksheet),
        lsp_realisasi: getData(('AH' + row), worksheet),
      });

      labaSetelahPajak.push({
        month: 4,
        year: YEAR,
        lsp_rkap: getData(('AK' + row), worksheet),
        lsp_prognosa: getData(('AN' + row), worksheet),
        lsp_realisasi: getData(('AQ' + row), worksheet),
      });

      labaSetelahPajak.push({
        month: 5,
        year: YEAR,
        lsp_rkap: getData(('AT' + row), worksheet),
        lsp_prognosa: getData(('AW' + row), worksheet),
        lsp_realisasi: getData(('AZ' + row), worksheet),
      });

      labaSetelahPajak.push({
        month: 6,
        year: YEAR,
        lsp_rkap: getData(('BC' + row), worksheet),
        lsp_prognosa: getData(('BF' + row), worksheet),
        lsp_realisasi: getData(('BI' + row), worksheet),
      });

      labaSetelahPajak.push({
        month: 7,
        year: YEAR,
        lsp_rkap: getData(('BL' + row), worksheet),
        lsp_prognosa: getData(('BO' + row), worksheet),
        lsp_realisasi: getData(('BR' + row), worksheet),
      });

      labaSetelahPajak.push({
        month: 8,
        year: YEAR,
        lsp_rkap: getData(('BU' + row), worksheet),
        lsp_prognosa: getData(('BX' + row), worksheet),
        lsp_realisasi: getData(('CA' + row), worksheet),
      });

      labaSetelahPajak.push({
        month: 9,
        year: YEAR,
        lsp_rkap: getData(('CD' + row), worksheet),
        lsp_prognosa: getData(('CG' + row), worksheet),
        lsp_realisasi: getData(('CJ' + row), worksheet),
      });

      labaSetelahPajak.push({
        month: 10,
        year: YEAR,
        lsp_rkap: getData(('CM' + row), worksheet),
        lsp_prognosa: getData(('CP' + row), worksheet),
        lsp_realisasi: getData(('CS' + row), worksheet),
      });

      labaSetelahPajak.push({
        month: 11,
        year: YEAR,
        lsp_rkap: getData(('CV' + row), worksheet),
        lsp_prognosa: getData(('CY' + row), worksheet),
        lsp_realisasi: getData(('DB' + row), worksheet),
      });

      labaSetelahPajak.push({
        month: 12,
        year: YEAR,
        lsp_rkap: getData(('DE' + row), worksheet),
        lsp_prognosa: getData(('DH' + row), worksheet),
        lsp_realisasi: getData(('DK' + row), worksheet),
      });
      break;
    }
  }

  callback({ year: YEAR, labaSetelahPajak });
};

exports.readClaim = function (fileName, year, callback) {
  const workbook = XLSX.readFile(fileName);

  const firstSheetName = workbook.SheetNames[RINCIAN_SHEET_POSITION];
  const worksheet = workbook.Sheets[firstSheetName];

  const getData = (cellName, ws) => (ws[cellName] ? ws[cellName].v : 0);
  const ok = getData('DV10', worksheet);
  const op = getData('DW10', worksheet);
  const lk = getData('DX10', worksheet);

  callback({ year, payload: { ok, op, lk } });
}
