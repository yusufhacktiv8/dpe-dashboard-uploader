const XLSX = require('xlsx');

const CASH_FLOW_SHEET_POSITION = 1;

const readCell = (worksheet, rowNum, colNum) => {
  const cellIdentity = XLSX.utils.encode_cell({ r: rowNum, c: colNum });
  return worksheet[cellIdentity] ? worksheet[cellIdentity].v : '';
};

const readCashFlowItems = (worksheet, startRow) => {
  const cashFlowItems = [];
  const initialColumnPosition = 21;
  for (let month = 1; month <= 12; month += 1) {
    const columnPosition = initialColumnPosition + ((month - 1) * 3);
    const projectData = {
      ra: readCell(worksheet, startRow, columnPosition) || 0,
      prog: readCell(worksheet, startRow, columnPosition + 1) || 0,
      ri: readCell(worksheet, startRow, columnPosition + 1) || 0,
      month,
    };
    cashFlowItems.push(projectData);
  }
  return cashFlowItems;
};

const readAllCashFlowDataInAYear = (worksheet, year) => {
  const allCashFlowDataInAYear = [];
  const saldoAwal = {
    typeCode: 1,
    year,
    rkap: readCell(worksheet, 11, 19) || 0,
    rolling: readCell(worksheet, 11, 20) || 0,
    items: readCashFlowItems(worksheet, 11),
  };
  const penerimaan = {
    typeCode: 2,
    year,
    rkap: readCell(worksheet, 15, 19) || 0,
    rolling: readCell(worksheet, 15, 20) || 0,
    items: readCashFlowItems(worksheet, 15),
  };
  const pengeluaran = {
    typeCode: 3,
    year,
    rkap: readCell(worksheet, 22, 19) || 0,
    rolling: readCell(worksheet, 22, 20) || 0,
    items: readCashFlowItems(worksheet, 22),
  };
  const kelebihanKas = {
    typeCode: 4,
    year,
    rkap: readCell(worksheet, 28, 19) || 0,
    rolling: readCell(worksheet, 28, 20) || 0,
    items: readCashFlowItems(worksheet, 28),
  };
  const setoran = {
    typeCode: 5,
    year,
    rkap: readCell(worksheet, 36, 19) || 0,
    rolling: readCell(worksheet, 36, 20) || 0,
    items: readCashFlowItems(worksheet, 36),
  };
  const saldoKasAkhir = {
    typeCode: 6,
    year,
    rkap: readCell(worksheet, 38, 19) || 0,
    rolling: readCell(worksheet, 38, 20) || 0,
    items: readCashFlowItems(worksheet, 38),
  };
  const saldoAwalRK = {
    typeCode: 7,
    year,
    rkap: readCell(worksheet, 40, 19) || 0,
    rolling: readCell(worksheet, 40, 20) || 0,
    items: readCashFlowItems(worksheet, 40),
  };
  const jumlahMutasi = {
    typeCode: 8,
    year,
    rkap: readCell(worksheet, 46, 19) || 0,
    rolling: readCell(worksheet, 46, 20) || 0,
    items: readCashFlowItems(worksheet, 46),
  };
  const jumlahRK = {
    typeCode: 9,
    year,
    rkap: readCell(worksheet, 47, 19) || 0,
    rolling: readCell(worksheet, 47, 20) || 0,
    items: readCashFlowItems(worksheet, 47),
  };
  const jumlahPemakaianDana = {
    typeCode: 10,
    year,
    rkap: readCell(worksheet, 48, 19) || 0,
    rolling: readCell(worksheet, 48, 20) || 0,
    items: readCashFlowItems(worksheet, 48),
  };
  const totalBunga = {
    typeCode: 11,
    year,
    rkap: readCell(worksheet, 52, 19) || 0,
    rolling: readCell(worksheet, 52, 20) || 0,
    items: readCashFlowItems(worksheet, 52),
  };
  const saldoAkhir = {
    typeCode: 12,
    year,
    rkap: readCell(worksheet, 53, 19) || 0,
    rolling: readCell(worksheet, 53, 20) || 0,
    items: readCashFlowItems(worksheet, 53),
  };
  allCashFlowDataInAYear.push(saldoAwal);
  allCashFlowDataInAYear.push(penerimaan);
  allCashFlowDataInAYear.push(pengeluaran);
  allCashFlowDataInAYear.push(kelebihanKas);
  allCashFlowDataInAYear.push(setoran);
  allCashFlowDataInAYear.push(saldoKasAkhir);
  allCashFlowDataInAYear.push(saldoAwalRK);
  allCashFlowDataInAYear.push(jumlahMutasi);
  allCashFlowDataInAYear.push(jumlahRK);
  allCashFlowDataInAYear.push(jumlahPemakaianDana);
  allCashFlowDataInAYear.push(totalBunga);
  allCashFlowDataInAYear.push(saldoAkhir);
  return allCashFlowDataInAYear;
};

exports.read = function (fileName, callback) {
  const workbook = XLSX.readFile(fileName);
  const worksheet = workbook.Sheets[workbook.SheetNames[CASH_FLOW_SHEET_POSITION]];
  // const yearCellValue = worksheet.F6.v;
  const year = 2017;
  const allCashFlowDataInAYear = readAllCashFlowDataInAYear(worksheet, year);
  callback({ year, payload: allCashFlowDataInAYear });
};
