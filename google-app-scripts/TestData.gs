/**
 * Populates FIFO, Crypto, and Dividends sheets with sample test data.
 * Covers multiple countries to verify country-grouped report calculations.
 * Run from the Apps Script editor or add to a custom menu.
 */
function insertTestData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  insertFifoTestData(spreadsheet);
  insertCryptoTestData(spreadsheet);
  insertDividendsTestData(spreadsheet);

  SpreadsheetApp.getUi().alert('Test data inserted');
}

function insertFifoTestData(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(FIFO_SHEET_NAME);
  // Columns B through L (2-12)
  const data = [
    // USA stocks — buy then sell
    ['TEAM', 'Inna', RSU_COUNTRY,                     BUY_OPERATION,  new Date('2025-01-15'), 10, 200,   '', 'USD', 5,  'USD'],
    ['TEAM', 'Inna', RSU_COUNTRY,                     SELL_OPERATION, new Date('2025-06-10'), 10, 250,   '', 'USD', 5,  'USD'],
    // Polish stocks — buy then sell
    ['CDR',  'Inna', 'Polska',                        BUY_OPERATION,  new Date('2025-02-01'), 20, 150,   '', CURRENCIES.PLN, 10, CURRENCIES.PLN],
    ['CDR',  'Inna', 'Polska',                        SELL_OPERATION, new Date('2025-07-15'), 20, 180,   '', CURRENCIES.PLN, 10, CURRENCIES.PLN],
    // German stocks — buy then sell
    ['SAP',  'Inna', 'Niemcy',                        BUY_OPERATION,  new Date('2025-03-01'), 5,  120,   '', 'EUR', 8,  'EUR'],
    ['SAP',  'Inna', 'Niemcy',                        SELL_OPERATION, new Date('2025-08-20'), 5,  140,   '', 'EUR', 8,  'EUR'],
  ];

  const startRow = findFirstEmptyRow(sheet, FIFO_COL.symbol.index);
  sheet.getRange(startRow, 2, data.length, data[0].length).setValues(data);

  // Set transactionSum formulas (col I = G * H)
  for (let i = 0; i < data.length; i++) {
    const rowNum = startRow + i;
    sheet.getRange('I' + rowNum).setFormula('=G' + rowNum + '*H' + rowNum);
  }
}

function insertCryptoTestData(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(CRYPTO_SHEET_NAME);
  // Columns B through L (2-12)
  const data = [
    // USA — BTC buy & sell
    ['BTC',  'Inna', RSU_COUNTRY,  BUY_OPERATION,  new Date('2025-01-20'), 0.5,   40000, '', 'USD', 15,  'USD'],
    ['BTC',  'Inna', RSU_COUNTRY,  SELL_OPERATION, new Date('2025-09-15'), 0.5,   55000, '', 'USD', 15,  'USD'],
    // USA — SOL buy & sell
    ['SOL',  'Inna', RSU_COUNTRY,  BUY_OPERATION,  new Date('2025-03-05'), 100,   95,    '', 'USD', 10,  'USD'],
    ['SOL',  'Inna', RSU_COUNTRY,  SELL_OPERATION, new Date('2025-11-20'), 100,   180,   '', 'USD', 10,  'USD'],
    // Poland — ETH buy & sell
    ['ETH',  'Inna', 'Polska',     BUY_OPERATION,  new Date('2025-02-10'), 2,     10000, '', CURRENCIES.PLN, 20, CURRENCIES.PLN],
    ['ETH',  'Inna', 'Polska',     SELL_OPERATION, new Date('2025-10-05'), 2,     12000, '', CURRENCIES.PLN, 20, CURRENCIES.PLN],
    // Poland — XRP buy & sell
    ['XRP',  'Inna', 'Polska',     BUY_OPERATION,  new Date('2025-04-12'), 5000,  2.5,   '', CURRENCIES.PLN, 15, CURRENCIES.PLN],
    ['XRP',  'Inna', 'Polska',     SELL_OPERATION, new Date('2025-12-01'), 5000,  4.8,   '', CURRENCIES.PLN, 15, CURRENCIES.PLN],
    // Germany — ADA buy & sell
    ['ADA',  'Inna', 'Niemcy',     BUY_OPERATION,  new Date('2025-05-18'), 8000,  0.45,  '', 'EUR', 12,  'EUR'],
    ['ADA',  'Inna', 'Niemcy',     SELL_OPERATION, new Date('2025-10-30'), 8000,  0.72,  '', 'EUR', 12,  'EUR'],
    // Germany — DOT buy & sell
    ['DOT',  'Inna', 'Niemcy',     BUY_OPERATION,  new Date('2025-06-01'), 300,   6.2,   '', 'EUR', 8,   'EUR'],
    ['DOT',  'Inna', 'Niemcy',     SELL_OPERATION, new Date('2025-11-15'), 300,   9.5,   '', 'EUR', 8,   'EUR'],
  ];

  const startRow = findFirstEmptyRow(sheet, CRYPTO_COL.operationType.index);
  sheet.getRange(startRow, 2, data.length, data[0].length).setValues(data);

  for (let i = 0; i < data.length; i++) {
    const rowNum = startRow + i;
    sheet.getRange('I' + rowNum).setFormula('=G' + rowNum + '*H' + rowNum);
  }
}

function insertDividendsTestData(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(DIVIDENDS_SHEET_NAME);
  // Columns B through I (2-9): Symbol, Gielda, Kraj, Data, Kwota, Waluta, Prowizja, Waluta Prowizji
  const data = [
    //  B        C            D                    E                        F     G      H   I
    // USA dividends
    ['TEAM', 'Inna', RSU_COUNTRY,  new Date('2025-03-15'), 500,  'USD', 0,  'USD'],
    ['AAPL', 'Inna', RSU_COUNTRY,  new Date('2025-06-15'), 300,  'USD', 0,  'USD'],
    ['MSFT', 'Inna', RSU_COUNTRY,  new Date('2025-09-10'), 750,  'USD', 0,  'USD'],
    ['GOOG', 'Inna', RSU_COUNTRY,  new Date('2025-12-05'), 180,  'USD', 0,  'USD'],
    // German dividends
    ['SAP',  'Inna', 'Niemcy',     new Date('2025-05-20'), 200,  'EUR', 0,  'EUR'],
    ['BMW',  'Inna', 'Niemcy',     new Date('2025-07-10'), 450,  'EUR', 0,  'EUR'],
    ['SIE',  'Inna', 'Niemcy',     new Date('2025-11-25'), 320,  'EUR', 0,  'EUR'],
    // Polish dividends
    ['CDR',  'Inna', 'Polska',     new Date('2025-04-01'), 1000, CURRENCIES.PLN, 0, CURRENCIES.PLN],
    ['PKN',  'Inna', 'Polska',     new Date('2025-08-15'), 650,  CURRENCIES.PLN, 0, CURRENCIES.PLN],
    ['PZU',  'Inna', 'Polska',     new Date('2025-10-20'), 420,  CURRENCIES.PLN, 0, CURRENCIES.PLN],
  ];

  const startRow = findFirstEmptyRow(sheet, DIVIDENDS_COL.transactionDate.index);
  sheet.getRange(startRow, 2, data.length, data[0].length).setValues(data);
}

function findFirstEmptyRow(sheet, checkColIndex) {
  const allData = sheet.getDataRange().getValues();
  for (let i = allData.length - 1; i >= 1; i--) {
    if (allData[i][checkColIndex]) {
      return i + 2; // 1-based + next row
    }
  }
  return 2; // first data row after header
}
