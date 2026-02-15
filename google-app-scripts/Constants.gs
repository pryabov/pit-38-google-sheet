// Sheet names
var FIFO_SHEET_NAME = 'FIFO Stocks Transactions';
var CRYPTO_SHEET_NAME = 'Crypto Currencies';
var DIVIDENDS_SHEET_NAME = 'Dividends';
var SETTINGS_SHEET_NAME = 'Settings';
var REPORT_SHEET_NAME = 'Report';
var CALC_LOG_SHEET_NAME = 'Calculation Log';
var HOME_PAGE_SHEET_NAME = 'Home Page';

// Currencies
var CURRENCIES = {
  PLN: 'PLN'
};

// Operation types
var BUY_OPERATION = 'Kupowanie';
var SELL_OPERATION = 'Sprzedaż';

// RSU upload defaults
var RSU_SYMBOL = 'TEAM';
var RSU_STOCK_TYPE = 'Inna';
var RSU_COUNTRY = 'Stany Zjednoczone Ameryki';

// Report sheet layout
var REPORT_FIFO_START_ROW = 4;  // FIFO per-country rows start here
var REPORT_FIFO_COUNTRY_COL = 'A';
var REPORT_FIFO_REVENUE_COL = 'B';
var REPORT_FIFO_COST_COL = 'C';

// Column mappings for batch-read arrays (.index) and getRange calls (.letter).
// Usage: row[FIFO_COL.currency.index] or sheet.getRange(FIFO_COL.currency.letter + row)
var FIFO_COL = {
  symbol:          { index: 1,  letter: 'B' },
  country:         { index: 3,  letter: 'D' },
  operationType:   { index: 4,  letter: 'E' },
  transactionDate: { index: 5,  letter: 'F' },
  count:           { index: 6,  letter: 'G' },
  price:           { index: 7,  letter: 'H' },
  transactionSum:  { index: 8,  letter: 'I' },
  currency:        { index: 9,  letter: 'J' },
  costs:           { index: 10, letter: 'K' },
  nbpRateDate:     { index: 12, letter: 'M' },
  exchangeRate:    { index: 13, letter: 'N' },
  sumPLN:          { index: 14, letter: 'O' },
  costsPLN:        { index: 15, letter: 'P' }
};

var CRYPTO_COL = {
  country:         { index: 3,  letter: 'D' },
  operationType:   { index: 4,  letter: 'E' },
  transactionDate: { index: 5,  letter: 'F' },
  transactionSum:  { index: 8,  letter: 'I' },
  currency:        { index: 9,  letter: 'J' },
  costs:           { index: 10, letter: 'K' },
  nbpRateDate:     { index: 12, letter: 'M' },
  exchangeRate:    { index: 13, letter: 'N' },
  sumPLN:          { index: 14, letter: 'O' },
  costsPLN:        { index: 15, letter: 'P' }
};

var DIVIDENDS_COL = {
  country:         { index: 3,  letter: 'D' },
  transactionDate: { index: 4,  letter: 'E' },
  transactionSum:  { index: 5,  letter: 'F' },
  currency:        { index: 6,  letter: 'G' },
  costs:           { index: 7,  letter: 'H' },
  nbpRateDate:     { index: 9,  letter: 'J' },
  exchangeRate:    { index: 10, letter: 'K' },
  sumPLN:          { index: 11, letter: 'L' },
  costsPLN:        { index: 12, letter: 'M' }
};
