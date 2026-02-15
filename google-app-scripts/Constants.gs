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

// Tax rate
var TAX_RATE = 0.19;

// Report sheet layout — three side-by-side tables
var REPORT_HEADER_ROW = 1;
var REPORT_DATA_START_ROW = 2;
var REPORT_HEADER_COLOR = '#4472C4';
var REPORT_HEADER_FONT_COLOR = '#FFFFFF';
var REPORT_COL_LABELS = { country: 'Kraj', revenue: 'Przychód', cost: 'Koszt', tax: 'Podatek' };
var REPORT_SUM_LABEL = 'Suma';

var REPORT_FIFO    = { label: 'Akcje (FIFO)',   countryCol: 'A', revenueCol: 'B', costCol: 'C', taxCol: 'D' };
var REPORT_CRYPTO  = { label: 'Kryptowaluty',   countryCol: 'F', revenueCol: 'G', costCol: 'H', taxCol: 'I' };
var REPORT_DIV     = { label: 'Dywidendy',      countryCol: 'K', revenueCol: 'L', costCol: 'M', taxCol: 'N' };

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
