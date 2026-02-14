/**
 * Poland Stocks Tax Calculator - NBP Exchange Rates Module (Optimized)
 * 
 * Fetches and caches NBP (National Bank of Poland) Table A exchange rates.
 * Supports USD, EUR, GBP, and CHF currencies.
 */

const NBP_CACHE_KEY = 'nbpRates';
const NBP_CACHE_TTL = 21600; // 6 hours (maximum allowed by CacheService)
const NBP_YEARS_TO_FETCH = 3;

// Column indices in NBP Table A CSV (0-based, after date column)
const CURRENCY_COLUMNS = {
  usd: 2,  // Column C
  eur: 8,  // Column I (adjust based on actual CSV layout)
  gbp: 10, // Column K (adjust based on actual CSV layout)
  chf: 6   // Column G (adjust based on actual CSV layout)
};

/**
 * Main entry point: imports NBP rates for the current and previous years.
 */
function importNbpRates() {
  const currentYear = new Date().getFullYear();

  for (let i = 0; i < NBP_YEARS_TO_FETCH; i++) {
    importNbpRatesForYear(currentYear - i);
  }

  refreshNbpRatesCache();
}

/**
 * Imports NBP Table A CSV data into a dedicated sheet for the given year.
 * Creates the sheet if it doesn't exist, clears it if it does.
 * 
 * @param {number} year - The year to import rates for.
 */
function importNbpRatesForYear(year) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'NBP_Rates_' + year;
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  } else {
    sheet.clear();
  }

  sheet.getRange('A1').setFormula(
    '=IMPORTDATA("https://static.nbp.pl/dane/kursy/Archiwum/archiwum_tab_a_' + year + '.csv",";","PL")'
  );
}

/**
 * Retrieves NBP rates from cache or refreshes if cache is empty/expired.
 * 
 * @returns {Map<string, Object>} Map of date strings to currency rate objects.
 */
function getNbpRates() {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get(NBP_CACHE_KEY);

  if (cachedData && cachedData !== '[]') {
    return new Map(JSON.parse(cachedData));
  }

  return refreshNbpRatesCache();
}

/**
 * Rebuilds the NBP rates cache from spreadsheet data.
 * Reads all years and stores the combined result in CacheService.
 * 
 * @returns {Map<string, Object>} Map of date strings to currency rate objects.
 */
function refreshNbpRatesCache() {
  const cache = CacheService.getScriptCache();
  const nbpRates = new Map();
  const currentYear = new Date().getFullYear();

  for (let i = 0; i < NBP_YEARS_TO_FETCH; i++) {
    populateRatesToMap(nbpRates, currentYear - i);
  }

  const serialized = JSON.stringify(Array.from(nbpRates.entries()));

  // CacheService has a 100KB value limit. If data exceeds it, chunk or skip caching.
  if (serialized.length < 100000) {
    cache.put(NBP_CACHE_KEY, serialized, NBP_CACHE_TTL);
  } else {
    Logger.log('Warning: NBP rates data exceeds cache size limit. Caching skipped.');
  }

  return nbpRates;
}

/**
 * Reads exchange rates from a yearly NBP sheet and adds them to the rates map.
 * Uses a single batch read of the entire sheet to minimize API calls.
 *
 * @param {Map<string, Object>} nbpRates - Map to populate with rates.
 * @param {number} year - The year sheet to read from.
 */
function populateRatesToMap(nbpRates, year) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('NBP_Rates_' + year);

  if (!sheet) {
    Logger.log('Sheet NBP_Rates_' + year + ' not found. Skipping.');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    Logger.log('Sheet NBP_Rates_' + year + ' has no data rows. Skipping.');
    return;
  }

  // Read ALL data in one batch call (rows 3 to lastRow, all used columns)
  const lastCol = sheet.getLastColumn();
  const dataRange = sheet.getRange(3, 1, lastRow - 2, lastCol);
  const allData = dataRange.getValues();

  for (let row = 0; row < allData.length; row++) {
    const dateValue = convertToDateFormat(allData[row][0]);
    if (!dateValue) continue;

    const rates = {};
    for (const [currency, colIndex] of Object.entries(CURRENCY_COLUMNS)) {
      const value = allData[row][colIndex];
      if (value !== '' && value !== null && value !== undefined) {
        rates[currency] = typeof value === 'string' ? parsePolishDecimal(value) : value;
      }
    }

    nbpRates.set(dateValue, rates);
  }
}

/**
 * Parses a Polish-format decimal string (comma as decimal separator).
 * 
 * @param {string} value - Polish decimal string, e.g. "4,0325".
 * @returns {number} Parsed float value.
 */
function parsePolishDecimal(value) {
  return parseFloat(value.toString().replace(',', '.'));
}

/**
 * Converts a date input (string or Date object) to "YYYY-MM-DD" format.
 * Accepts formats: "yyyyMMdd" strings and Date objects.
 * 
 * @param {string|Date} input - The date to convert.
 * @returns {string|null} Formatted date string or null if invalid.
 */
function convertToDateFormat(input) {
  if (input instanceof Date) {
    if (isNaN(input.getTime())) return null;
    const y = input.getFullYear();
    const m = input.getMonth() + 1;
    const d = input.getDate();
    return `${y.toString().padStart(4, '0')}-${m.toString().padStart(2, '0')}-${d.toString().padStart(2, '0')}`;
  }

  const inputString = input.toString().trim();
  if (!/^\d{8}$/.test(inputString)) return null;

  const year = parseInt(inputString.substring(0, 4), 10);
  const month = parseInt(inputString.substring(4, 6), 10) - 1;
  const day = parseInt(inputString.substring(6, 8), 10);
  const date = new Date(year, month, day);

  if (date.getFullYear() !== year || date.getMonth() !== month || date.getDate() !== day) {
    return null;
  }

  return `${year.toString().padStart(4, '0')}-${(month + 1).toString().padStart(2, '0')}-${day.toString().padStart(2, '0')}`;
}