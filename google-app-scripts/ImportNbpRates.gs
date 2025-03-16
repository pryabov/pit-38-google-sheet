function importNbpRates() {
  let currentYear = new Date().getFullYear();

  for (let i = 0; i < 3; i++){
    importNbpRatesForYear(currentYear - i);
  }

  refreshNbpRatesCache();
}

function importNbpRatesForYear(year) {
  // Open the spreadsheet and get the sheet named "nbp"
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('NBP_Rates_' + year);
  if (!sheet) {
    // If the sheet doesn't exist, create it
    sheet = spreadsheet.insertSheet('NBP_Rates_' + year);
  } else {
    // Clear the sheet if it already exists
    sheet.clear();
  }

  // Paste the parsed data into the sheet
  sheet.getRange('A1').setFormula('=IMPORTDATA("https://static.nbp.pl/dane/kursy/Archiwum/archiwum_tab_a_' + year + '.csv",";","PL")');
}

const nbpRatesCacheKey = 'nbpRates'

function getNbpRates() {
  let cache = CacheService.getScriptCache();
  let cachedData = cache.get(nbpRatesCacheKey);

  if (cachedData && cachedData != '[]') {
    const mapArray = JSON.parse(cachedData);
    return new Map(mapArray);
  }

  let nbpRates = refreshNbpRatesCache()
  // Store the result in cache for 360000 seconds
  cache.put(nbpRatesCacheKey, JSON.stringify(Array.from(nbpRates.entries())), 360000);
  return nbpRates;
}

function refreshNbpRatesCache() {
  let cache = CacheService.getScriptCache();

  let nbpRates = new Map();

  let currentYear = new Date().getFullYear();
  for (let i = 0; i < 3; i++){
    populateRatesToMap(nbpRates, currentYear - i);
  }

  // Store the result in cache for 360000 seconds
  var serializedNbpRates = JSON.stringify(Array.from(nbpRates.entries()))
  cache.put(nbpRatesCacheKey, serializedNbpRates, 360000);

  return nbpRates;
}

function populateRatesToMap(nbpRates, year) {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let nbpRatesSheet = spreadsheet.getSheetByName(`NBP_Rates_${year}`);

  let nbpRateRowNumber = 1
  let nbpRateDateValue = convertToDateFormat(nbpRatesSheet.getRange(`A${nbpRateRowNumber + 2}`).getValue()); 

  while (nbpRateDateValue) {
    nbpRates.set(nbpRateDateValue, {
      usd: nbpRatesSheet.getRange('C' + (nbpRateRowNumber + 2)).getValue()
    });

    nbpRateRowNumber++;
    nbpRateDateValue = convertToDateFormat(nbpRatesSheet.getRange('A' + (nbpRateRowNumber + 2)).getValue()); 
  }

  return nbpRates;
}

function convertToDateFormat(input) {
  let inputString = input.toString();

  // Regular expression to validate the input format "yyyyMMdd"
  const regex = /^\d{8}$/;
  // Check if the input matches the expected format
  if (!regex.test(inputString)) {
    return null;
  }
  // Extract year, month, and day from the input string
  const year = parseInt(inputString.substring(0, 4), 10);
  const month = parseInt(inputString.substring(4, 6), 10) - 1; // Month is zero-indexed in JavaScript
  const day = parseInt(inputString.substring(6, 8), 10);
  // Create a new Date object
  const date = new Date(year, month, day);
  // Check if the date is valid

  // Ensure the date object is valid by comparing its components to the input
  if (date.getFullYear() !== year || date.getMonth() !== month || date.getDate() !== day) {
    return null;
  }

  // Format the date as "yyyy-MM-DD"
  const formattedDate = `${year.toString().padStart(4, '0')}-${(month + 1).toString().padStart(2, '0')}-${day.toString().padStart(2, '0')}`;

  return formattedDate;
}

function formatDateString(dateStr) {
  // Extract year, month, and day from the YYYYMMDD format
  var year = dateStr.substring(0, 4);   // Extract the first four characters as the year
  var month = dateStr.substring(4, 6);  // Extract the next two characters as the month
  var day = dateStr.substring(6, 8);    // Extract the last two characters as the day
  // Return the formatted date as YYYY-MM-DD
  return `${year}-${month}-${day}`;
}
