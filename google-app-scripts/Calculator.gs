/**
 * Main entry point for PIT-38 tax calculation.
 * Fetches NBP exchange rates, then calculates FIFO stocks, crypto, and dividends
 * for the year specified in the "Home Page" sheet.
 */
function calculate() {
  setPreviousWorkingDayRate();

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const homePageSheet = spreadsheet.getSheetByName(HOME_PAGE_SHEET_NAME);
  const calculationYear = +homePageSheet.getRange('H16').getValue();

  const reportSheet = spreadsheet.getSheetByName(REPORT_SHEET_NAME) || spreadsheet.insertSheet(REPORT_SHEET_NAME);
  reportSheet.clear();

  const calcLog = [];

  // FIFO Stocks section
  reportSheet.getRange('A3').setValue('Akcje (FIFO)');
  const fifoCountryRows = calculateFifo(spreadsheet, calculationYear, calcLog, reportSheet);

  // Crypto section
  const cryptoHeaderRow = REPORT_FIFO_START_ROW + Math.max(fifoCountryRows, 0) + 1;
  reportSheet.getRange('A' + cryptoHeaderRow).setValue('Kryptowaluty');
  const cryptoStartRow = cryptoHeaderRow + 1;
  const cryptoCountryRows = calculateCrypto(spreadsheet, calculationYear, calcLog, reportSheet, cryptoStartRow);

  // Dividends section
  const divHeaderRow = cryptoStartRow + Math.max(cryptoCountryRows, 0) + 1;
  reportSheet.getRange('A' + divHeaderRow).setValue('Dywidendy');
  const divStartRow = divHeaderRow + 1;
  calculateDividends(spreadsheet, calculationYear, calcLog, reportSheet, divStartRow);

  processCalcLog(spreadsheet, calcLog);

  SpreadsheetApp.getUi().alert('Calculation finished');
}

/**
 * Writes calculation log entries to the console and to a dedicated sheet.
 *
 * @param {Spreadsheet} spreadsheet - The active spreadsheet.
 * @param {string[]} calcLog - Array of log messages to output.
 */
function processCalcLog(spreadsheet, calcLog) {
  calcLog.forEach(logEntry => {
    console.log(logEntry);
  });

  if (calcLog.length > 0) {
    const logSheet = spreadsheet.getSheetByName(CALC_LOG_SHEET_NAME) || spreadsheet.insertSheet(CALC_LOG_SHEET_NAME);
    logSheet.clear();
    logSheet.getRange(1, 1, calcLog.length, 1).setValues(calcLog.map(entry => [entry]));
  }
}

/**
 * Calculates FIFO-based gains/losses for stock transactions, grouped by country.
 * Loads all transactions up to calculationYear to build the buy queue,
 * but only accumulates revenue/cost for sells in the calculation year.
 *
 * @returns {number} Number of country rows written to the report sheet.
 */
function calculateFifo(spreadsheet, calculationYear, calcLog, reportSheet) {
  const sheet = spreadsheet.getSheetByName(FIFO_SHEET_NAME);
  const allData = sheet.getDataRange().getValues();
  const inMemoryFifo = new Map();

  // Phase 1: Data Loading (batch read)
  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    const currentSymbol = row[FIFO_COL.symbol.index];
    if (!currentSymbol) break;

    const transactionDate = new Date(row[FIFO_COL.transactionDate.index]);
    const transactionYear = transactionDate.getFullYear();

    if (transactionYear <= calculationYear) {
      const transaction = {
        date: transactionDate,
        operationType: row[FIFO_COL.operationType.index] === BUY_OPERATION ? 'Buy' : 'Sell',
        count: row[FIFO_COL.count.index],
        price: row[FIFO_COL.price.index],
        currency: row[FIFO_COL.currency.index],
        costs: row[FIFO_COL.costs.index],
        exchangeRate: row[FIFO_COL.exchangeRate.index],
        country: row[FIFO_COL.country.index] || ''
      };

      if (!inMemoryFifo.has(currentSymbol)) {
        inMemoryFifo.set(currentSymbol, []);
      }
      inMemoryFifo.get(currentSymbol).push(transaction);
    }
  }

  // Phase 2: Sorting
  inMemoryFifo.forEach((transactions) => {
    transactions.sort((a, b) => a.date - b.date);
  });

  // Phase 3: FIFO Calculation grouped by country
  const countryTotals = new Map();

  inMemoryFifo.forEach((transactions, symbol) => {
    const buyQueue = [];

    transactions.forEach(transaction => {
      if (transaction.operationType === 'Buy') {
        buyQueue.push({ ...transaction });
      } else if (transaction.operationType === 'Sell') {
        let remainingToSell = transaction.count;
        let totalCost = 0;
        let totalTransactionCost = transaction.costs * transaction.exchangeRate;
        const sellDetails = [];

        while (remainingToSell > 0 && buyQueue.length > 0) {
          const buyTransaction = buyQueue[0];

          if (buyTransaction.count <= remainingToSell) {
            const cost = buyTransaction.count * buyTransaction.price * buyTransaction.exchangeRate;
            totalCost += cost;
            totalTransactionCost += buyTransaction.costs * buyTransaction.exchangeRate;

            remainingToSell -= buyTransaction.count;
            sellDetails.push(`Sold ${buyTransaction.count} shares bought on ${buyTransaction.date.toDateString()} at ${buyTransaction.price} ${buyTransaction.currency} (Cost: ${cost.toFixed(2)} ${CURRENCIES.PLN}, Transaction Cost: ${(buyTransaction.costs * buyTransaction.exchangeRate).toFixed(2)} ${CURRENCIES.PLN})`);
            buyQueue.shift();
          } else {
            const partialCost = remainingToSell * buyTransaction.price * buyTransaction.exchangeRate;
            const partialTransactionCost = (remainingToSell / buyTransaction.count) * buyTransaction.costs * buyTransaction.exchangeRate;

            totalCost += partialCost;
            totalTransactionCost += partialTransactionCost;

            sellDetails.push(`Sold ${remainingToSell} shares bought on ${buyTransaction.date.toDateString()} at ${buyTransaction.price} ${buyTransaction.currency} (Cost: ${partialCost.toFixed(2)} ${CURRENCIES.PLN}, Transaction Cost: ${partialTransactionCost.toFixed(2)} ${CURRENCIES.PLN})`);

            const originalCount = buyTransaction.count;
            buyTransaction.count -= remainingToSell;
            buyTransaction.costs -= (remainingToSell / originalCount) * buyTransaction.costs;
            remainingToSell = 0;
          }
        }

        // Only accumulate for the calculation year
        if (transaction.date.getFullYear() === calculationYear) {
          const totalRevenue = transaction.count * transaction.price * transaction.exchangeRate;
          const gainOrLoss = totalRevenue - totalCost - totalTransactionCost;

          const country = transaction.country || 'Nieznany';
          if (!countryTotals.has(country)) {
            countryTotals.set(country, { revenue: 0, cost: 0, transactionCost: 0 });
          }
          const totals = countryTotals.get(country);
          totals.revenue += totalRevenue;
          totals.cost += totalCost;
          totals.transactionCost += totalTransactionCost;

          calcLog.push(`[${country}] Sold ${transaction.count} shares of ${symbol} on ${transaction.date.toDateString()}:`);
          calcLog.push(...sellDetails);
          calcLog.push(`Total Revenue: ${totalRevenue.toFixed(2)} ${CURRENCIES.PLN}`);
          calcLog.push(`Total Cost: ${totalCost.toFixed(2)} ${CURRENCIES.PLN}`);
          calcLog.push(`Total Transaction Cost: ${totalTransactionCost.toFixed(2)} ${CURRENCIES.PLN}`);
          calcLog.push(`Gain/Loss: ${gainOrLoss.toFixed(2)} ${CURRENCIES.PLN}`);
          calcLog.push('---');
        }
      }
    });
  });

  // Write per-country rows to report
  const countries = Array.from(countryTotals.keys()).sort();
  countries.forEach((country, idx) => {
    const row = REPORT_FIFO_START_ROW + idx;
    const totals = countryTotals.get(country);
    reportSheet.getRange(REPORT_FIFO_COUNTRY_COL + row).setValue(country);
    reportSheet.getRange(REPORT_FIFO_REVENUE_COL + row).setFormula(`=ROUND(${totals.revenue}, 2)`);
    reportSheet.getRange(REPORT_FIFO_COST_COL + row).setFormula(`=ROUND(${totals.cost + totals.transactionCost}, 2)`);
  });

  return countries.length;
}

/**
 * Calculates crypto transaction totals for the given year, grouped by country.
 * Revenue from sells and costs from buys are accumulated separately per country.
 *
 * @returns {number} Number of country rows written to the report sheet.
 */
function calculateCrypto(spreadsheet, calculationYear, calcLog, reportSheet, reportRow) {
  const sheet = spreadsheet.getSheetByName(CRYPTO_SHEET_NAME);
  const allData = sheet.getDataRange().getValues();
  const countryTotals = new Map();

  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    const transactionDate = new Date(row[CRYPTO_COL.transactionDate.index]);
    if (!isValidDate(transactionDate)) break;

    const transactionYear = transactionDate.getFullYear();

    if (transactionYear === calculationYear) {
      const operationType = row[CRYPTO_COL.operationType.index] === BUY_OPERATION ? 'Buy' : 'Sell';
      const amount = row[CRYPTO_COL.transactionSum.index];
      const currency = row[CRYPTO_COL.currency.index];
      const costs = row[CRYPTO_COL.costs.index];
      const exchangeRate = row[CRYPTO_COL.exchangeRate.index];
      const country = row[CRYPTO_COL.country.index] || 'Nieznany';

      const transactionCostPLN = costs * exchangeRate;
      const amountPLN = amount * exchangeRate;

      if (!countryTotals.has(country)) {
        countryTotals.set(country, { revenue: 0, cost: 0, transactionCost: 0 });
      }
      const totals = countryTotals.get(country);

      if (operationType === 'Sell') {
        totals.revenue += amountPLN;
      } else if (operationType === 'Buy') {
        totals.cost += amountPLN;
      }

      totals.transactionCost += transactionCostPLN;

      calcLog.push(`[${country}] Crypto Transaction ${operationType} ${amount} ${currency} on ${transactionDate.toDateString()}:`);
      calcLog.push(`Amount: ${amount}`);
      calcLog.push(`Exchange Rate: ${exchangeRate}`);
      calcLog.push(`Transaction Cost: ${transactionCostPLN.toFixed(2)} ${CURRENCIES.PLN}`);
      calcLog.push(`Total Revenue: ${amountPLN.toFixed(2)} ${CURRENCIES.PLN}`);
      calcLog.push('---');
    }
  }

  const countries = Array.from(countryTotals.keys()).sort();
  countries.forEach((country, idx) => {
    const row = reportRow + idx;
    const totals = countryTotals.get(country);
    reportSheet.getRange('A' + row).setValue(country);
    reportSheet.getRange('B' + row).setFormula(`=ROUND(${totals.revenue}, 2)`);
    reportSheet.getRange('C' + row).setFormula(`=ROUND(${totals.cost + totals.transactionCost}, 2)`);
  });

  return countries.length;
}

/**
 * Calculates dividend totals for the given year, grouped by country.
 *
 * @returns {number} Number of country rows written to the report sheet.
 */
function calculateDividends(spreadsheet, calculationYear, calcLog, reportSheet, reportRow) {
  const sheet = spreadsheet.getSheetByName(DIVIDENDS_SHEET_NAME);
  const allData = sheet.getDataRange().getValues();
  const countryTotals = new Map();

  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    const transactionDate = new Date(row[DIVIDENDS_COL.transactionDate.index]);
    if (!isValidDate(transactionDate)) break;

    const transactionYear = transactionDate.getFullYear();

    if (transactionYear === calculationYear) {
      const amount = row[DIVIDENDS_COL.transactionSum.index];
      const currency = row[DIVIDENDS_COL.currency.index];
      const costs = row[DIVIDENDS_COL.costs.index];
      const exchangeRate = row[DIVIDENDS_COL.exchangeRate.index];
      const country = row[DIVIDENDS_COL.country.index] || 'Nieznany';

      const transactionCostPLN = costs * exchangeRate;
      const amountPLN = amount * exchangeRate;

      if (!countryTotals.has(country)) {
        countryTotals.set(country, { revenue: 0, transactionCost: 0 });
      }
      const totals = countryTotals.get(country);
      totals.revenue += amountPLN;
      totals.transactionCost += transactionCostPLN;

      calcLog.push(`[${country}] Dividends Transaction ${amount} ${currency} on ${transactionDate.toDateString()}:`);
      calcLog.push(`Amount: ${amount}`);
      calcLog.push(`Exchange Rate: ${exchangeRate}`);
      calcLog.push(`Transaction Cost: ${transactionCostPLN.toFixed(2)} ${CURRENCIES.PLN}`);
      calcLog.push(`Total Revenue: ${amountPLN.toFixed(2)} ${CURRENCIES.PLN}`);
      calcLog.push('---');
    }
  }

  const countries = Array.from(countryTotals.keys()).sort();
  countries.forEach((country, idx) => {
    const row = reportRow + idx;
    const totals = countryTotals.get(country);
    reportSheet.getRange('A' + row).setValue(country);
    reportSheet.getRange('B' + row).setFormula(`=ROUND(${totals.revenue}, 2)`);
    reportSheet.getRange('C' + row).setFormula(`=ROUND(${totals.transactionCost}, 2)`);
  });

  return countries.length;
}

/**
 * Populates NBP exchange rates for all transaction sheets (FIFO, Crypto, Dividends).
 * For each transaction, finds the previous working day's rate per Polish tax law.
 */
function setPreviousWorkingDayRate() {
  const nbpRates = getNbpRates();

  setPreviousWorkingDayWithParams(FIFO_SHEET_NAME, nbpRates, FIFO_COL);
  setPreviousWorkingDayWithParams(CRYPTO_SHEET_NAME, nbpRates, CRYPTO_COL);
  setPreviousWorkingDayWithParams(DIVIDENDS_SHEET_NAME, nbpRates, DIVIDENDS_COL);
}

/**
 * Fills NBP exchange rate columns for each row in the given sheet.
 * Looks up the previous working day rate from the NBP rates map and writes
 * the rate date, rate value, and PLN-converted sum/cost formulas.
 *
 * @param {string} sheetName - Name of the sheet to process.
 * @param {Map<string, Object>} nbpRates - Map of date strings to currency rate objects.
 * @param {Object} col - Column map (e.g. FIFO_COL) with properties containing {index, letter}.
 */
function setPreviousWorkingDayWithParams(sheetName, nbpRates, col) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  const allData = sheet.getDataRange().getValues();

  // Collect updates to write in batch
  const updates = [];

  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    const dateValue = row[col.transactionDate.index];
    if (!isValidDate(dateValue)) break;

    const nbpRateDate = new Date(dateValue);
    const existingNbpDate = row[col.nbpRateDate.index];

    // Use previous working day rate (per Polish tax law)
    if (!existingNbpDate) {
      nbpRateDate.setDate(nbpRateDate.getDate() - 1);
    }

    let maxDepth = 9;
    while (!nbpRates.has(formatDate(nbpRateDate)) && maxDepth > 0) {
      nbpRateDate.setDate(nbpRateDate.getDate() - 1);
      maxDepth--;
    }

    if (maxDepth > 0) {
      const formattedDate = formatDate(nbpRateDate);
      const currency = row[col.currency.index];
      const rate = currency === CURRENCIES.PLN ? 1 : nbpRates.get(formattedDate)[currency.toLowerCase()];
      const rowNum = i + 1;

      updates.push({
        rowNum: rowNum,
        nbpRateDate: nbpRateDate,
        rate: rate,
        sumFormula: `=${col.transactionSum.letter}${rowNum}*${col.exchangeRate.letter}${rowNum}`,
        costsFormula: `=${col.costs.letter}${rowNum}*${col.exchangeRate.letter}${rowNum}`
      });
    } else {
      Logger.log(`Issue with processing record: ${i}`);
    }
  }

  // Batch write all updates
  updates.forEach(function(update) {
    sheet.getRange(`${col.nbpRateDate.letter}${update.rowNum}`).setValue(update.nbpRateDate);
    sheet.getRange(`${col.exchangeRate.letter}${update.rowNum}`).setValue(update.rate);
    sheet.getRange(`${col.sumPLN.letter}${update.rowNum}`).setFormula(update.sumFormula);
    sheet.getRange(`${col.costsPLN.letter}${update.rowNum}`).setFormula(update.costsFormula);
  });
}

/**
 * Checks if a value is a valid date.
 * @param {any} dateValue
 * @returns {boolean}
 */
function isValidDate(dateValue) {
  return Object.prototype.toString.call(dateValue) === "[object Date]" && !isNaN(dateValue);
}

/**
 * Formats a date to 'yyyy-MM-dd' based on the script timezone.
 * @param {Date} date
 * @returns {string}
 */
function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}
