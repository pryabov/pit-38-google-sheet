/**
 * Main entry point for PIT-38 tax calculation.
 * Fetches NBP exchange rates, then calculates FIFO stocks, crypto, and dividends
 * for the year specified in the Settings sheet.
 */
function calculate() {
  setPreviousWorkingDayRate();

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const settingsSheet = spreadsheet.getSheetByName(SETTINGS_SHEET_NAME);
  const calculationYear = +settingsSheet.getRange('B2').getValue();

  const reportSheet = spreadsheet.getSheetByName(REPORT_SHEET_NAME) || spreadsheet.insertSheet(REPORT_SHEET_NAME);

  const calcLog = [];

  calculateFifo(spreadsheet, calculationYear, calcLog, reportSheet);
  calculateCrypto(spreadsheet, calculationYear, calcLog, reportSheet);
  calculateDividends(spreadsheet, calculationYear, calcLog, reportSheet);

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
 * Calculates FIFO-based gains/losses for stock transactions.
 * Loads all transactions up to calculationYear to build the buy queue,
 * but only accumulates revenue/cost for sells in the calculation year.
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
        exchangeRate: row[FIFO_COL.exchangeRate.index]
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

  // Phase 3: FIFO Calculation
  let totalRevenueAccumulated = 0;
  let totalCostAccumulated = 0;
  let totalTransactionsCostAccumulated = 0;

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
            sellDetails.push(`Sold ${buyTransaction.count} shares bought on ${buyTransaction.date.toDateString()} at ${buyTransaction.price} ${buyTransaction.currency} (Cost: ${cost.toFixed(2)} PLN, Transaction Cost: ${(buyTransaction.costs * buyTransaction.exchangeRate).toFixed(2)} PLN)`);
            buyQueue.shift();
          } else {
            const partialCost = remainingToSell * buyTransaction.price * buyTransaction.exchangeRate;
            const partialTransactionCost = (remainingToSell / buyTransaction.count) * buyTransaction.costs * buyTransaction.exchangeRate;

            totalCost += partialCost;
            totalTransactionCost += partialTransactionCost;

            sellDetails.push(`Sold ${remainingToSell} shares bought on ${buyTransaction.date.toDateString()} at ${buyTransaction.price} ${buyTransaction.currency} (Cost: ${partialCost.toFixed(2)} PLN, Transaction Cost: ${partialTransactionCost.toFixed(2)} PLN)`);

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

          totalRevenueAccumulated += totalRevenue;
          totalCostAccumulated += totalCost;
          totalTransactionsCostAccumulated += totalTransactionCost;

          calcLog.push(`Sold ${transaction.count} shares of ${symbol} on ${transaction.date.toDateString()}:`);
          calcLog.push(...sellDetails);
          calcLog.push(`Total Revenue: ${totalRevenue.toFixed(2)} PLN`);
          calcLog.push(`Total Cost: ${totalCost.toFixed(2)} PLN`);
          calcLog.push(`Total Transaction Cost: ${totalTransactionCost.toFixed(2)} PLN`);
          calcLog.push(`Gain/Loss: ${gainOrLoss.toFixed(2)} PLN`);
          calcLog.push('---');
        }
      }
    });
  });

  reportSheet.getRange(REPORT_CELLS.fifoRevenue).setFormula(`=ROUND(${totalRevenueAccumulated}, 2)`);
  reportSheet.getRange(REPORT_CELLS.fifoCost).setFormula(`=ROUND(${totalCostAccumulated}+${totalTransactionsCostAccumulated}, 2)`);
}

/**
 * Calculates crypto transaction totals for the given year.
 * Revenue from sells and costs from buys are accumulated separately.
 */
function calculateCrypto(spreadsheet, calculationYear, calcLog, reportSheet) {
  const sheet = spreadsheet.getSheetByName(CRYPTO_SHEET_NAME);
  const allData = sheet.getDataRange().getValues();
  let totalRevenueAccumulated = 0;
  let totalCostAccumulated = 0;
  let totalTransactionsCostAccumulated = 0;

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

      const transactionCostPLN = costs * exchangeRate;
      const amountPLN = amount * exchangeRate;

      if (operationType === 'Sell') {
        totalRevenueAccumulated += amountPLN;
      } else if (operationType === 'Buy') {
        totalCostAccumulated += amountPLN;
      }

      totalTransactionsCostAccumulated += transactionCostPLN;

      calcLog.push(`Crypto Transaction ${operationType} ${amount} ${currency} on ${transactionDate.toDateString()}:`);
      calcLog.push(`Amount: ${amount}`);
      calcLog.push(`Exchange Rate: ${exchangeRate}`);
      calcLog.push(`Transaction Cost: ${transactionCostPLN.toFixed(2)} PLN`);
      calcLog.push(`Total Revenue: ${amountPLN.toFixed(2)} PLN`);
      calcLog.push('---');
    }
  }

  reportSheet.getRange(REPORT_CELLS.cryptoRevenue).setFormula(`=ROUND(${totalRevenueAccumulated}, 2)`);
  reportSheet.getRange(REPORT_CELLS.cryptoCost).setFormula(`=ROUND(${totalCostAccumulated}+${totalTransactionsCostAccumulated}, 2)`);
}

/**
 * Calculates dividend totals for the given year.
 */
function calculateDividends(spreadsheet, calculationYear, calcLog, reportSheet) {
  const sheet = spreadsheet.getSheetByName(DIVIDENDS_SHEET_NAME);
  const allData = sheet.getDataRange().getValues();
  let totalRevenueAccumulated = 0;
  let totalTransactionsCostAccumulated = 0;

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

      const transactionCostPLN = costs * exchangeRate;
      const amountPLN = amount * exchangeRate;

      totalRevenueAccumulated += amountPLN;
      totalTransactionsCostAccumulated += transactionCostPLN;

      calcLog.push(`Dividends Transaction ${amount} ${currency} on ${transactionDate.toDateString()}:`);
      calcLog.push(`Amount: ${amount}`);
      calcLog.push(`Exchange Rate: ${exchangeRate}`);
      calcLog.push(`Transaction Cost: ${transactionCostPLN.toFixed(2)} PLN`);
      calcLog.push(`Total Revenue: ${amountPLN.toFixed(2)} PLN`);
      calcLog.push('---');
    }
  }

  reportSheet.getRange(REPORT_CELLS.dividendsRevenue).setFormula(`=ROUND(${totalRevenueAccumulated}, 2)`);
  reportSheet.getRange(REPORT_CELLS.dividendsCost).setFormula(`=ROUND(${totalTransactionsCostAccumulated}, 2)`);
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
      const rate = currency === 'PLN' ? 1 : nbpRates.get(formattedDate)[currency.toLowerCase()];
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
