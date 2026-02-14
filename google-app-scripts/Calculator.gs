function calculate() {
  // SpreadsheetApp.getUi().alert('Calculation started. Pls, wait for the finish notification!');

  setPreviousWorkingDayRate();

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  var settingsSheet = spreadsheet.getSheetByName('Settings');
  var calculationYear = +settingsSheet.getRange('B2').getValue();

  var calcLog = [];

  calculateFifo(spreadsheet, calculationYear, calcLog);
  calculateCrypto(spreadsheet, calculationYear, calcLog);
  calculateDividends(spreadsheet, calculationYear, calcLog);

  processCalcLog(spreadsheet, calcLog);

  SpreadsheetApp.getUi().alert('Calculation finished')
}

function processCalcLog(spreadsheet, calcLog) {
  calcLog.forEach(logEntry => {
    console.log(logEntry);
  });

  if (calcLog.length > 0) {
    var logSheet = spreadsheet.getSheetByName('Calculation Log') || spreadsheet.insertSheet('Calculation Log');
    logSheet.clear();
    logSheet.getRange(1, 1, calcLog.length, 1).setValues(calcLog.map(entry => [entry]));
  }
}

function calculateFifo(spreadsheet, calculationYear, calcLog) {
  var sheet = spreadsheet.getSheetByName('FIFO Stocks Transactions');

  var inMemoryFifo = new Map();

  var totalRevenueAccumulated = 0;
  var totalCostAccumulated = 0;
  var totalTransactionsCostAccumulated = 0;

  var i = 2;
  var currentSymbol = sheet.getRange(`B${i}`).getValue();

  // Phase 1: Data Loading
  while (currentSymbol) {
    var transactionDate = new Date(sheet.getRange(`F${i}`).getValue());
    // Extract the year from the transactionDate
    var transactionYear = transactionDate.getFullYear();

    if (transactionYear <= calculationYear) {
        var transaction = {
          date: transactionDate,
          operationType: sheet.getRange(`E${i}`).getValue() === 'Kupowanie' ? 'Buy' : 'Sell',
          count: sheet.getRange(`G${i}`).getValue(),
          price: sheet.getRange(`H${i}`).getValue(),
          currency: sheet.getRange(`J${i}`).getValue(),
          costs: sheet.getRange(`K${i}`).getValue(),
          exchangeRate: sheet.getRange(`N${i}`).getValue()
        };

        if (!inMemoryFifo.has(currentSymbol)) {
          inMemoryFifo.set(currentSymbol, []);
        }

        inMemoryFifo.get(currentSymbol).push(transaction);
    }

    i++;
    currentSymbol = sheet.getRange(`B${i}`).getValue();
  }

  // Phase 2: Sorting
  inMemoryFifo.forEach((transactions, symbol) => {
    transactions.sort((a, b) => a.date - b.date);
  });

  // Phase 3: Calculation
  inMemoryFifo.forEach((transactions, symbol) => {
    var buyQueue = [];

    transactions.forEach(transaction => {
      if (transaction.operationType === 'Buy') {
        buyQueue.push({ ...transaction });
      } else if (transaction.operationType === 'Sell') {
        var remainingToSell = transaction.count;
        var totalCost = 0;
        var totalTransactionCost = transaction.costs * transaction.exchangeRate;
        var sellDetails = [];

        while (remainingToSell > 0 && buyQueue.length > 0) {
          var buyTransaction = buyQueue[0];

          if (buyTransaction.count <= remainingToSell) {
            var cost = buyTransaction.count * buyTransaction.price * buyTransaction.exchangeRate;
            totalCost += cost;
            totalTransactionCost += buyTransaction.costs * buyTransaction.exchangeRate;

            remainingToSell -= buyTransaction.count;
            sellDetails.push(`Sold ${buyTransaction.count} shares bought on ${buyTransaction.date.toDateString()} at ${buyTransaction.price} ${buyTransaction.currency} (Cost: ${(cost).toFixed(2)} PLN, Transaction Cost: ${(buyTransaction.costs * buyTransaction.exchangeRate).toFixed(2)} PLN)`);
            buyQueue.shift();
          } else {
            var partialCost = remainingToSell * buyTransaction.price * buyTransaction.exchangeRate;
            var transactionCost = (remainingToSell / buyTransaction.count) * buyTransaction.costs * buyTransaction.exchangeRate;

            totalCost += partialCost;
            totalTransactionCost += transactionCost;

            sellDetails.push(`Sold ${remainingToSell} shares bought on ${buyTransaction.date.toDateString()} at ${buyTransaction.price} ${buyTransaction.currency} (Cost: ${(partialCost).toFixed(2)} PLN, Transaction Cost: ${(transactionCost).toFixed(2)} PLN)`);
            buyTransaction.count -= remainingToSell;
            buyTransaction.costs -= (remainingToSell / buyTransaction.count) * buyTransaction.costs;
            remainingToSell = 0;
          }
        }

        var totalRevenue = transaction.count * transaction.price * transaction.exchangeRate;
        var gainOrLoss = totalRevenue - totalCost - totalTransactionCost;

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
    });
  });

  var reportSheet = spreadsheet.getSheetByName('Report') || spreadsheet.insertSheet('Report');
  reportSheet.getRange('A4').setFormula(`=ROUND(${totalRevenueAccumulated}, 2)`);
  reportSheet.getRange('B4').setFormula(`=ROUND(${totalCostAccumulated}+${totalTransactionsCostAccumulated}, 2)`);
}

function calculateCrypto(spreadsheet, calculationYear, calcLog) {
  var sheet = spreadsheet.getSheetByName('Crypto Currencies');

  var inMemoryCrypto = [];

  var totalRevenueAccumulated = 0;
  var totalCostAccumulated = 0;
  var totalTransactionsCostAccumulated = 0;

  var i = 2;
  var transactionDate = new Date(sheet.getRange(`F${i}`).getValue());

  // Phase 1: Data Loading
  while (isValidDate(transactionDate)) {
    var transactionYear = transactionDate.getFullYear();

    if (transactionYear === calculationYear) {
      var transaction = {
        rowNumber: i,
        date: transactionDate,
        operationType: sheet.getRange(`E${i}`).getValue() === 'Kupowanie' ? 'Buy' : 'Sell',
        amount: sheet.getRange(`I${i}`).getValue(),
        currency: sheet.getRange(`J${i}`).getValue(),
        costs: sheet.getRange(`K${i}`).getValue(),
        exchangeRate: sheet.getRange(`N${i}`).getValue()
      };

      inMemoryCrypto.push(transaction);

      var transactionCostPLN = transaction.costs * transaction.exchangeRate;
      var amountPLN = transaction.amount * transaction.exchangeRate;

      if (transaction.operationType === 'Sell') {
        totalRevenueAccumulated += amountPLN;
      } else if (transaction.operationType === 'Buy') {
        totalCostAccumulated += amountPLN;
      }

      totalTransactionsCostAccumulated += transactionCostPLN;

      // Log the transaction details
      calcLog.push(`Crypto Transaction ${transaction.operationType} ${transaction.amount} ${transaction.currency} on ${transaction.date.toDateString()}:`);
      calcLog.push(`Amount: ${transaction.amount}`);
      calcLog.push(`Exchange Rate: ${transaction.exchangeRate}`);
      calcLog.push(`Transaction Cost: ${transactionCostPLN.toFixed(2)} PLN`);
      calcLog.push(`Total Revenue: ${amountPLN.toFixed(2)} PLN`);
      calcLog.push('---');
    }

    i++;
    transactionDate = new Date(sheet.getRange(`F${i}`).getValue());
  }

  // Optionally, write the results to a sheet named 'Crypto Report'
  var reportSheet = spreadsheet.getSheetByName('Report') || spreadsheet.insertSheet('Report');
  reportSheet.getRange('D4').setFormula(`=ROUND(${totalRevenueAccumulated}, 2)`);
  reportSheet.getRange('E4').setFormula(`=ROUND(${totalCostAccumulated}+${totalTransactionsCostAccumulated}, 2)`);
}

function calculateDividends(spreadsheet, calculationYear, calcLog) {
  var sheet = spreadsheet.getSheetByName('Dividends');

  var inMemoryDividends = [];

  var totalRevenueAccumulated = 0;
  var totalTransactionsCostAccumulated = 0;

  var i = 2;
  var transactionDate = new Date(sheet.getRange(`E${i}`).getValue());

  // Phase 1: Data Loading
  while (isValidDate(transactionDate)) {
    var transactionYear = transactionDate.getFullYear();

    if (transactionYear === calculationYear) {
      var transaction = {
        rowNumber: i,
        date: transactionDate,
        amount: sheet.getRange(`F${i}`).getValue(),
        currency: sheet.getRange(`G${i}`).getValue(),
        costs: sheet.getRange(`H${i}`).getValue(),
        exchangeRate: sheet.getRange(`K${i}`).getValue()
      };

      inMemoryDividends.push(transaction);

      var transactionCostPLN = transaction.costs * transaction.exchangeRate;
      var amountPLN = transaction.amount * transaction.exchangeRate;

      totalRevenueAccumulated += amountPLN;
      totalTransactionsCostAccumulated += transactionCostPLN;

      // Log the transaction details
      calcLog.push(`Dividends Transaction ${transaction.amount} ${transaction.currency} on ${transaction.date.toDateString()}:`);
      calcLog.push(`Amount: ${transaction.amount}`);
      calcLog.push(`Exchange Rate: ${transaction.exchangeRate}`);
      calcLog.push(`Transaction Cost: ${transactionCostPLN.toFixed(2)} PLN`);
      calcLog.push(`Total Revenue: ${amountPLN.toFixed(2)} PLN`);
      calcLog.push('---');
    }

    i++;
    transactionDate = new Date(sheet.getRange(`E${i}`).getValue());
  }

  // Optionally, write the results to a sheet named 'Crypto Report'
  var reportSheet = spreadsheet.getSheetByName('Report') || spreadsheet.insertSheet('Report');
  reportSheet.getRange('G4').setFormula(`=ROUND(${totalRevenueAccumulated}, 2)`);
  reportSheet.getRange('H4').setFormula(`=ROUND(${totalTransactionsCostAccumulated}, 2)`);
}

function setPreviousWorkingDayRate() {
  const nbpRates = getNbpRates();

  setPreviousWorkingDayWithParams('FIFO Stocks Transactions', nbpRates, 'F', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P');
  setPreviousWorkingDayWithParams('Crypto Currencies', nbpRates, 'F', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P');
  setPreviousWorkingDayWithParams('Dividends', nbpRates, 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M');
}

function setPreviousWorkingDayWithParams(
  sheetName,
  nbpRates,
  transactionDateColumn,
  transactionSumColumn,
  transactionCurrencyColumn,
  transactionCommissionColumn,
  transactionCommissionCurrencyColumn,
  nbpRateDateColumn,
  nbpRateValueColumn,
  sumColumn,
  sumCostsColumn
) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);

  let rowNumber = 1;
  let dateValue = sheet.getRange(`${transactionDateColumn}${rowNumber + 1}`).getValue();

  while (isValidDate(dateValue)) {
    let nbpRateDate = new Date(dateValue);

    if (sheet.getRange(`${nbpRateDateColumn}${rowNumber + 1}`))
    nbpRateDate.setDate(nbpRateDate.getDate() - 1);

    let maxDepth = 9;
    while (!nbpRates.has(formatDate(nbpRateDate)) && maxDepth > 0) {
      nbpRateDate.setDate(nbpRateDate.getDate() - 1);
      maxDepth--;
    }

    if (maxDepth > 0) {
      let formattedDate = formatDate(nbpRateDate);
      let nbpRateUsd = sheet.getRange(`${transactionCurrencyColumn}${rowNumber + 1}`).getValue() === 'PLN' ? 1 : nbpRates.get(formattedDate).usd;

      sheet.getRange(`${nbpRateDateColumn}${rowNumber + 1}`).setValue(nbpRateDate);
      sheet.getRange(`${nbpRateValueColumn}${rowNumber + 1}`).setValue(nbpRateUsd);

      sheet.getRange(`${sumColumn}${rowNumber + 1}`).setFormula(
        `=${transactionSumColumn}${rowNumber + 1}*${nbpRateValueColumn}${rowNumber + 1}`
      );
      sheet.getRange(`${sumCostsColumn}${rowNumber + 1}`).setFormula(
        `=${transactionCommissionColumn}${rowNumber + 1}*${nbpRateValueColumn}${rowNumber + 1}`
      );
    } else {
      Logger.log(`Issue with processing record: ${rowNumber}`);
    }

    rowNumber++;
    dateValue = sheet.getRange(`${transactionDateColumn}${rowNumber + 1}`).getValue();
  }
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
