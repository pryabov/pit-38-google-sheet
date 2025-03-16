function calculate() {
  // SpreadsheetApp.getUi().alert('Calculation started. Pls, wait for the finish notification!');

  setPreviousWorkingDayRate();

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  var settingsSheet = spreadsheet.getSheetByName('Settings');
  var calculationYear = +settingsSheet.getRange('B2').getValue();

  var calcLog = [];

  calculateFifo(spreadsheet, calculationYear, calcLog);
  // calculateCrypto(spreadsheet, calculationYear, calcLog);
  // calculateDividends(spreadsheet, calculationYear, calcLog);

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

  // Map to store transactions grouped by country and broker
  var countryGroups = new Map();

  var i = 2; // Start from row 2, assuming headers in row 1
  var currentSymbol = sheet.getRange(`B${i}`).getValue();

  // Phase 1: Data loading and grouping by country and broker
  while (currentSymbol) {
    var country = sheet.getRange(`D${i}`).getValue(); // Get country from column D
    var broker = sheet.getRange(`C${i}`).getValue(); // Get broker from column C
    var transactionDate = new Date(sheet.getRange(`F${i}`).getValue()); // Transaction date
    var transactionYear = transactionDate.getFullYear(); // Extract year from date

    if (transactionYear <= calculationYear) {
      var transaction = {
        symbol: currentSymbol,
        date: transactionDate,
        operationType: sheet.getRange(`E${i}`).getValue() === 'Kupowanie' ? 'Buy' : 'Sell',
        count: sheet.getRange(`G${i}`).getValue(),
        price: sheet.getRange(`I${i}`).getValue(),
        costs: sheet.getRange(`K${i}`).getValue(),
        exchangeRate: sheet.getRange(`N${i}`).getValue()
      };

      if (!countryGroups.has(country)) {
        countryGroups.set(country, new Map());
      }

      var brokerGroup = countryGroups.get(country);
      if (!brokerGroup.has(broker)) {
        brokerGroup.set(broker, new Map());
      }

      var symbolGroup = brokerGroup.get(broker);
      if (!symbolGroup.has(currentSymbol)) {
        symbolGroup.set(currentSymbol, []);
      }

      symbolGroup.get(currentSymbol).push(transaction);
    }

    i++;
    currentSymbol = sheet.getRange(`B${i}`).getValue();
  }

  var reportRow = 4; // Initial row for report data

  // Phase 2 & 3: Sorting and FIFO calculation by country, broker, and symbol
  countryGroups.forEach((brokersMap, country) => {
    calcLog.push(`Country: ${country}`);

    var totalRevenueAccumulated = 0;
    var totalCostAccumulated = 0;

    brokersMap.forEach((symbolsMap, broker) => {
      calcLog.push(`Broker: ${broker}`);

      symbolsMap.forEach((symbolTransactions, symbol) => {
        symbolTransactions.sort((a, b) => a.date - b.date); // Sort transactions by date

        var buyQueue = [];

        symbolTransactions.forEach(transaction => {
          if (transaction.operationType === 'Buy') {
            buyQueue.push({ ...transaction }); // Add buy transactions to queue
          } else if (transaction.operationType === 'Sell') {
            var remainingToSell = transaction.count;
            var totalCost = 0;
            var totalTransactionCost = transaction.costs * transaction.exchangeRate;
            var sellDetails = [];

            // FIFO logic for selling
            while (remainingToSell > 0 && buyQueue.length > 0) {
              var buyTransaction = buyQueue[0];

              if (buyTransaction.count <= remainingToSell) {
                var cost = buyTransaction.count * buyTransaction.price * buyTransaction.exchangeRate;
                totalCost += cost;
                totalTransactionCost += buyTransaction.costs * buyTransaction.exchangeRate;
                remainingToSell -= buyTransaction.count;
                sellDetails.push(`Sold ${buyTransaction.count} shares bought on ${buyTransaction.date.toDateString()}`);
                buyQueue.shift();
              } else {
                var partialCost = remainingToSell * buyTransaction.price * buyTransaction.exchangeRate;
                var transactionCost = (remainingToSell / buyTransaction.count) * buyTransaction.costs * buyTransaction.exchangeRate;
                totalCost += partialCost;
                totalTransactionCost += transactionCost;
                sellDetails.push(`Sold ${remainingToSell} shares bought on ${buyTransaction.date.toDateString()}`);
                buyTransaction.count -= remainingToSell;
                buyTransaction.costs -= transactionCost;
                remainingToSell = 0;
              }
            }

            var totalRevenue = transaction.count * transaction.price * transaction.exchangeRate;
            totalRevenueAccumulated += totalRevenue;
            totalCostAccumulated += totalCost;

            // Log transaction details
            calcLog.push(...sellDetails);
            calcLog.push(`Total Revenue: ${totalRevenue.toFixed(2)} PLN`);
            calcLog.push(`Total Cost: ${totalCost.toFixed(2)} PLN`);
            calcLog.push(`Total Transaction Cost: ${totalTransactionCost.toFixed(2)} PLN`);
          }
        });
      });
    });

    // Write totals per country
    var reportSheet = spreadsheet.getSheetByName('Report') || spreadsheet.insertSheet('Report');
    reportSheet.getRange(`A${reportRow}`).setValue(country);
    reportSheet.getRange(`B${reportRow}`).setValue(totalRevenueAccumulated);
    reportSheet.getRange(`C${reportRow}`).setValue(totalCostAccumulated);

    reportRow++;

    calcLog.push(`Country ${country} totals recorded.`);
    calcLog.push('===');
  });

  // Add totals row with sum formula
  reportSheet.getRange(`A${reportRow}`).setValue('Total');
  reportSheet.getRange(`B${reportRow}`).setFormula(`=SUM(B4:B${reportRow - 1})`);
  reportSheet.getRange(`C${reportRow}`).setFormula(`=ROUND(SUM(C4:C${reportRow}), 2)`);
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
