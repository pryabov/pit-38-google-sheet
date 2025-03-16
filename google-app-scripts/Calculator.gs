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
  var reportSheet = spreadsheet.getSheetByName('Report') || spreadsheet.insertSheet('Report');

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
        price: sheet.getRange(`H${i}`).getValue(),
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
      calcLog.push('---------');
      calcLog.push(`Broker: ${broker}`);
      calcLog.push('---------');

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
            totalCostAccumulated += totalCost + totalTransactionCost;

            // Log transaction details
            calcLog.push(...sellDetails);
            calcLog.push(`Total Revenue: ${totalRevenue.toFixed(2)} PLN`);
            calcLog.push(`Total Cost: ${totalCost.toFixed(2)} PLN`);
            calcLog.push(`Total Transaction Cost: ${totalTransactionCost.toFixed(2)} PLN`);
            calcLog.push('---');
          }
        });
      });
    });

    // Write totals per country
    reportSheet.getRange(`A${reportRow}`).setValue(country);
    reportSheet.getRange(`B${reportRow}`).setValue(totalRevenueAccumulated);
    reportSheet.getRange(`C${reportRow}`).setValue(totalCostAccumulated);

    reportRow++;

    calcLog.push(`Country ${country} totals recorded.`);
    calcLog.push('---------');
  });

  // Add totals row with sum formula
  reportSheet.getRange(`A${reportRow}`).setValue('Total');
  reportSheet.getRange(`B${reportRow}`).setFormula(`=ROUND(SUM(B4:B${reportRow - 1}), 2)`);
  reportSheet.getRange(`C${reportRow}`).setFormula(`=ROUND(SUM(C4:C${reportRow - 1}), 2)`);
}

function calculateCrypto(spreadsheet, calculationYear, calcLog) {
  var sheet = spreadsheet.getSheetByName('Crypto Currencies');
  var reportSheet = spreadsheet.getSheetByName('Report') || spreadsheet.insertSheet('Report');

  var countryGroups = new Map();

  var i = 2;
  var transactionDate = new Date(sheet.getRange(`F${i}`).getValue());

  // Phase 1: Data loading and grouping by country
  while (isValidDate(transactionDate)) {
    var transactionYear = transactionDate.getFullYear();

    if (transactionYear === calculationYear) {
      var country = sheet.getRange(`D${i}`).getValue();
      var transaction = {
        date: transactionDate,
        operationType: sheet.getRange(`E${i}`).getValue() === 'Kupowanie' ? 'Buy' : 'Sell',
        amount: sheet.getRange(`I${i}`).getValue(),
        currency: sheet.getRange(`J${i}`).getValue(),
        costs: sheet.getRange(`K${i}`).getValue(),
        exchangeRate: sheet.getRange(`N${i}`).getValue()
      };

      if (!countryGroups.has(country)) {
        countryGroups.set(country, []);
      }

      countryGroups.get(country).push(transaction);
    }

    i++;
    transactionDate = new Date(sheet.getRange(`F${i}`).getValue());
  }

  var reportRow = 4;

  // Phase 2: Calculations by country
  countryGroups.forEach((transactions, country) => {
    var totalRevenue = 0;
    var totalCost = 0;
    var totalTransactionCosts = 0;

    transactions.forEach(transaction => {
      var transactionCostPLN = transaction.costs * transaction.exchangeRate;
      var amountPLN = transaction.amount * transaction.exchangeRate;

      if (transaction.operationType === 'Sell') {
        totalRevenue += amountPLN;
      } else if (transaction.operationType === 'Buy') {
        totalCost += amountPLN;
      }

      totalTransactionCosts += transactionCostPLN;

      calcLog.push(`Crypto Transaction ${transaction.operationType} ${transaction.amount} ${transaction.currency} on ${transaction.date.toDateString()} (${country}):`);
      calcLog.push(`Amount: ${amountPLN.toFixed(2)} PLN, Transaction Cost: ${transactionCostPLN.toFixed(2)} PLN`);
      calcLog.push('---');
    });

    // Write results per country into columns E-G
    reportSheet.getRange(`E${reportRow}`).setValue(country);
    reportSheet.getRange(`F${reportRow}`).setValue(totalRevenue.toFixed(2));
    reportSheet.getRange(`G${reportRow}`).setValue((totalCost + totalTransactionCosts).toFixed(2));

    calcLog.push(`Country: ${country}, Revenue: ${totalRevenue.toFixed(2)} PLN, Total Costs: ${(totalCost + totalTransactionCosts).toFixed(2)} PLN`);
    calcLog.push('----');

    reportRow++;
  });

  // Totals row
  reportSheet.getRange(`E${reportRow}`).setValue('Total');
  reportSheet.getRange(`F${reportRow}`).setFormula(`=ROUND(SUM(F4:F${reportRow - 1}), 2)`);
  reportSheet.getRange(`G${reportRow}`).setFormula(`=ROUND(SUM(G4:G${reportRow - 1}), 2)`);
}

function calculateDividends(spreadsheet, calculationYear, calcLog) {
  var sheet = spreadsheet.getSheetByName('Dividends');
  var reportSheet = spreadsheet.getSheetByName('Report') || spreadsheet.insertSheet('Report');

  var countryGroups = new Map();

  var i = 2;
  var transactionDate = new Date(sheet.getRange(`E${i}`).getValue());

  // Phase 1: Data loading and grouping by country
  while (isValidDate(transactionDate)) {
    var transactionYear = transactionDate.getFullYear();

    if (transactionYear === calculationYear) {
      var country = sheet.getRange(`D${i}`).getValue();
      var transaction = {
        date: transactionDate,
        amount: sheet.getRange(`F${i}`).getValue(),
        currency: sheet.getRange(`G${i}`).getValue(),
        costs: sheet.getRange(`H${i}`).getValue(),
        exchangeRate: sheet.getRange(`K${i}`).getValue()
      };

      if (!countryGroups.has(country)) {
        countryGroups.set(country, []);
      }

      countryGroups.get(country).push(transaction);
    }

    i++;
    transactionDate = new Date(sheet.getRange(`E${i}`).getValue());
  }

  var reportRow = 4;

  // Phase 2: Calculations by country
  countryGroups.forEach((transactions, country) => {
    var totalRevenue = 0;
    var totalTransactionCosts = 0;

    transactions.forEach(transaction => {
      var transactionCostPLN = transaction.costs * transaction.exchangeRate;
      var amountPLN = transaction.amount * transaction.exchangeRate;

      totalRevenue += amountPLN;
      totalTransactionCosts += transactionCostPLN;

      calcLog.push(`Dividend Transaction ${transaction.amount} ${transaction.currency} on ${transaction.date.toDateString()} (${country}):`);
      calcLog.push(`Amount: ${amountPLN.toFixed(2)} PLN, Transaction Cost: ${transactionCostPLN.toFixed(2)} PLN`);
      calcLog.push('---');
    });

    // Write results per country into columns H-J
    reportSheet.getRange(`I${reportRow}`).setValue(country);
    reportSheet.getRange(`J${reportRow}`).setValue(totalRevenue.toFixed(2));
    reportSheet.getRange(`K${reportRow}`).setValue(totalTransactionCosts.toFixed(2));

    calcLog.push(`Country: ${country}, Revenue: ${totalRevenue.toFixed(2)} PLN, Transaction Costs: ${totalTransactionCosts.toFixed(2)} PLN`);
    calcLog.push('----');

    reportRow++;
  });

  // Totals row
  reportSheet.getRange(`I${reportRow}`).setValue('Total');
  reportSheet.getRange(`J${reportRow}`).setFormula(`=ROUND(SUM(J4:J${reportRow - 1}), 2)`);
  reportSheet.getRange(`K${reportRow}`).setFormula(`=ROUND(SUM(K4:K${reportRow - 1}), 2)`);
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
