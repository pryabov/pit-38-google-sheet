/**
 * Processes an uploaded RSU Shareworks CSV file.
 * Parses the base64-encoded CSV and writes transactions to the FIFO sheet.
 *
 * @param {string} fileName - Original filename (for logging).
 * @param {string} base64Data - Base64-encoded CSV file content.
 */
function processFile(fileName, base64Data) {
  try {
    Logger.log('Processing file: ' + fileName);
    const reportObject = buildReportForRsuShareworks(base64Data);
    return writeDataToSheet(reportObject);
  } catch (error) {
    Logger.log('Error processing file ' + fileName + ': ' + error.toString());
    throw new Error('Failed to process file: ' + error.toString());
  }
}

/**
 * Opens a modal dialog for uploading an RSU report CSV file.
 */
function showUploadReportDialog() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('UploadReport.html')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Upload RSU Shareworks Report CSV');
}

/**
 * Writes parsed RSU transaction data to the FIFO Stocks Transactions sheet.
 * Skips duplicate transactions and appends new rows starting from the first empty row.
 *
 * @param {Object[]} reportObject - Array of parsed transaction objects from buildReportForRsuShareworks.
 * @returns {string} Status message with count of added and skipped transactions.
 */
function writeDataToSheet(reportObject) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = FIFO_SHEET_NAME;
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('Sheet with name "' + sheetName + '" not found.');
  }

  // Read existing data to detect duplicates
  const existingData = sheet.getDataRange().getValues();
  const existingKeys = new Set();
  for (let i = 1; i < existingData.length; i++) {
    const row = existingData[i];
    if (!row[FIFO_COL.symbol.index]) break;
    const dateValue = row[FIFO_COL.transactionDate.index];
    const dateKey = dateValue instanceof Date
      ? dateValue.getFullYear() + '-' + (dateValue.getMonth() + 1) + '-' + dateValue.getDate()
      : String(dateValue);
    const key = [
      row[FIFO_COL.symbol.index],
      dateKey,
      row[FIFO_COL.count.index],
      row[FIFO_COL.price.index],
      row[FIFO_COL.currency.index]
    ].join('|');
    existingKeys.add(key);
  }

  // Filter out duplicates
  const newRows = reportObject.filter(function(reportRow) {
    const parsedDate = new Date(reportRow.saleDate);
    const dateKey = parsedDate.getFullYear() + '-' + (parsedDate.getMonth() + 1) + '-' + parsedDate.getDate();
    const key = [
      RSU_SYMBOL,
      dateKey,
      reportRow.sharesSold,
      reportRow.salePrice.amount,
      reportRow.salePrice.currency
    ].join('|');
    return !existingKeys.has(key);
  });

  const skippedCount = reportObject.length - newRows.length;

  if (newRows.length === 0) {
    Logger.log('No new transactions to add (all duplicates).');
    return 'No new transactions added (' + skippedCount + ' duplicates skipped).';
  }

  // Derive first empty row from existing data
  let firstEmptyRow = existingData.length;
  for (let i = existingData.length - 1; i >= 1; i--) {
    if (existingData[i][FIFO_COL.symbol.index]) break;
    firstEmptyRow = i;
  }
  firstEmptyRow++; // Convert to 1-based sheet row

  Logger.log('Adding ' + newRows.length + ' new transactions (skipped ' + skippedCount + ' duplicates).');

  // Build batch data array (columns B through L = 11 columns)
  const batchData = newRows.map(function(reportRow) {
    const totalFees = reportRow.brokerageCommission.amount + reportRow.supplementalTransactionFee.amount;
    return [
      RSU_SYMBOL,                    // B
      RSU_STOCK_TYPE,                // C
      RSU_COUNTRY,                   // D
      SELL_OPERATION,                // E
      reportRow.saleDate,            // F
      reportRow.sharesSold,          // G
      reportRow.salePrice.amount,    // H
      '',                            // I (formula, set separately)
      reportRow.salePrice.currency,  // J
      totalFees,                     // K
      reportRow.brokerageCommission.currency  // L
    ];
  });

  // Write all rows in a single batch call (columns B:L = columns 2-12)
  sheet.getRange(firstEmptyRow, 2, batchData.length, batchData[0].length).setValues(batchData);

  // Set formulas for column I (transactionSum = count * price)
  for (let i = 0; i < newRows.length; i++) {
    const rowNum = firstEmptyRow + i;
    sheet.getRange('I' + rowNum).setFormula('=G' + rowNum + '*H' + rowNum);
  }

  return 'Added ' + newRows.length + ' transactions' + (skippedCount > 0 ? ' (' + skippedCount + ' duplicates skipped).' : '.');
}
