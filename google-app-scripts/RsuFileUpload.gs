/**
 * Processes an uploaded RSU Shareworks CSV file.
 * Parses the base64-encoded CSV and writes transactions to the FIFO sheet.
 *
 * @param {string} fileName - Original filename (for logging).
 * @param {string} base64Data - Base64-encoded CSV file content.
 */
function processFile(fileName, base64Data) {
  try {
    const reportObject = buildReportForRsuShareworks(base64Data);

    writeDataToSheet(reportObject);
  } catch (error) {
    Logger.log('Error processing file: ' + error.toString());
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
 * Appends rows starting from the first empty row in column B.
 *
 * @param {Object[]} reportObject - Array of parsed transaction objects from buildReportForRsuShareworks.
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

  if (newRows.length === 0) {
    Logger.log('No new transactions to add (all duplicates).');
    return;
  }

  // Find the first empty row in column B
  let filledRowsCount = 1;
  while (sheet.getRange('B' + filledRowsCount).getValue() && filledRowsCount < 10000) {
    filledRowsCount++;
  }

  Logger.log('Adding ' + newRows.length + ' new transactions (skipped ' + (reportObject.length - newRows.length) + ' duplicates).');

  for (let i = 0; i < newRows.length; i++) {
    const row = filledRowsCount + i;
    const reportRow = newRows[i];
    const totalFees = reportRow.brokerageCommission.amount + reportRow.supplementalTransactionFee.amount;

    sheet.getRange('B' + row).setValue(RSU_SYMBOL);
    sheet.getRange('C' + row).setValue('Inna');
    sheet.getRange('D' + row).setValue('Stany Zjednoczone Ameryki');
    sheet.getRange('E' + row).setValue('Sprzedaż');

    sheet.getRange('F' + row).setValue(reportRow.saleDate);
    sheet.getRange('G' + row).setValue(reportRow.sharesSold);
    sheet.getRange('H' + row).setValue(reportRow.salePrice.amount);
    sheet.getRange('I' + row).setFormula('=G' + row + '*H' + row);
    sheet.getRange('J' + row).setValue(reportRow.salePrice.currency);

    sheet.getRange('K' + row).setValue(totalFees);
    sheet.getRange('L' + row).setValue(reportRow.brokerageCommission.currency);
  }
}
