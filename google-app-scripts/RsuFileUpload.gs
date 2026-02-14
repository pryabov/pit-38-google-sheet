function processFile(fileName, base64Data) {
  try {
    var reportObject = buildReportForRsuShareworks(base64Data);

    writeDataToSheet(reportObject);
  } catch (error) {
    Logger.log('Error processing file: ' + error.toString());
    throw new Error('Failed to process file: ' + error.toString());
  }
}

function showUploadReportDialog() {
    // Create the HTML output from the file
  const htmlOutput = HtmlService.createHtmlOutputFromFile('UploadReport.html')
    .setWidth(400)  // Set the width of the dialog
    .setHeight(300); // Set the height of the dialog
  // Show the modal dialog with a custom title
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Upload RSU Shareworks Report CSV');
}

function writeDataToSheet(reportObject) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = FIFO_SHEET_NAME;
  var sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('Sheet with name "' + sheetName + '" not found.');
  }

  // Find the first empty row in column B
  var filledRowsCount = 1;
  while (sheet.getRange('B' + filledRowsCount).getValue() && filledRowsCount < 10000) {
    filledRowsCount++;
  }
  Logger.log('First empty row: ' + filledRowsCount);

  for (var i = 0; i < reportObject.length; i++) {
    var row = filledRowsCount + i;
    var reportRow = reportObject[i];
    var totalFees = reportRow.brokerageCommission.amount + reportRow.supplementalTransactionFee.amount;

    sheet.getRange('B' + row).setValue('TEAM');
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
