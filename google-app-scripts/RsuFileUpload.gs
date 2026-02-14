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

  // htmlOutput.setTitle('Your Sidebar Title Here');
  // SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function writeDataToSheet(reportObject) {
  // Get the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('FIFO Stocks Transactions');

  if (!sheet) {
    throw new Error('Sheet with name "' + sheetName + '" not found.');
  }

  let filledRowsCount = 1;
  while (sheet.getRange(`B${filledRowsCount}`).getValue() && filledRowsCount < 10000) {
    filledRowsCount++;
  }
  Logger.log('FilledRows:' + filledRowsCount)

  // Set each value in the spreadsheet individually
  for (var i = filledRowsCount; i < reportObject.length + filledRowsCount - 1; i++) {
    let reportRow = reportObject[i - filledRowsCount];

    // sheet.getRange(i + filledRowsCount, 1).setValue(i + filledRowsCount - 1);

    sheet.getRange(`B${i}`).setValue('TEAM')
    sheet.getRange(`C${i}`).setValue('Inna')
    sheet.getRange(`D${i}`).setValue('Stany Zjednoczone Ameryki')
    sheet.getRange(`E${i}`).setValue('Sprzedaż')

    sheet.getRange(`F${i}`).setValue(reportRow.saleDate)
    sheet.getRange(`G${i}`).setValue(reportRow.sharesSold)
    sheet.getRange(`H${i}`).setValue(reportRow.salePrice.amount)
    sheet.getRange(`I${i}`).setFormula(`=G${i}*H${i}`)
    sheet.getRange(`J${i}`).setValue(reportRow.salePrice.currency)

    sheet.getRange(`K${i}`).setFormula(`=${reportRow.brokerageCommission.amount} + ${reportRow.supplementalTransactionFee.amount}`)
    sheet.getRange(`L${i}`).setValue(reportRow.brokerageCommission.currency)
  }
}