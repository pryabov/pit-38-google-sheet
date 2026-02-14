function buildReportForRsuShareworks(base64Data) {
  var data = parseCSVData(base64Data);

  // Assuming the first row contains headers
  var headers = data[0];
  
  // Find indices of the specific columns
  var salePricePerShareIndex = headers.indexOf('Sale Price Per Share');
  var sharesSoldIndex = headers.indexOf('Shares Sold');
  var brokerageCommissionIndex = headers.indexOf('Brokerage Commission');
  var supplementalTransactionFeeIndex = headers.indexOf('Supplemental Transaction Fee');
  var paymentFeeIndex = headers.indexOf('Payment Fee');
  var saleDateIndex = headers.indexOf('Sale Date');

  // Extract the specific columns for each row
  var extractedData = data.slice(1).map(function(row) {
    return {
      salePrice: formatMonetaryValue(row[salePricePerShareIndex]),
      sharesSold: row[sharesSoldIndex],
      brokerageCommission: formatMonetaryValue(row[brokerageCommissionIndex]),
      supplementalTransactionFee: formatMonetaryValue(row[supplementalTransactionFeeIndex]),
      paymentFee: formatMonetaryValue(row[paymentFeeIndex]),
      saleDate: row[saleDateIndex],
    };
  });
  return extractedData;
}

function formatMonetaryValue(value) {
  if (typeof value === 'string' && value.startsWith('$') && value.length > 1) {
    let amount = parseFloat(value.substring(1)); // Remove the '$' and parse as float
    return {
      currency: 'USD',
      amount: amount
    };
  } else {
    return null; // Return null for invalid input
  }
}
