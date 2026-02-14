/**
 * Parses a base64-encoded "Sales - Long Shares" CSV from Shareworks and extracts
 * transaction data for RSU stock sales. The CSV uses paired columns for monetary
 * values (amount column followed by currency column, e.g. "$218.33,USD").
 * The last row (summary/totals) is automatically excluded.
 *
 * @param {string} base64Data - Base64-encoded CSV file content.
 * @returns {Object[]} Array of transaction objects with: saleDate, originalAcquisitionDate,
 *   sharesSold, salePrice, originalCostBasis, brokerageCommission, supplementalTransactionFee,
 *   and withdrawalReferenceNumber.
 */
function buildReportForRsuShareworks(base64Data) {
  var data = parseCSVData(base64Data);

  // First row contains headers
  var headers = data[0];

  // Find indices of the specific columns
  var salePricePerShareIndex = headers.indexOf('Sale Price Per Share');
  var sharesSoldIndex = headers.indexOf('Shares Sold');
  var brokerageCommissionIndex = headers.indexOf('Brokerage Commission');
  var supplementalTransactionFeeIndex = headers.indexOf('Supplemental Transaction Fee');
  var saleDateIndex = headers.indexOf('Sale Date');

  // Extract data rows, excluding the last summary/totals row
  var dataRows = data.slice(1).filter(function(row) {
    return row[saleDateIndex] && row[saleDateIndex].trim() !== '';
  });

  var extractedData = dataRows.map(function(row) {
    return {
      saleDate: row[saleDateIndex],
      sharesSold: parseInt(row[sharesSoldIndex], 10),
      salePrice: formatMonetaryValue(row[salePricePerShareIndex], row[salePricePerShareIndex + 1]),
      brokerageCommission: formatMonetaryValue(row[brokerageCommissionIndex], row[brokerageCommissionIndex + 1]),
      supplementalTransactionFee: formatMonetaryValue(row[supplementalTransactionFeeIndex], row[supplementalTransactionFeeIndex + 1]),
    };
  });
  return extractedData;
}

/**
 * Parses a dollar-formatted string into a monetary object.
 * Handles comma-separated values (e.g. "$5,391.60") and reads the currency
 * from the adjacent CSV column.
 *
 * @param {string} value - Dollar-formatted string (e.g. "$218.33").
 * @param {string} currency - Currency code from the adjacent CSV column (e.g. "USD").
 * @returns {{currency: string, amount: number}} Parsed monetary value, defaults to {USD, 0} on invalid input.
 */
function formatMonetaryValue(value, currency) {
  if (typeof value === 'string' && value.startsWith('$') && value.length > 1) {
    // Remove '$' and commas before parsing (e.g. "$5,391.60" -> 5391.60)
    var amount = parseFloat(value.substring(1).replace(/,/g, ''));
    return {
      currency: (currency && currency.trim()) || 'USD',
      amount: amount
    };
  } else {
    return { currency: 'USD', amount: 0 };
  }
}
