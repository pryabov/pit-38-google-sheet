/**
 * Decodes a base64-encoded CSV string and parses it into a 2D array.
 * The first row of the CSV (file title) is excluded.
 *
 * @param {string} base64Data - Base64-encoded CSV file content.
 * @returns {string[][]} Parsed CSV data as a 2D array of strings.
 */
function parseCSVData(base64Data) {
  // Decode the base64 data
  const decodedBytes = Utilities.base64Decode(base64Data);
  const decodedString = Utilities.newBlob(decodedBytes).getDataAsString();

  // Parse the CSV data
  const parsedData = parseCSV(decodedString);

  return parsedData;
}

/**
 * Parses a CSV string into a 2D array using the built-in Utilities.parseCsv().
 * The first row (file title) is excluded from the result.
 *
 * @param {string} csvString - Raw CSV content.
 * @returns {string[][]} Parsed rows and columns.
 */
function parseCSV(csvString) {
  const data = Utilities.parseCsv(csvString.trim());
  // Exclude the first row as it contains name of the file
  return data.slice(1);
}
