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
 * Parses a CSV string into a 2D array, handling quoted fields with commas.
 * The first row (file title) is excluded from the result.
 *
 * @param {string} csvString - Raw CSV content.
 * @returns {string[][]} Parsed rows and columns.
 */
function parseCSV(csvString) {
  // Split the string into rows by new line
  const rows = csvString.trim().split('\n');

  const data = rows.map(function(row) {
    // Use a regular expression to correctly split the row by commas
    // while ignoring commas within quotes
    const regex = /,(?=(?:(?:[^"]*"){2})*[^"]*$)/;
    return row.split(regex).map(function(cell) {
      // Remove any surrounding quotes and trim whitespace
      return cell.trim().replace(/^"|"$/g, '');
    });
  });
  // Exclude the first row as it contains name of the file
  return data.slice(1);
}
