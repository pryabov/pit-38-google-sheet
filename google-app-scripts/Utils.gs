function parseCSVData(base64Data) {
  // Decode the base64 data
  var decodedBytes = Utilities.base64Decode(base64Data);
  var decodedString = Utilities.newBlob(decodedBytes).getDataAsString();

  // Parse the CSV data
  var parsedData = parseCSV(decodedString);

  return parsedData;
}

function parseCSV(csvString) {
  // Split the string into rows by new line
  var rows = csvString.trim().split('\n');
  // Split each row into columns

  var data = rows.map(function(row) {
    // Use a regular expression to correctly split the row by commas
    // while ignoring commas within quotes
    var regex = /,(?=(?:(?:[^"]*"){2})*[^"]*$)/;
    return row.split(regex).map(function(cell) {
      // Remove any surrounding quotes and trim whitespace
      return cell.trim().replace(/^"|"$/g, '');
    });
  });
  // Exclude the first row as it contains name of the file
  return data.slice(1);
}
