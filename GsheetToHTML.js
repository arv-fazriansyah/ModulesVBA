function doGet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TOKEN");
  var dataRange = sheet.getRange("A:B");
  var data = dataRange.getValues();

  var jsonData = [];

  // Assuming the first row contains headers
  var headers = data[0];

  // Loop through the rows starting from the second row
  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    // Skip the row if the first cell is empty
    if (!row[0]) {
      continue;
    }

    var rowData = {};

    // Loop through each cell in the row
    for (var j = 0; j < headers.length; j++) {
      rowData[headers[j]] = row[j];
    }

    // Add the row data to the JSON array
    jsonData.push(rowData);
  }

  // Convert the JSON array to a string
  var jsonString = JSON.stringify(jsonData);

  // Create HTML table
  var htmlTable = '<html><body><table border="1"><tr>';

  // Add headers to the table
  for (var k = 0; k < headers.length; k++) {
    htmlTable += '<th>' + headers[k] + '</th>';
  }

  htmlTable += '</tr>';

  // Add data rows to the table
  for (var l = 0; l < jsonData.length; l++) {
    htmlTable += '<tr>';
    for (var m = 0; m < headers.length; m++) {
      htmlTable += '<td>' + jsonData[l][headers[m]] + '</td>';
    }
    htmlTable += '</tr>';
  }

  htmlTable += '</table></body></html>';

  // Set the content type and return the HTML string
  return ContentService.createTextOutput(htmlTable).setMimeType(ContentService.MimeType.HTML);
}
