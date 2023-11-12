function doGet(e) {
  if (!e || !e.parameter || !e.parameter.sheet || !e.parameter.range) {
    var message = "Berikan parameter berikut:\n\n" +
      "sheet=NamaLembar\n" +
      "range=A1:B10\n" +
      "format=json\n\n" +
      "Misal: ?sheet=NamaLembar&range=A1:B10&format=json";
    return ContentService.createTextOutput(message).setMimeType(ContentService.MimeType.TEXT);
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(e.parameter.sheet);
  var dataRange = sheet.getRange(e.parameter.range);
  var data = dataRange.getValues();

  var jsonData = [];
  var headers = data[0];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) {
      continue;
    }
    var rowData = {};
    for (var j = 0; j < headers.length; j++) {
      rowData[headers[j]] = row[j];
    }
    jsonData.push(rowData);
  }

  var jsonString = JSON.stringify(jsonData);

  if (e.parameter.format && e.parameter.format.toLowerCase() === "json") {
    return ContentService.createTextOutput(jsonString).setMimeType(ContentService.MimeType.JSON);
  } else {
    var message = "Silakan tambahkan parameter 'format=json' di URL untuk mendapatkan data dalam format JSON.";
    return ContentService.createTextOutput(message).setMimeType(ContentService.MimeType.TEXT);
  }
}
