function doGet(e) {
  if (!e || !e.parameter || !e.parameter.sheet || !e.parameter.range) {
    return ContentService.createTextOutput('Silakan berikan parameter "sheet" dan "range" di URL untuk mengakses data.').setMimeType(ContentService.MimeType.TEXT);
  }

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(e.parameter.sheet);
    if (!sheet) {
      return ContentService.createTextOutput('Lembar tidak ditemukan.').setMimeType(ContentService.MimeType.TEXT);
    }

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
      var htmlTable = '<html><body><table border="1"><tr>';

      for (var k = 0; k < headers.length; k++) {
        htmlTable += '<th>' + headers[k] + '</th>';
      }

      htmlTable += '</tr>';

      for (var l = 0; l < jsonData.length; l++) {
        htmlTable += '<tr>';
        for (var m = 0; m < headers.length; m++) {
          htmlTable += '<td>' + jsonData[l][headers[m]] + '</td>';
        }
        htmlTable += '</tr>';
      }

      htmlTable += '</table></body></html>';

      return ContentService.createTextOutput(htmlTable).setMimeType(ContentService.MimeType.HTML);
    }
  } catch (error) {
    return ContentService.createTextOutput('Terjadi kesalahan: ' + error).setMimeType(ContentService.MimeType.TEXT);
  }
}
