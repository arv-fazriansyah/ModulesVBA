function doGet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TOKEN");
  var data = sheet.getRange("A:B").getValues();
  
  var html = '<html><head><style>table {font-family: Arial, sans-serif; border-collapse: collapse; width: 100%;} th, td {border: 1px solid #dddddd; text-align: left; padding: 8px;} th {background-color: #f2f2f2;} </style></head><body>';
  
  html += '<table>';
  
  // Header
  html += '<tr>';
  for (var i = 0; i < data[0].length; i++) {
    if (data[0][i] !== "") {
      html += '<th>' + data[0][i] + '</th>';
    }
  }
  html += '</tr>';
  
  // Data
  for (var i = 1; i < data.length; i++) {
    html += '<tr>';
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j] !== "") {
        html += '<td>' + data[i][j] + '</td>';
      }
    }
    html += '</tr>';
  }
  
  html += '</table></body></html>';
  
  return HtmlService.createHtmlOutput(html);
}
