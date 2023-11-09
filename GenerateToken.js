function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Generate')
      .addItem('Generate Tokens', 'generateTokens')
      .addToUi();
}

function generateTokens() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DATA');
  var dataColumn = sheet.getRange('A:A').getValues();
  var outputColumn = sheet.getRange('F:F');

  // Hapus konten di kolom output
  outputColumn.clearContent();

  // Tambahkan header "GENERATE" ke kolom output
  outputColumn.setValue([['GENERATE']]);

  // Loop melalui data di kolom A dan hasilkan token
  for (var i = 1; i < dataColumn.length; i++) {
    var cellValue = dataColumn[i][0];
    if (cellValue !== "") {
      var token = generateRandomToken(5);
      outputColumn.getCell(i + 1, 1).setValue(token);
    } else {
      outputColumn.getCell(i + 1, 1).setValue("");
    }
  }
}

function generateRandomToken(length) {
  var characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  var token = '';
  for (var i = 0; i < length; i++) {
    var randomIndex = Math.floor(Math.random() * characters.length);
    token += characters.charAt(randomIndex);
  }
  return token;
}
