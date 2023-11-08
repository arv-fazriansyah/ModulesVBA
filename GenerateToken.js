# APPSCRIPT
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Generate Token')
      .addItem('Generate', 'generateToken')
      .addToUi();
}

function generateToken() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DATA');
  var dataRange = sheet.getRange('A2:A');
  var tokenRange = sheet.getRange('F2:F');
  
  // Dapatkan data teks dari kolom A
  var dataValues = dataRange.getValues();
  
  var tokens = dataValues.map(function(row) {
    if (row[0] !== "") {
      return [generateRandomToken()];
    } else {
      return [""];
    }
  });

  // Tulis token ke kolom F
  tokenRange.setValues(tokens);
}

function generateRandomToken() {
  var characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  var charactersLength = characters.length;
  var token = '';

  for (var i = 0; i < 5; i++) {
    var randomIndex = Math.floor(Math.random() * charactersLength);
    token += characters.charAt(randomIndex);
  }

  return token;
}
