function doPost(e) {
  // Mendapatkan nilai parameter dari URL
  var sheetName = e.parameter.sheet;
  var rangeA1Notation = e.parameter.range; // Menyimpan rentang dalam format A1 Notation
  
  // Membaca nilai POST body
  var data = JSON.parse(e.postData.contents);
  var values = data.values.slice(1); // Mengabaikan baris pertama (header)
  
  // Mendapatkan spreadsheet aktif
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  // Memastikan sheet ditemukan
  if (!sheet) {
    return ContentService
      .createTextOutput("Sheet tidak ditemukan")
      .setMimeType(ContentService.MimeType.TEXT);
  }
  
  // Mengonversi rentang dari A1 Notation ke indeks baris dan kolom
  var range = sheet.getRange(rangeA1Notation);
  var startRow = range.getRow();
  var startColumn = range.getColumn();
  
  // Mendapatkan data dari kolom pertama di spreadsheet
  var columnData = sheet.getRange(startRow, startColumn, sheet.getLastRow() - startRow + 1, 1).getValues();
  
  // Memeriksa apakah ada data yang sama di kolom pertama dari JSON
  for (var i = 0; i < values.length; i++) {
    var newData = values[i];
    var newDataFirstColumnValue = newData[0];
    for (var j = 0; j < columnData.length; j++) {
      var existingDataFirstColumnValue = columnData[j][0];
      if (existingDataFirstColumnValue === newDataFirstColumnValue) {
        // Jika ada data yang sama, ganti baris dengan data baru
        sheet.getRange(startRow + j, startColumn, 1, newData.length).setValues([newData]);
        // Menghapus data yang sudah diganti agar tidak ada duplikasi
        values.splice(i, 1);
        break;
      }
    }
  }
  
  // Menulis nilai baru ke spreadsheet pada baris kosong berikutnya
  if (values.length > 0) {
    var numRows = values.length;
    var numCols = values[0].length;
    sheet.getRange(sheet.getLastRow() + 1, startColumn, numRows, numCols).setValues(values);
  }
  
  // Menampilkan pesan sukses
  return ContentService
    .createTextOutput("Data berhasil ditambahkan dan diganti di " + sheetName)
    .setMimeType(ContentService.MimeType.TEXT);
}
