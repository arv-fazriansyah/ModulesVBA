function autoMergeFiles() {
  const settingsSheetId = '1_tj2qwnxDwtQDLzUVQtPHJidSS9cz5vMwTogsFDow5s';
  const settingsSheet = SpreadsheetApp.openById(settingsSheetId).getSheetByName('SETTING');

  if (!settingsSheet) {
    throw new Error('SETTING sheet not found in the provided spreadsheet.');
  }

  const settings = settingsSheet.getDataRange().getValues();
  const headers = settings.shift();
  const jobSettings = parseJobSettings(settings, headers);

  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // Use the currently active spreadsheet for output

  jobSettings.forEach(setting => processJobSetting(activeSpreadsheet, setting));
}

function parseJobSettings(settings, headers) {
  return settings.map(setting => ({
    jobName: setting[headers.indexOf("Job Name")],
    templateId: setting[headers.indexOf("Template ID")],
    outputSheetName: setting[headers.indexOf("Data Sheet ID")],
    outputFileNameTemplate: setting[headers.indexOf("File Name")],
    outputFileType: setting[headers.indexOf("File Type")],
    folderId: getFolderId(),
    conditionals: JSON.parse(setting[headers.indexOf("Conditionals")] || '[]')
  }));
}

function getFolderId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("DATA");
  return dataSheet.getRange("AH2").getValue(); // Mengambil folder ID dari cell AH2 di sheet DATA
}

function processJobSetting(outputSpreadsheet, setting) {
  const { jobName, templateId, outputSheetName, outputFileNameTemplate, outputFileType, folderId, conditionals } = setting;
  const outputSheet = outputSpreadsheet.getSheetByName(outputSheetName);

  if (!outputSheet) {
    Logger.log(`Sheet '${outputSheetName}' not found in the active spreadsheet.`);
    return;
  }

  const outputData = outputSheet.getDataRange().getDisplayValues();
  if (outputData.length <= 1) return;

  const headerIndices = getHeaderIndices(outputSheet, outputData[0], jobName);
  const rowsToProcess = outputData.slice(1);

  rowsToProcess.forEach((row, rowIndex) => {
    const outputId = row[headerIndices.id]; // Get the Output ID for the row
    const outputUrl = row[headerIndices.url]; // Get the URL for the row

    // Check if URL is empty, then regenerate
    if (!outputId || !outputUrl || outputUrl === "") {
      if (checkConditionals(row, outputData[0], conditionals)) {
        processRow(row, outputData[0], templateId, outputFileNameTemplate, outputFileType, folderId, headerIndices, rowIndex + 2, outputSheet);
      }
    }
  });
}

function getHeaderIndices(sheet, headers, jobName) {
  const headersToAdd = {
    id: `Merged Doc ID - ${jobName}`,
    url: `${jobName}`,
    downloadLink: `Link Download - ${jobName}`,
    timestamp: `Timestamp - ${jobName}`
  };

  Object.keys(headersToAdd).forEach(key => {
    headersToAdd[key] = ensureHeader(sheet, headers, headersToAdd[key], key === 'id' ? '#e06666' : '#D3D3D3');
  });

  return headersToAdd;
}

function needsProcessing(row, headerIndices) {
  return !row[headerIndices.id] || !row[headerIndices.url] || !row[headerIndices.downloadLink] || !row[headerIndices.timestamp];
}

function processRow(row, headers, templateId, outputFileNameTemplate, outputFileType, folderId, headerIndices, rowIndex, outputSheet) {
  const outputFileName = generateFileName(row, headers, outputFileNameTemplate);
  const mergedFile = createMergedFile(templateId, row, headers, outputFileType, outputFileName);

  // Hanya proses file jika berhasil dibuat
  const fileId = saveFile(mergedFile, folderId);

  const fileUrl = mergedFile.getUrl();
  const downloadLink = generateDownloadLink(fileId, outputFileType);
  const timestamp = new Date().toLocaleString('en-GB', { hour12: false }).replace(',', '');

  const nextRow = rowIndex;
  outputSheet.getRange(nextRow, headerIndices.id + 1).setValue(fileId);
  outputSheet.getRange(nextRow, headerIndices.url + 1).setValue(fileUrl);
  outputSheet.getRange(nextRow, headerIndices.downloadLink + 1).setValue(downloadLink);
  outputSheet.getRange(nextRow, headerIndices.timestamp + 1).setValue(timestamp);

  SpreadsheetApp.flush(); // Pastikan semua perubahan diterapkan
}

function ensureHeader(sheet, headers, headerName, color) {
  let index = headers.indexOf(headerName);
  if (index === -1) {
    index = headers.length;
    headers.push(headerName);
    sheet.getRange(1, index + 1).setValue(headerName).setFontWeight('bold').setBackground(color);
  }
  return index;
}

function checkConditionals(row, headers, conditionals) {
  return conditionals.every(conditional => {
    const columnIndex = headers.indexOf(conditional.headerMap);
    return row[columnIndex] != null && row[columnIndex].toString().trim() !== '';
  });
}

function generateFileName(row, headers, template) {
  return template.replace(/<<([^>>]+)>>/g, (_, p1) => {
    const columnIndex = headers.indexOf(p1);
    return row[columnIndex] ? row[columnIndex].toString().trim() : '';
  });
}

function createMergedFile(templateId, row, headers, fileType, fileName) {
  const templateFile = DriveApp.getFileById(templateId);
  const mimeType = templateFile.getMimeType();

  // Buat salinan file dan langsung pindahkan ke folder trash
  const fileCopy = templateFile.makeCopy(fileName);
  fileCopy.setTrashed(true); // Pindahkan ke folder trash

  const file = mimeType === MimeType.GOOGLE_SLIDES
    ? createMergedSlideFile(fileCopy, row, headers, fileType, fileName)
    : createMergedDocFile(fileCopy, row, headers, fileType, fileName);

  // Cek jika nama file mengandung "RAPOR"
  if (fileName.toUpperCase().includes("RAPOR")) {
    Logger.log(`Nama file '${fileName}' mengandung kata "RAPOR". Memeriksa tabel...`);
    checkTables(file.getId());
  } else {
    Logger.log(`Nama file '${fileName}' tidak mengandung kata "RAPOR". Melewati pemeriksaan tabel.`);
  }

  // Konversi ke PDF jika diperlukan
  if (fileType.toLowerCase() === 'pdf') {
    const pdfBlob = DriveApp.getFileById(file.getId()).getAs('application/pdf');
    file.setTrashed(true); // Pindahkan file sumber ke trash setelah konversi
    return DriveApp.createFile(pdfBlob).setName(fileName);
  }

  return file;
}

function createMergedSlideFile(templateFile, row, headers, fileType, fileName) {
  const slideFile = templateFile.makeCopy(fileName);
  const slides = SlidesApp.openById(slideFile.getId());
  const slideRequests = [];

  slides.getSlides().forEach(slide => {
    headers.forEach((header, index) => {
      const text = row[index] ? row[index].toString().trim() : '';
      const request = {
        replaceAllText: {
          containsText: { text: `<<${header}>>` },
          replaceText: text
        }
      };

      if (header.toLowerCase().includes('image')) {
        const imageUrl = row[index];
        if (imageUrl) {
          slideRequests.push({
            replaceAllShapesWithImage: {
              imageUrl: imageUrl,
              containsText: { text: `<<${header}>>` }
            }
          });
        }
      } else {
        slideRequests.push(request);
      }
    });
  });

  // Make batch update once after processing all requests
  if (slideRequests.length > 0) {
    Slides.Presentations.batchUpdate({ requests: slideRequests }, slides.getId());
  }

  slides.saveAndClose();
  return slideFile;
}

function createMergedDocFile(templateFile, row, headers, fileType, fileName) {
  const docFile = templateFile.makeCopy(fileName);
  const doc = DocumentApp.openById(docFile.getId());
  const requests = [];

  const p = doc.getBody().getParent();

  for (let i = 0; i < p.getNumChildren(); i++) {
    const element = p.getChild(i);
    const t = element.getType();
    if (t === DocumentApp.ElementType.BODY_SECTION) {
      mergeSectionContent(element.asBody(), row, headers, requests);
    } else if (t === DocumentApp.ElementType.HEADER_SECTION) {
      mergeSectionContent(element.asHeaderSection(), row, headers, requests);
    } else if (t === DocumentApp.ElementType.FOOTER_SECTION) {
      mergeSectionContent(element.asFooterSection(), row, headers, requests);
    }
  }

  // Make batch update once after processing all requests
  if (requests.length > 0) {
    Docs.Documents.batchUpdate({ requests: requests }, doc.getId());
  }

  doc.saveAndClose();
  return docFile;
}

function checkTables(docId) {
  // Open the Google Doc once
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  const tables = body.getTables();

  // Get the value from cell D11 in the "1. SEKOLAH" sheet
  const sheetSekolah = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetSEKOLAH);
  const cellValueSekolah = sheetSekolah.getRange("D11").getValue();

  // If the value in D11 is "Ganjil", delete specific columns in table 6 (if it exists)
  if (cellValueSekolah === "Ganjil") {
    deleteKenaikan(tables); // Pass tables directly to deleteKenaikan
  }

  // Get the value from cell M5 in the "3. MAPEL" sheet
  const sheetMapel = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetMAPEL);
  const cellValueMapel = sheetMapel.getRange("M5").getValue();

  // Ambil tabel ke-3 dan ke-4
  const table3 = tables[2];
  const table4 = tables[3];

  // Hapus baris kosong di kolom pertama untuk kedua tabel
  deleteRow(table3, 0);
  deleteRow(table4, 0);

  // If the value in M5 is 0, delete table 4 (index 3)
  if (cellValueMapel === 0 && tables.length > 3) {
    Logger.log("Deleting Table 4 because M5 is 0...");

    // Remove Table 4
    body.removeChild(tables[3]);

    const paragraphs = body.getParagraphs();
    const paragraphIndexToRemove = 81;

    // Remove the specific paragraph if it exists
    if (paragraphs[paragraphIndexToRemove]) {
      body.removeChild(paragraphs[paragraphIndexToRemove]);
      Logger.log('Paragraph break removed at index ' + paragraphIndexToRemove);
    } else {
      Logger.log('No paragraph found at index ' + paragraphIndexToRemove);
    }
  }

  // Save and close the document
  doc.saveAndClose();
}

// Fungsi untuk menghapus baris kosong di kolom tertentu
function deleteRow(table, columnIndex) {
  for (var i = table.getNumRows() - 1; i >= 0; i--) {
    var cell = table.getRow(i).getCell(columnIndex);
    if (cell.getText().trim() === "") {
      table.removeRow(i);
    }
  }
}

// Fungsi untuk menghapus kolom 5 dan 6 pada tabel ke-6 (indeks 5)
function deleteKenaikan(tables) {
  // Check if there are at least 6 tables in the document
  if (tables.length >= 6) {
    const table = tables[5]; // Table 6 is at index 5
    Logger.log("Deleting Columns 5 and 6 in Table 6...");

    const rowCount = table.getNumRows();
    
    // Delete Column 6 (index 5) and Column 5 (index 4) across all rows
    for (let row = 0; row < rowCount; row++) {
      table.getRow(row).removeCell(5); // Removes Column 6 (index 5)
      table.getRow(row).removeCell(4); // Removes Column 5 (index 4)
    }

    Logger.log("Columns 5 and 6 have been deleted.");
  } else {
    Logger.log("Table 6 does not exist in this document.");
  }
}

function mergeSectionContent(section, row, headers, requests) {
  headers.forEach((header, index) => {
    const text = row[index] ? row[index].toString().trim() : '';
    const request = {
      replaceAllText: {
        containsText: { text: `<<${header}>>` },
        replaceText: text
      }
    };

    if (header.toLowerCase().includes('image')) {
      const imageUrl = row[index];
      if (imageUrl) {
        insertImageIntoDocTable(section, imageUrl, `<<${header}>>`);
      }
    } else {
      requests.push(request);
    }
  });
}

function insertImageIntoDocTable(section, imageUrl, placeholder) {
  const tables = section.getTables();
  const imageBlob = UrlFetchApp.fetch(imageUrl).getBlob();

  tables.forEach(table => {
    const numRows = table.getNumRows();
    for (let rowIndex = 0; rowIndex < numRows; rowIndex++) {
      const row = table.getRow(rowIndex);
      const numCells = row.getNumCells();

      for (let cellIndex = 0; cellIndex < numCells; cellIndex++) {
        const cell = row.getCell(cellIndex);
        if (cell.getText().includes(placeholder)) {
          const image = cell.insertImage(0, imageBlob);
          
          // Ambil ukuran kolom dalam point
          const columnWidthPoints = table.getColumnWidth(cellIndex); // Dapatkan lebar kolom dalam point
          
          // Sesuaikan ukuran gambar dengan ukuran kolom
          resizeImageToFitColumn(image, columnWidthPoints);
          
          cell.setText('');
        }
      }
    }
  });
}

function processHeaderAndFooter(doc, row, headers) {
  ['Header', 'Footer'].forEach(sectionName => {
    const section = doc[`get${sectionName}`]();
    const textElements = section.getText().match(/<<([^>>]+)>>/g) || [];

    textElements.forEach(placeholder => {
      const header = placeholder.replace(/<<|>>/g, '');
      const value = row[headers.indexOf(header)];

      if (header.toLowerCase().includes('image')) {
        const imageUrl = value;
        if (imageUrl) {
          insertImageIntoSection(section, placeholder, imageUrl);
        }
      } else {
        section.replaceText(placeholder, value);
      }
    });
  });
}

function insertImageIntoDocTable(section, imageUrl, placeholder) {
  const tables = section.getTables();
  const imageBlob = UrlFetchApp.fetch(imageUrl).getBlob();

  tables.forEach(table => {
    const numRows = table.getNumRows();
    for (let rowIndex = 0; rowIndex < numRows; rowIndex++) {
      const row = table.getRow(rowIndex);
      const numCells = row.getNumCells();

      for (let cellIndex = 0; cellIndex < numCells; cellIndex++) {
        const cell = row.getCell(cellIndex);
        if (cell.getText().includes(placeholder)) {
          const image = cell.insertImage(0, imageBlob);
          resizeImageToFitColumn(image, table.getColumnWidth(cellIndex)); // Resize the image
          cell.setText(''); // Clear the text placeholder
        }
      }
    }
  });
}

function resizeImageToFitColumn(image, columnWidthInPixels) {
  const pixelsToCm = 2.54 / 72;
  const columnWidthInCm = columnWidthInPixels * pixelsToCm;
  const aspectRatio = image.getWidth() / image.getHeight();

  const newWidthPixels = columnWidthInCm * 36;
  const newHeightPixels = newWidthPixels / aspectRatio;

  image.setWidth(newWidthPixels).setHeight(newHeightPixels);
}

function saveFile(file, folderId) {
  const folder = DriveApp.getFolderById(folderId);
  return file.moveTo(folder).getId();
}

function generateDownloadLink(fileId, fileType) {
  const baseLinks = {
    doc: `https://docs.google.com/document/d/${fileId}/export?format=doc`,
    pdf: `https://drive.google.com/uc?export=download&id=${fileId}`,
    ppt: `https://docs.google.com/presentation/d/${fileId}/export/pptx`
  };
  return baseLinks[fileType.toLowerCase()] || '';
}
