function autoMergeFiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName("SETTING");
  const settings = settingsSheet.getDataRange().getValues();
  const headers = settings.shift();

  settings.forEach(setting => {
    const jobName = setting[headers.indexOf("Job Name")];
    const templateId = setting[headers.indexOf("Template ID")];
    const outputSheetName = setting[headers.indexOf("Data Sheet ID")];
    const outputFileNameTemplate = setting[headers.indexOf("File Name")];
    const outputFileType = setting[headers.indexOf("File Type")];
    const folderId = setting[headers.indexOf("Folders")];
    const conditionals = JSON.parse(setting[headers.indexOf("Conditionals")]);

    const outputSheet = ss.getSheetByName(outputSheetName);
    const outputData = outputSheet.getDataRange().getValues();
    let outputHeaders = outputData.length > 0 ? outputData.shift() : [];

    // Create headers if they don't exist
    const idHeader = `Merged Doc ID - ${jobName}`;
    const urlHeader = `Merged Doc URL - ${jobName}`;
    const downloadLinkHeader = `Link Download - ${jobName}`;
    const timestampHeader = `Timestamp - ${jobName}`;

    let idColumn = outputHeaders.indexOf(idHeader);
    if (idColumn === -1) {
      idColumn = outputHeaders.length;
      outputHeaders.push(idHeader);
    }

    let urlColumn = outputHeaders.indexOf(urlHeader);
    if (urlColumn === -1) {
      urlColumn = outputHeaders.length;
      outputHeaders.push(urlHeader);
    }

    let downloadLinkColumn = outputHeaders.indexOf(downloadLinkHeader);
    if (downloadLinkColumn === -1) {
      downloadLinkColumn = outputHeaders.length;
      outputHeaders.push(downloadLinkHeader);
    }

    let timestampColumn = outputHeaders.indexOf(timestampHeader);
    if (timestampColumn === -1) {
      timestampColumn = outputHeaders.length;
      outputHeaders.push(timestampHeader);
    }

    if (outputHeaders.length > 0) {
      outputSheet.getRange(1, idColumn + 1, 1, 1).setValue(idHeader).setFontWeight('bold').setBackground('#e06666');
      outputSheet.getRange(1, urlColumn + 1, 1, 1).setValue(urlHeader).setFontWeight('bold').setBackground('#D3D3D3');
      outputSheet.getRange(1, downloadLinkColumn + 1, 1, 1).setValue(downloadLinkHeader).setFontWeight('bold').setBackground('#D3D3D3');
      outputSheet.getRange(1, timestampColumn + 1, 1, 1).setValue(timestampHeader).setFontWeight('bold').setBackground('#D3D3D3');
    }

    // Update headers range after potential changes
    outputHeaders = outputSheet.getRange(1, 1, 1, outputHeaders.length).getValues()[0];

    const lastColumn = outputHeaders.length;

    // Process data rows
    outputData.forEach((row, rowIndex) => {
      if (!row[idColumn] || !row[urlColumn] || !row[downloadLinkColumn] || !row[timestampColumn]) {
        if (checkConditionals(row, outputHeaders, conditionals)) {
          const outputFileName = generateFileName(row, outputHeaders, outputFileNameTemplate);
          const mergedFile = createMergedFile(templateId, row, outputHeaders, outputFileType, outputFileName);
          const fileId = saveFile(mergedFile, folderId);
          const fileUrl = `https://drive.google.com/file/d/${fileId}/view?usp=drivesdk`;
          const downloadLink = generateDownloadLink(fileId, outputFileType);
          const timestamp = new Date().toLocaleString('en-GB', { hour12: false }).replace(',', '');

          const nextRow = rowIndex + 2; // +2 because of headers and 0-based index
          outputSheet.getRange(nextRow, idColumn + 1).setValue(fileId);
          outputSheet.getRange(nextRow, urlColumn + 1).setValue(fileUrl);
          outputSheet.getRange(nextRow, downloadLinkColumn + 1).setValue(downloadLink);
          outputSheet.getRange(nextRow, timestampColumn + 1).setValue(timestamp);

          // Flush to ensure the changes are saved before proceeding
          SpreadsheetApp.flush();
        }
      }
    });
  });
}

function checkConditionals(row, headers, conditionals) {
  for (const conditional of conditionals) {
    const columnIndex = headers.indexOf(conditional.headerMap);
    if (row[columnIndex] == null || row[columnIndex].toString().trim() === '') {
      return false;
    }
  }
  return true;
}

function generateFileName(row, headers, template) {
  return template.replace(/<<([^>>]+)>>/g, (match, p1) => {
    const columnIndex = headers.indexOf(p1);
    return row[columnIndex] ? row[columnIndex].toString().trim() : '';
  });
}

function createMergedFile(templateId, row, headers, fileType, fileName) {
  const template = DriveApp.getFileById(templateId);
  const file = template.makeCopy(fileName);
  const doc = DocumentApp.openById(file.getId());
  const body = doc.getBody();

  headers.forEach((header, index) => {
    body.replaceText(`<<${header}>>`, row[index] ? row[index].toString().trim() : '');
  });

  doc.saveAndClose();

  if (fileType.toLowerCase() === 'pdf') {
    const pdfBlob = DriveApp.getFileById(file.getId()).getAs('application/pdf');
    file.setTrashed(true);
    const pdfFile = DriveApp.createFile(pdfBlob);
    pdfFile.setName(fileName);
    return pdfFile;
  }
  
  return file;
}

function saveFile(file, folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const newFile = file.moveTo(folder); // Use moveTo method to move the file to the specified folder
  return newFile.getId();
}

function generateDownloadLink(fileId, fileType) {
  if (fileType.toLowerCase() === 'doc') {
    return `https://docs.google.com/document/d/${fileId}/export?format=doc`;
  } else if (fileType.toLowerCase() === 'pdf') {
    return `https://drive.google.com/uc?export=download&id=${fileId}`;
  }
  return '';
}

function deleteFolderContents() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName("SETTING");
  const settings = settingsSheet.getDataRange().getValues();
  const headers = settings.shift();
  const folderIdColumn = headers.indexOf("Folders");

  settings.forEach(setting => {
    const folderId = setting[folderIdColumn];
    if (folderId) {
      const folder = DriveApp.getFolderById(folderId);
      const files = folder.getFiles();
      let count = 0;
      while (files.hasNext()) {
        const file = files.next();
        file.setTrashed(true);
        count++;
        // Limit batch size to 100 deletions per iteration
        if (count >= 100) {
          Utilities.sleep(500); // Pause briefly to prevent rate limiting
          count = 0; // Reset count for next batch
        }
      }
    }
  });

  // ui.alert('Isi dari semua folder yang terdaftar telah dihapus.');
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom-Menu')
    .addItem('Run AutoMerge', 'autoMergeFiles')
    .addItem('Delete isi Folder', 'deleteFolderContents')
    .addToUi();
}
