function autoMergeFiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let settingsSheet = ss.getSheetByName("SETTING");

  // Check if SETTINGS sheet exists, if not, create it and set headers
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet("SETTING");
    const headers = ["Job Name", "Template ID", "Data Sheet ID", "File Name", "File Type", "Folders", "Conditionals"];
    settingsSheet.appendRow(headers);
  }

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
    const outputData = outputSheet.getDataRange().getDisplayValues();
    let outputHeaders = outputData.length > 0 ? outputData.shift() : [];

    const idHeader = `Merged Doc ID - ${jobName}`;
    const urlHeader = `Merged Doc URL - ${jobName}`;
    const downloadLinkHeader = `Link Download - ${jobName}`;
    const timestampHeader = `Timestamp - ${jobName}`;

    const headerIndices = {
      id: ensureHeader(outputSheet, outputHeaders, idHeader, '#e06666'),
      url: ensureHeader(outputSheet, outputHeaders, urlHeader, '#D3D3D3'),
      downloadLink: ensureHeader(outputSheet, outputHeaders, downloadLinkHeader, '#D3D3D3'),
      timestamp: ensureHeader(outputSheet, outputHeaders, timestampHeader, '#D3D3D3')
    };

    outputHeaders = outputSheet.getRange(1, 1, 1, outputHeaders.length).getValues()[0];

    outputData.forEach((row, rowIndex) => {
      if (!row[headerIndices.id] || !row[headerIndices.url] || !row[headerIndices.downloadLink] || !row[headerIndices.timestamp]) {
        if (checkConditionals(row, outputHeaders, conditionals)) {
          const outputFileName = generateFileName(row, outputHeaders, outputFileNameTemplate);
          const mergedFile = createMergedFile(templateId, row, outputHeaders, outputFileType, outputFileName);
          const fileId = saveFile(mergedFile, folderId);
          const fileUrl = mergedFile.getUrl();
          const downloadLink = generateDownloadLink(fileId, outputFileType);
          const timestamp = new Date().toLocaleString('en-GB', { hour12: false }).replace(',', '');

          const nextRow = rowIndex + 2; // +2 because of headers and 0-based index
          outputSheet.getRange(nextRow, headerIndices.id + 1).setValue(fileId);
          outputSheet.getRange(nextRow, headerIndices.url + 1).setValue(fileUrl);
          outputSheet.getRange(nextRow, headerIndices.downloadLink + 1).setValue(downloadLink);
          outputSheet.getRange(nextRow, headerIndices.timestamp + 1).setValue(timestamp);

          SpreadsheetApp.flush();
        }
      }
    });
  });
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
  const newFile = file.moveTo(folder);
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
        if (count >= 100) {
          Utilities.sleep(500);
          count = 0;
        }
      }
    }
  });
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('CustomMenu')
    .addItem('Run AutoMerge', 'autoMergeFiles')
    .addItem('Delete isi Folder', 'deleteFolderContents')
    .addToUi();
}
