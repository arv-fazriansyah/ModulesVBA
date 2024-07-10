function autoMergeFiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName("SETTING");
  const settings = settingsSheet.getDataRange().getValues();
  const headers = settings.shift();

  settings.forEach(setting => {
    const jobName = setting[headers.indexOf("Job Name")];
    const templateId = setting[headers.indexOf("Template ID")];
    const outputSheetName = setting[headers.indexOf("Output Sheet")];
    const outputFileNameTemplate = setting[headers.indexOf("Output File Name")];
    const outputFileType = setting[headers.indexOf("Output File Type")];
    const folderId = setting[headers.indexOf("Folders ID")];
    const conditionals = JSON.parse(setting[headers.indexOf("Conditionals")]);

    const outputSheet = ss.getSheetByName(outputSheetName);
    const outputData = outputSheet.getDataRange().getValues();
    let outputHeaders = outputData.length > 0 ? outputData.shift() : [];

    // Create headers if they don't exist
    const idHeader = `Merged Doc ID - ${jobName}`;
    const urlHeader = `Merged Doc URL - ${jobName}`;
    const downloadLinkHeader = `Merged Link Download - ${jobName}`;

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

    if (outputHeaders.length > 0) {
      outputSheet.getRange(1, idColumn + 1, 1, 1).setValue(idHeader).setFontWeight('bold').setBackground('#e06666');
      outputSheet.getRange(1, urlColumn + 1, 1, 1).setValue(urlHeader).setFontWeight('bold').setBackground('#D3D3D3');
      outputSheet.getRange(1, downloadLinkColumn + 1, 1, 1).setValue(downloadLinkHeader).setFontWeight('bold').setBackground('#D3D3D3');
    }

    // Update headers range after potential changes
    outputHeaders = outputSheet.getRange(1, 1, 1, outputHeaders.length).getValues()[0];

    const lastColumn = outputHeaders.length;

    // Process data rows
    outputData.forEach((row, rowIndex) => {
      if (!row[idColumn] || !row[urlColumn] || !row[downloadLinkColumn]) {
        if (checkConditionals(row, outputHeaders, conditionals)) {
          const outputFileName = generateFileName(row, outputHeaders, outputFileNameTemplate);
          const mergedFile = createMergedFile(templateId, row, outputHeaders, outputFileType, outputFileName);
          const fileId = saveFile(mergedFile, folderId);
          const fileUrl = `https://drive.google.com/file/d/${fileId}/view?usp=drivesdk`;
          const downloadLink = generateDownloadLink(fileId, outputFileType);

          const nextRow = rowIndex + 2; // +2 because of headers and 0-based index
          outputSheet.getRange(nextRow, idColumn + 1).setValue(fileId);
          outputSheet.getRange(nextRow, urlColumn + 1).setValue(fileUrl);
          outputSheet.getRange(nextRow, downloadLinkColumn + 1).setValue(downloadLink);
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
