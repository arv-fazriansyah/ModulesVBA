function autoMergeFiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = getOrCreateSettingsSheet(ss);
  const settings = settingsSheet.getDataRange().getValues();
  const headers = settings.shift();
  const jobSettings = parseJobSettings(settings, headers);

  jobSettings.forEach(setting => processJobSetting(ss, setting));
}

function getOrCreateSettingsSheet(ss) {
  let sheet = ss.getSheetByName("SETTING");
  if (!sheet) {
    sheet = ss.insertSheet("SETTING");
    sheet.appendRow(["Job Name", "Template ID", "Data Sheet ID", "File Name", "File Type", "Folders", "Conditionals"]);
  }
  return sheet;
}

function parseJobSettings(settings, headers) {
  return settings.map(setting => ({
    jobName: setting[headers.indexOf("Job Name")],
    templateId: setting[headers.indexOf("Template ID")],
    outputSheetName: setting[headers.indexOf("Data Sheet ID")],
    outputFileNameTemplate: setting[headers.indexOf("File Name")],
    outputFileType: setting[headers.indexOf("File Type")],
    folderId: setting[headers.indexOf("Folders")],
    conditionals: JSON.parse(setting[headers.indexOf("Conditionals")])
  }));
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
    url: `Merged Doc URL - ${jobName}`,
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

  // Only generate and save if it's a new file
  const fileId = saveFile(mergedFile, folderId);
  const fileUrl = mergedFile.getUrl();
  const downloadLink = generateDownloadLink(fileId, outputFileType);
  const timestamp = new Date().toLocaleString('en-GB', { hour12: false }).replace(',', '');

  const nextRow = rowIndex;
  outputSheet.getRange(nextRow, headerIndices.id + 1).setValue(fileId);
  outputSheet.getRange(nextRow, headerIndices.url + 1).setValue(fileUrl);
  outputSheet.getRange(nextRow, headerIndices.downloadLink + 1).setValue(downloadLink);
  outputSheet.getRange(nextRow, headerIndices.timestamp + 1).setValue(timestamp);

  SpreadsheetApp.flush(); // Flush to ensure all changes are applied
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

  Logger.log(`Column Width (cm): ${columnWidthInCm}`);
  Logger.log(`Original Image Width: ${image.getWidth()}`);
  Logger.log(`Original Image Height: ${image.getHeight()}`);
  Logger.log(`Resized Image Width (px): ${newWidthPixels}`);
  Logger.log(`Resized Image Height (px): ${newHeightPixels}`);

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
