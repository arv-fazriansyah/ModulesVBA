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
    jobName: setting[0],
    templateId: setting[1],
    outputSheetName: setting[2],
    outputFileNameTemplate: setting[3],
    outputFileType: setting[4],
    folderId: setting[5],
    conditionals: JSON.parse(setting[headers.indexOf("Conditionals")])
  }));
}

function processJobSetting(ss, setting) {
  const { jobName, templateId, outputSheetName, outputFileNameTemplate, outputFileType, folderId, conditionals } = setting;
  const outputSheet = ss.getSheetByName(outputSheetName);
  const outputData = outputSheet.getDataRange().getDisplayValues();

  if (outputData.length <= 1) return;

  const headerIndices = getHeaderIndices(outputSheet, outputData[0], jobName);
  const rowsToProcess = outputData.slice(1);

  rowsToProcess.forEach((row, rowIndex) => {
    const outputId = row[headerIndices.id]; // Get the Output ID for the row

    // Check if Output ID is empty, and regenerate if necessary
    if (!outputId || outputId === "") {
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
  const file = mimeType === MimeType.GOOGLE_SLIDES ? createMergedSlideFile(templateFile, row, headers, fileType, fileName) : createMergedDocFile(templateFile, row, headers, fileType, fileName);
  
  if (fileType.toLowerCase() === 'pdf') {
    const pdfBlob = DriveApp.getFileById(file.getId()).getAs('application/pdf');
    file.setTrashed(true);
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
          resizeImageToFitColumn(image, table.getColumnWidth(cellIndex));
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

function insertImageIntoSection(section, placeholder, imageUrl) {
  const images = section.getImages();
  const imageBlob = UrlFetchApp.fetch(imageUrl).getBlob();

  images.forEach(image => {
    const parentParagraph = image.getParent();
    if (parentParagraph.getText().includes(placeholder)) {
      const insertedImage = parentParagraph.insertImage(0, imageBlob);
      resizeImageToFitColumn(insertedImage, section.getWidth());
      image.removeFromParent();
    }
  });
}

function resizeImageToFitColumn(image, columnWidth) {
  const aspectRatio = image.getWidth() / image.getHeight();
  const newWidth = Math.min(columnWidth, image.getWidth());
  const newHeight = newWidth / aspectRatio;

  image.setWidth(newWidth).setHeight(newHeight);
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
