function oldDocFinalize(id, document, pdfFile) {
  try {
    var fileId = pdfFile['parameters']['id'];
    if (id.includes('pdfFormatIssue')) {
      var card = printMessage('<font color = #ff0000>The PDf file is not in the correct format</font>');
      return [card];
    }

    var ocrName = pdfFile['parameters']['title'].replace(/\\n/g, '').replace('\n', '');;
    if (ocrName.endsWith('.pdf')) {
      ocrName = ocrName.substring(0, ocrName.length - 4);
    }

    var doc = DriveApp.getFileById(id);
    var folder = DriveApp.createFolder(ocrName);
    var extractedResult = JSON.stringify(document);
    extractedResult = extractedResult.replace(/\\n/g, '\n');

    var jsondoc = DocumentApp.create(ocrName + ' OCR');
    var body = jsondoc.getBody();
    body.appendParagraph(extractedResult);
    DriveApp.getFileById(jsondoc.getId()).moveTo(folder);

    doc.moveTo(folder);
    if (pdfFile['parameters']['createdWith'] != undefined) {
      Drive.Files.remove(fileId);
    }
    if (pdfFile['parameters']['folderId'] != undefined) {
      var tempFolderId = pdfFile['parameters']['folderId'];
      var tempFolder = DriveApp.getFolderById(tempFolderId);
      tempFolder.setTrashed(true);
    }
  } catch ({ exception }) {
    Logger.log('oldDocFinalize Exception: ' + exception);
  }
}

function newDocFinalize(claimsId, applicantsArgumentId, amendmentId, document, pdfFile) {
  try {
    var fileId = pdfFile['parameters']['id'];
    if (claimsId.includes('pdfFormatIssue')) {
      var card = printMessage('<font color = #ff0000>The PDf file is not in the correct format</font>');
      return [card];
    }
    var claimDoc = DriveApp.getFileById(claimsId);
    var applicantsDoc = DriveApp.getFileById(applicantsArgumentId);
    var ammendmentDoc = DriveApp.getFileById(amendmentId);

    var ocrName = pdfFile['parameters']['title'].replace(/\\n/g, '').replace('\n', '');;
    if (ocrName.endsWith('.pdf')) {
      ocrName = ocrName.substring(0, ocrName.length - 4);
    }

    var folder = DriveApp.createFolder(ocrName);
    var extractedResult = JSON.stringify(document.text);
    extractedResult = extractedResult.replace(/\\n/g, '\n');


    var jsondoc = DocumentApp.create(ocrName + ' OCR');
    var body = jsondoc.getBody();
    body.appendParagraph(extractedResult);
    DriveApp.getFileById(jsondoc.getId()).moveTo(folder);

    claimDoc.moveTo(folder);
    applicantsDoc.moveTo(folder);
    ammendmentDoc.moveTo(folder);

    if (pdfFile['parameters']['createdWith'] != undefined) {
      Drive.Files.remove(fileId);
    }
    Logger.log(pdfFile['parameters']['folderId']);
    if (pdfFile['parameters']['folderId'] != undefined) {
      var tempFolderId = pdfFile['parameters']['folderId'];
      var tempFolder = DriveApp.getFolderById(tempFolderId);
      tempFolder.setTrashed(true);
    }
  } catch ({ message }) {
    Logger.log('newDocFinalize Exception: ' + message);
  }
}
