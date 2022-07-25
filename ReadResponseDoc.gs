function readResponse(returnedDocument, pdfFile, functionStartedTime) {
  try {
    var resultDocument = JSON.parse(returnedDocument);

    var documentSection = resultDocument.document;
    var page01Blocks = documentSection.pages[0].blocks;
    var firstPageFields = getFirstPageFields(page01Blocks, documentSection, functionStartedTime);
    
    let scriptProperties = PropertiesService.getScriptProperties();
    
    if (!scriptProperties.getProperty('timeout') || scriptProperties.getProperty('timeoutabort')) {
      scriptProperties.deleteProperty('timeoutabort');
      if (documentSection.pages.length < 3) {
        if ((firstPageFields['applicationnumber'] == null || firstPageFields['applicationnumber'] == undefined) && (firstPageFields['filingdate'] == null || firstPageFields['filingdate'] == undefined) && (firstPageFields['firstnamedapplicant'] == null || firstPageFields['firstnamedapplicant'] == undefined)) {
          return 'pdfFormatIssue';
        }
        if (firstPageFields['applicationnumber'] != null && firstPageFields['applicationnumber'] != undefined) {
          pdfFile['parameters']['title'] = firstPageFields['applicationnumber'];
        }
        if (!firstPageFields['noticename']) {
          let secondPageFields = getSecondPageFields(page01Blocks, documentSection, firstPageFields);
          if (documentSection.pages[1].blocks) {
            secondPageFields = getSecondPageFields(documentSection.pages[1].blocks, documentSection, firstPageFields);
          }
          let claimId = generateClaimsTemplate(secondPageFields);
          let applicantsArgumentId = generateApplicantArgumentsTemplate(firstPageFields);
          let amendmentId = generateAmendmentTemplate(firstPageFields);
          if (claimId == 'exceeded_limit' || applicantsArgumentId == 'exceeded_limit' || amendmentId == 'exceeded_limit') {
            return ['exceeded_limit'];
          }
          Logger.log(amendmentId + ' ammendment');
          newDocFinalize(claimId, applicantsArgumentId, amendmentId, documentSection, pdfFile);
          let docsId = [claimId, applicantsArgumentId, amendmentId];
          return docsId;
        }
        let docId = generateOldFormatTemplate(firstPageFields);
        if (docId == 'exceeded_limit') {
          return docId;
        }
        oldDocFinalize(docId, documentSection, pdfFile);
        return docId;
      }
      else {
        if ((firstPageFields['applicationnumber'] == null || firstPageFields['applicationnumber'] == undefined) && (firstPageFields['filingdate'] == null || firstPageFields['filingdate'] == undefined) && (firstPageFields['firstnamedapplicant'] == null || firstPageFields['firstnamedapplicant'] == undefined)) {
          return ['pdfFormatIssue'];
        }
        firstPageFields = getSecondPageFields(page01Blocks, documentSection, firstPageFields);
        let secondPageFields = getSecondPageFields(page01Blocks, documentSection, firstPageFields);
        if (documentSection.pages[1].blocks) {
          secondPageFields = getSecondPageFields(documentSection.pages[1].blocks, documentSection, firstPageFields);
        }

        let claimId = generateClaimsTemplate(secondPageFields);
        let applicantsArgumentId = generateApplicantArgumentsTemplate(firstPageFields);
        let amendmentId = generateAmendmentTemplate(firstPageFields);
        if (claimId == 'exceeded_limit' || applicantsArgumentId == 'exceeded_limit' || amendmentId == 'exceeded_limit') {
          return ['exceeded_limit'];
        }

        if (firstPageFields['applicationnumber'] != null && firstPageFields['applicationnumber'] != undefined) {
          if (typeof (firstPageFields['applicationnumber']) != 'number') {
            pdfFile['parameters']['title'] = firstPageFields['applicationnumber'].replace(/\\n/g, '').replace('\n', '');
          }
        }

        newDocFinalize(claimId, applicantsArgumentId, amendmentId, documentSection, pdfFile);
        let docsId = [claimId, applicantsArgumentId, amendmentId];
        return docsId;
      }
    }
    else {
      Logger.log(pdfFile);
      return ['timeout', resultDocument, pdfFile];
    }
  } catch ({ message }) {
    Logger.log('Read Response  Exception: ' + message);
  }
}

function getAllBlocksFields(blocks, document) {
  var formFields = {};
  for (let i = 0; i < blocks.length; i++) {
    let textSegments = blocks[i].layout.textAnchor.textSegments;
    commonFirstPageFields(extractTextSegments(textSegments, document), formFields, i, blocks, document);
  }
  return formFields;
}

function extractTextSegments(textSegments, document) {
  let blockText = '';
  for (let j = 0; j < textSegments.length; j++) {
    let startIndex = +textSegments[j].startIndex;
    let endIndex = +textSegments[j].endIndex;
    blockText += document.text.substring(startIndex, endIndex);
  }
  return blockText;
}

function getDefaultValues() {
  var spreadsheetName = 'OA default values sheet';
  var defaultValues = {};
  var files = DriveApp.getFilesByName(spreadsheetName);
  while (files.hasNext()) {
    var file = files.next();
    if (file.getMimeType() == 'application/vnd.google-apps.spreadsheet') {
      try {
        var spreadsheet = SpreadsheetApp.openById(file.getId());
        var sheet = spreadsheet.getActiveSheet();
        var data = sheet.getDataRange().getValues();
        for (var i = 0; i < data.length; i++) {
          if (data[i][0].toString().toLowerCase().includes('phone')) {
            defaultValues['phone'] = data[i][1];
          }
          else if (data[i][0].toString().toLowerCase().includes('fax')) {
            defaultValues['fax'] = data[i][1];
          }
          else if (data[i][0].toString().toLowerCase().includes('email')) {
            defaultValues['email'] = data[i][1];
          }
          else if (data[i][0].toString().toLowerCase().includes('attorney name')) {
            defaultValues['attorneyName'] = data[i][1];
          }
          else if (data[i][0].toString().toLowerCase().includes('reg no')) {
            defaultValues['regNo'] = data[i][1];
          }
          else if (data[i][0].toString().toLowerCase().includes('for')) {
            defaultValues['for'] = data[i][1];
          }
          else if (data[i][0].toString().toLowerCase().includes('group art unit')) {
            defaultValues['groupArtUnit'] = data[i][1];
          }
          else if (data[i][0].toString().toLowerCase().includes('examiner')) {
            defaultValues['examiner'] = data[i][1];
          }
        }
        break;
      }
      catch (e) {
        Logger.log('Default Values Exception: ' + e);
      }
    }
  }
  return defaultValues;
}

