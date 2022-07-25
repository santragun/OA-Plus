function read892Response(docAIResponse, originalFile) {
  var formFields = { "patternOne": "false", "patternTwo": "false", "patternThree": "false", "patternFour": "false" };
  var documentNumbers = [];
  var documentSection = docAIResponse.document;
  var pages = documentSection.pages;
  var blocks = pages[0].blocks;

  if (originalFile['parameters']['folderId'] != undefined) {
    var tempFolderId = originalFile['parameters']['folderId'];
    var tempFolder = DriveApp.getFolderById(tempFolderId);
    tempFolder.setTrashed(true);
  }

  for (let i = 0; i < blocks.length; i++) {
    let textSegments = blocks[i].layout.textAnchor.textSegments;
    checkCommonFields(extractTextSegments(textSegments, documentSection), formFields);
  }
  if (formFields["patternOne"] == "true" && formFields["patternTwo"] == "true" && formFields["patternThree"] == "true" && formFields["patternFour"] == "true") {

    for (let i = 0; i < pages.length; i++) {
      var block = pages[i].blocks;

      for (let j = 0; j < block.length; j++) {
        let textSegments = block[j].layout.textAnchor.textSegments;
        getDocumentNumbers(extractTextSegments(textSegments, documentSection), documentNumbers);
      }
    }

    documentNumbers.unshift("done");
    return documentNumbers;
  }

  else {
    return ['pdfFormatIssue'];
  }

}

function checkCommonFields(blockText, resultContainer) {

  if (blockText.includes("U.S. PATENT DOCUMENTS")) {
    resultContainer["patternOne"] = "true";
  }
  if (blockText.includes("Country Code-Number-Kind Code")) {
    resultContainer["patternTwo"] = "true";
  }
  if (blockText.includes("Notice of References Cited")) {
    resultContainer["patternThree"] = "true";
  }
  if (blockText.includes("Classification")) {
    resultContainer["patternFour"] = "true";
  }
}

function getDocumentNumbers(blockText, resultContainer) {
  var countryCodes = ["US-", "JP-", "KR-", "JP ", "KR "];
  for (let i = 0; i < countryCodes.length; i++) {

    while (blockText.includes(countryCodes[i])) {
      let documentNumber = blockText.substring(3, blockText.indexOf("\n"));
      if (documentNumber.length > 0) {
        if (documentNumber.includes(' ')) {
          documentNumber = documentNumber.substring(0, documentNumber.indexOf(' '));
        }
        if (documentNumber.includes('\n')) {
          documentNumber = documentNumber.substring(0, documentNumber.indexOf('\n'));
        }


        let fullDocNo = countryCodes[i] + documentNumber;
        fullDocNo = fullDocNo.replace(/,/g, '').replace(/-/g, '').replace(/\s/g, '').replace("/", '');

        resultContainer.push(fullDocNo);
      }
      blockText = blockText.substring(blockText.indexOf("\n") + 1);
    }

  }
}

function retrieveFromPatents(docNo, folder) {
  var request = UrlFetchApp.fetch("https://patents.google.com/patent/" + docNo, {
    contentType: "application/json",
    muteHttpExceptions: true,
  });

  let response = request.getContentText();
  let downloadUrl = 'content="https://patentimages.storage.googleapis.com/';
  if (request.getResponseCode() == 200) {
    if (response.includes(downloadUrl)) {
      downloadUrl = response.substring(response.indexOf(downloadUrl) + 9);
      downloadUrl = downloadUrl.substring(0, downloadUrl.indexOf("\""));
    }
    else {
      downloadUrl = "notFound";
    }
  }
  else {
    downloadUrl = "notFound";
  }
  if (!docNo.startsWith('US')) {
    Logger.log("other language");
    saveToFile(folder, "https://patents.google.com/patent/" + docNo + "/en")
  }
  return downloadUrl;
}

function saveToFile(folder, url) {
  if (!url.endsWith('en')) {
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() == 200) {

      var blob = response.getAs('application/pdf');
      var fileName = decodeURIComponent(url.split("/").pop());

      var file = folder.createFile(blob);
      file.setName(fileName);
      return "saved";
    }
    else {
      return "Document not found!";
    }
  }
  else {
    var response = UrlFetchApp.fetch(url);
    var htmlBody = response.getContentText();
    var blob = converthtmlTextToPDF(htmlBody, url);
    // var blob = Utilities.newBlob(htmlBody, 'text/html').getAs('application/pdf').setName('fileName.pdf');
    var removeEn = url.substring(0, url.length - 3);
    var fileName = decodeURIComponent(removeEn.split("/").pop());
    folder.createFile(blob).setName(fileName);
    return "saved";
  }

}

function converthtmlTextToPDF(htmlText, url) {

  const images = getAllImageTags(htmlText);
  images.forEach(image => {
    if (image.startsWith('//')) {
      image = url + image;
    }
    const res = UrlFetchApp.fetch(image);
    if (res.getResponseCode() == 200) {
      const imgbase64 = "data:" + res.getBlob().getContentType() + ';base64,' + Utilities.base64Encode(res.getBlob().getBytes());
      htmlText = replaceAll(htmlText, image, imgbase64);
    }
  });

  const blob = htmlToPDF(htmlText);

  return blob;

}

function htmlToPDF(text) {
  const html = HtmlService.createHtmlOutput(text).getContent();
  var blob = Utilities.newBlob(html, 'text/html').getAs('application/pdf');
  return blob;
}

function getAllImageTags(str) {
  let words = [];
  str.replace(/<img[^>]+src="([^">]+)"/g, function ($0, $1) { words.push($1) });
  return words;
}

function getAllScripts(str) {
  let words = [];
  str.replace(/<script[^>]+src="([^">]+)"/g, function ($0, $1) { words.push($1) });
  return words;
}

function getAllStylesheets(str) {
  let words = [];
  str.replace(/<link[^>]+rel="([^">]+)href="([^">]+)"/g, function ($0, $2) { words.push($2) });
  return words;
}




const escapeRegExp = (string) => string.replace(/[.*+\-?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
const replaceAll = (str, find, replace) => str.replace(new RegExp(escapeRegExp(find), 'g'), replace);

