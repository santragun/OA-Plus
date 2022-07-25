function extraction(pdfFile, type, operation) {
  var NOW = Date.now();
  if (operation == 'multi') {
    NOW = NOW - 10000;
  }
  let authotization = checkAuthrization();
  Logger.log("Authorization " + authotization);
  if (!authotization) {
    var func = "printMessage";
    var param = '<font color="#ff0000">Thanks for using the add-on!<br>We are sorry, Your quota limit exceeded, Please subscribe to use more. You will get the payment details via email.</font>';
    var paymentDetailsAddon = callSpreadsheetDatabaseAPI('GetMailSheetData', '')[0][1][2];
    if (paymentDetailsAddon != '' && paymentDetailsAddon != null) {
      param = '<font color="#ff0000">Thanks for using the add-on!<br>' + paymentDetailsAddon + '</font>';
    }
    return [func, param];
  }
  else {
    if (!pdfFile.hasOwnProperty('parameters')) {
      pdfFile['parameters'] = {
        id: pdfFile.id,
        folderId: pdfFile.folderId,
        title: pdfFile.title,
        iconUrl: pdfFile.iconUrl,
        createdWith: pdfFile.createdWith,
        mimeType: pdfFile.mimeType
      };
    }

    var fileId = pdfFile['parameters']['id'];
    var file = DriveApp.getFileById(fileId);

    if (file != null) {
      var documentOCRResponse = {};
      var documentOCREndpoint = "https://us-documentai.googleapis.com/v1/projects/940447923436/locations/us/processors/765e9d106d9e3ec4:process";

      var decodedFile = Utilities.base64Encode(file.getBlob().getBytes());

      var scope =
        [
          "https://www.googleapis.com/auth/drive",
          "https://www.googleapis.com/auth/cloud-platform",
          "https://www.googleapis.com/auth/script.send_mail"
        ];
      var service = getOAuthService(scope);
      // service.reset();
      try {
        if (service.hasAccess()) {
          documentOCRResponse = UrlFetchApp.fetch(documentOCREndpoint, {
            method: "post",
            contentType: "application/json",
            muteHttpExceptions: true,
            payload: JSON.stringify({
              "rawDocument":
              {
                "content": decodedFile,
                "mimeType": "application/pdf"
              }
            }),
            headers: { Authorization: "Bearer " + service.getAccessToken() }
          });

        }
        else {
          var authorizationUrl = service.getAuthorizationUrl();
          var message = CardService.newTextParagraph()
            .setText('<font color="#187bcd">Authorization failed<br>Follow the following link to authorize</font>');

          if (authorizationUrl == null || authorizationUrl == '') {
            authorizationUrl = "hexon-ip.com";
          }
          var authorizeAction = CardService.newTextButton()
            .setText("Authorize")
            .setOpenLink(CardService.newOpenLink()
              .setUrl(authorizationUrl)
              .setOpenAs(CardService.OpenAs.OVERLAY)
              .setOnClose(CardService.OnClose.RELOAD));

          var mainSection = CardService.newCardSection()
            .addWidget(banner)
            .addWidget(message)
            .addWidget(authorizeAction);

          var card = CardService.newCardBuilder().addSection(mainSection).build();
          return [card];
        }
      }
      catch (error) {
        Logger.log("Document Ai exception " + error);
      }
      if (Object.keys(documentOCRResponse).length != 0) {
        if (documentOCRResponse.getResponseCode() == '200') {
          var documentOCRResult = JSON.parse(documentOCRResponse.getContentText());
          if (type == 'oa') {
            if (documentOCRResult.document.pages.length < 3) {
              let templateId = readResponse(documentOCRResponse, pdfFile, NOW);
              if (typeof (templateId) == 'string') {
                if (templateId.includes('pdfFormatIssue')) {
                  var func = "printMessage";
                  var param = '<font color="#ff0000">The pdf file is not in correct format!</font>';
                  return [func, param];
                }
                else if (templateId.includes('exceeded_limit')) {
                  var func = "printMessage";
                  var param = '<font color="#ff0000">Exception: Service invoked too many times for one day! You have finished your daily quota.</font><br><br>Please try again in the next day.';
                  return [func, param];
                }
              }
              var func = "extractionResultUi";
              return [func, templateId];
            }
            else {
              let docIds = readResponse(documentOCRResponse, pdfFile, NOW);
              if (!typeof (docIds) == 'string') {
                if (docIds[0].includes('pdfFormatIssue')) {
                  var func = "printMessage";
                  var param = '<font color="#ff0000">The pdf file is not in correct format!!</font>';
                  return [func, param];
                }
                else if (docIds[0].includes('exceeded_limit')) {
                  var func = "printMessage";
                  var param = '<font color="#ff0000">Exception: Service invoked too many times for one day! You have finished your daily quota.</font><br><br>Please try again in the next day.';
                  return [func, param];
                }
              }

              var func = "extractionResultUi";
              return [func, docIds];
            }
          }
          else {
            let documentNumbers = read892Response(documentOCRResult, pdfFile);
            var result = documentNumbers.shift();
            if (result.includes('pdfFormatIssue')) {
              var func = "printMessage";
              var param = '<font color="#ff0000">The pdf file is not in correct format!!!</font>';
              return [func, param];
            }
            if (documentNumbers.length != 0) {
              if (pdfFile['parameters']['title'].endsWith('.pdf')) {
                pdfFile['parameters']['title'] = pdfFile['parameters']['title'].substring(0, pdfFile['parameters']['title'].length - 4);
              }
              var folder = DriveApp.createFolder(pdfFile['parameters']['title']);

              identity = (x) => x;
              var tempContainer = documentNumbers.map(identity);
              var total = tempContainer.length;
              var docIds = [];

              if (operation == 'multi') {
                docIds.push('test');
              }
              else {
                for (let i = 0; i < total && isTimeLeft(NOW); i++) {
                  let url = retrieveFromPatents(documentNumbers[i], folder);
                  if (url != "notFound") {
                    stat = saveToFile(folder, url);
                  }
                  tempContainer.shift();
                }
              }

              var files = folder.getFiles();
              while (files.hasNext()) {
                file = files.next();
                docIds.push(file.getId());
              }
              if (docIds.length == 0) {
                var func = "printMessage";
                var param = '<font color="#ff0000">Couldn\'t find Patent Documents from Google patents!</font>';
                return [func, param];
              }
              else {
                docIds.unshift(folder.getId());
                docIds.unshift("892");

                if (tempContainer.length > 0) {
                  docIds.unshift(folder.getId());
                  docIds.unshift(tempContainer);
                  docIds.unshift('exceeded');
                }
                var func = "extractionResultUi";
                return [func, docIds];
              }
            }
            else {
              var func = "printMessage";
              var param = '<font color="#ff0000">Couldn\'t find Patent Documents in the file!!</font>';
              return [func, param];
            }
          }
        }
        else if (documentOCRResponse.getResponseCode() == '400') {
          var documentOCRResult = JSON.parse(documentOCRResponse.getContentText());
          if (pdfFile['parameters']['folderId'] != undefined) {
            var tempFolderId = pdfFile['parameters']['folderId'];
            var tempFolder = DriveApp.getFolderById(tempFolderId);
            tempFolder.setTrashed(true);
          }
          if (documentOCRResult.error.message.toString().includes('Document pages exceed the limit')) {
            var func = "printMessage";
            var param = '<font color="#ff0000">Maximum no of pages allowed is 10</font>';
            return [func, param];
          }
          else {
            var func = "printMessage";
            var param = '<font color="#ff0000">Please try again!</font>';
            return [func, param];
          }
        }
        else if (JSON.parse(documentOCRResponse).error &&
          JSON.parse(documentOCRResponse).error.code == '403' &&
          JSON.parse(documentOCRResponse).error.message.includes('Request had insufficient authentication scopes')) {

          var authorizationUrl = service.getAuthorizationUrl();
          var message = CardService.newTextParagraph()
            .setText('<font color="#187bcd">Authorization failed<br>Follow the following link to authorize</font>');

          if (authorizationUrl == null || authorizationUrl == '') {
            authorizationUrl = "hexon-ip.com";
          }
          var authorizeAction = CardService.newTextButton()
            .setText("Authorize")
            .setOpenLink(CardService.newOpenLink()
              .setUrl(authorizationUrl)
              .setOpenAs(CardService.OpenAs.OVERLAY)
              .setOnClose(CardService.OnClose.RELOAD));

          var mainSection = CardService.newCardSection()
            .addWidget(banner)
            .addWidget(message)
            .addWidget(authorizeAction);

          return ['authorize', CardService.newCardBuilder().addSection(mainSection)];

        }
        else {
          if (pdfFile['parameters']['folderId'] != undefined) {
            var tempFolderId = pdfFile['parameters']['folderId'];
            var tempFolder = DriveApp.getFolderById(tempFolderId);
            tempFolder.setTrashed(true);
          }
          Logger.log(documentOCRResponse);
          var func = "printMessage";
          var param = '<font color="#ff0000">Please try again!!</font>';
          return [func, param];
        }
      }
      else {
        Utilities.sleep(5000);
        var func = "printMessage";
        var param = '<font color="#ff0000">Please try again!!!</font>';
        return [func, param];
      }
    }
    else {
      var func = "printMessage";
      var param = '<font color="#ff0000">File not found!</font>';
      return [func, param];
    }
  }
}
