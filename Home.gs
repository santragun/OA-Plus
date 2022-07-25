function singleFileOperation(fileId, folder, option, operation) {
  try {
    if (DriveApp.getFileById(fileId)) {
      var pdfFile = DriveApp.getFileById(fileId);

      if (pdfFile != null) {
        if (pdfFile.getMimeType() != 'application/pdf') {
          var card = printMessage('<font color="#ff0000">Please select PDF file only</font>');
          return [card];
        }
        else {
          var jsonFile = {
            'mimeType': pdfFile.getMimeType(),
            'title': pdfFile.getName(),
            'iconUrl': "https://drive-thirdparty.googleusercontent.com/32/type/application/pdf",
            'id': pdfFile.getId(),
            'createdWith': 'uploadFile',
            'folderId': folder.getId()
          };
          var uploadedFile = JSON.parse(JSON.stringify(jsonFile));

          if (option == "oa") {
            var response = extraction(uploadedFile, 'oa', operation);
          }
          else {
            var response = extraction(uploadedFile, '892', operation);
          }
          return response;
        }
      }
      else {
        var card = selectOption();
        return [card];
      }
    }
  } catch ({ message }) {
    Logger.log('Sigle File Operation Exception: ' + message);
    return printMessage('<br><font color="#ff0000">Please try again later, &nbsp;Sorry for the Incovinience!</font>');
  }
}

function homePage() {
  try {
    var footer = footersCardSection();
    if (typeof (footer[0]) == 'number' || typeof (footer[0]) != 'string') {
      try {
        var scriptProperties = PropertiesService.getScriptProperties();
        var option = scriptProperties.getProperty("option");

        var clicked = PropertiesService.getScriptProperties().getProperty('addFile');
        var fileName = "local files.txt";
        var card = selectOption();
        var folderList = DriveApp.getFoldersByName(folderName);

        if (folderList.hasNext()) {
          var folder = folderList.next();
          var fileList = folder.getFilesByName(fileName);
          if (fileList.hasNext()) {
            var txtFile = fileList.next();
            var id = txtFile.getBlob().getDataAsString();

            if (clicked == null) {
              scriptProperties.setProperty('addFile', 'no');
              scriptProperties.setProperty('firstFileId', id);
              var card = dialog(scriptProperties.getProperty("option"));
              return card;
            }
            else if (clicked == 'no') {
              try {
                scriptProperties.deleteProperty('addFile');
                Drive.Files.remove(txtFile.getId());

                var response = singleFileOperation(id, folder, option, 'single');
                if (response[0] == "printMessage") {
                  var card = printMessage(response[1]);
                  return [card];
                }
                else if (response[0] == "authorize") {
                  var card = response[1];
                  return [card.build()];
                }
                else {
                  var widget;
                  if (response[1][0] == 'exceeded') {
                    widget = executionTimeExceeded(response[1], ['one']);
                    return CardService.newCardBuilder().addSection(bannerSection).addSection(widget).setFixedFooter(CardService.newFixedFooter().setPrimaryButton(homeBtn)).build();
                  }
                  else if (response[1][0] == 'timeout') {
                    widget = OAexecutionTimeExceeded(response[1][1], response[1][2], ['one']);
                    return CardService.newCardBuilder().addSection(bannerSection).addSection(widget).setFixedFooter(CardService.newFixedFooter().setPrimaryButton(homeBtn)).build();
                  }
                  else {
                    updateCustomerDetails();
                    footer = footersCardSection();
                    widget = extractionResultUi(response[1]);
                    return CardService.newCardBuilder().addSection(widget).addSection(footer[0]).addSection(footer[1]).setFixedFooter(CardService.newFixedFooter().setPrimaryButton(homeBtn)).build();
                  }
                }
              } catch ({ message }) {
                Logger.log('Only one execute Exception: ' + message);
              }
            }
            else {
              try {
                scriptProperties.deleteProperty('addFile');
                Drive.Files.remove(txtFile.getId());

                var previousId = scriptProperties.getProperty('firstFileId');

                var firstResponse = singleFileOperation(previousId, folder, negate(option), 'multi');
                var secondResponse = singleFileOperation(id, folder, option, 'multi');
                var errorMessage, resultSection, exceed, firstSection, secondSection;

                if (firstResponse[0] == "printMessage") {
                  errorMessage = CardService.newDecoratedText()
                    .setTopLabel('Something went wrong with ' + negate(option).toUpperCase() + '...')
                    .setText(firstResponse[1])
                    .setWrapText(true);
                  firstSection = CardService.newCardSection().addWidget(errorMessage);
                }
                else {
                  if (firstResponse[1][0] == 'exceeded') {
                    exceed = true;
                    firstSection = executionTimeExceeded(firstResponse[1], secondResponse);
                    scriptProperties.setProperty('parallelExtraction892', 'first');
                  }
                  else if (firstResponse[1][0] == 'timeout') {
                    exceed = true;
                    firstSection = OAexecutionTimeExceeded(firstResponse[1][1], firstResponse[1][2], secondResponse);
                    scriptProperties.setProperty('parallelExtractionOA', 'first');
                  }
                  else {
                    updateCustomerDetails();
                    firstSection = extractionResultUi(firstResponse[1]);
                  }
                }

                if (secondResponse[0] == "printMessage") {
                  errorMessage = CardService.newDecoratedText()
                    .setTopLabel('Something went wrong with ' + option.toUpperCase() + '...')
                    .setText(secondResponse[1])
                    .setWrapText(true);
                  secondSection = CardService.newCardSection().addWidget(errorMessage);
                }
                else {
                  if (secondResponse[1][0] == 'exceeded') {
                    exceed = true;
                    secondSection = executionTimeExceeded(secondResponse[1], firstResponse);
                    scriptProperties.setProperty('parallelExtraction892', 'second');
                  }
                  else if (secondResponse[1][0] == 'timeout') {
                    exceed = true;
                    secondSection = OAexecutionTimeExceeded(secondResponse[1][1], secondResponse[1][2], firstResponse);
                    scriptProperties.setProperty('parallelExtractionOA', 'second');
                  }
                  else {
                    updateCustomerDetails();
                    secondSection = extractionResultUi(secondResponse[1]);
                  }
                }

                resultSection = CardService.newCardBuilder().addSection(firstSection).addSection(secondSection);
                if (!exceed) {
                  footer = footersCardSection();
                  resultSection.addSection(footer[0]).addSection(footer[1]);
                }
                return resultSection.setFixedFooter(CardService.newFixedFooter().setPrimaryButton(homeBtn)).build();
              }
              catch ({ message }) {
                Logger.log('Two files Execution: ' + message);
              }
            }
          }
          else {
            if (clicked == 'no') {
              scriptProperties.deleteProperty('addFile');
              let id = scriptProperties.getProperty('firstFileId');
              var response = singleFileOperation(id, folder, option, 'single');

              try {
                if (response[0] == "printMessage") {
                  var card = printMessage(response[1]);
                  return [card];
                }
                else if (response[0] == "authorize") {
                  var card = response[1];
                  return [card.build()];
                }
                else {
                  var widget;
                  if (response[1][0] == 'exceeded') {
                    widget = executionTimeExceeded(response[1], ['one']);
                    return CardService.newCardBuilder().addSection(bannerSection).addSection(widget).setFixedFooter(CardService.newFixedFooter().setPrimaryButton(homeBtn)).build();
                  }
                  else if (response[1][0] == 'timeout') {
                    widget = OAexecutionTimeExceeded(response[1][1], response[1][2], ['one']);
                    return CardService.newCardBuilder().addSection(bannerSection).addSection(widget).setFixedFooter(CardService.newFixedFooter().setPrimaryButton(homeBtn)).build();
                  }
                  else {
                    updateCustomerDetails();
                    footer = footersCardSection();
                    widget = extractionResultUi(response[1]);
                    return CardService.newCardBuilder().addSection(widget).addSection(footer[0]).addSection(footer[1]).setFixedFooter(CardService.newFixedFooter().setPrimaryButton(homeBtn)).build();
                  }
                }
              } catch ({ message }) {
                Logger.log('Only one execute Exception: ' + message);
              }
            }
            else if (clicked == 'yes') {
              scriptProperties.deleteProperty('addFile');
            }
            return [card];
          }
        }
        else {
          return [card];
        }
      } catch ({ message }) {
        Logger.log('HomePage First If(start) Exception: ' + message);
      }
    }
    else {
      try {
        var message = CardService.newTextParagraph()
          .setText('<font color="#187bcd">Authorization failed<br>Follow the following link to authorize</font>');
        var authUrl = footer[0].substring(7);
        if (authUrl == null || authUrl == '') {
          authUrl = "hexon-ip.com";
        }
        var authorizeAction = CardService.newTextButton()
          .setText("Authorize")
          .setOpenLink(CardService.newOpenLink()
            .setUrl(authUrl)
            .setOpenAs(CardService.OpenAs.OVERLAY)
            .setOnClose(CardService.OnClose.RELOAD));

        var mainSection = CardService.newCardSection()
          .addWidget(banner)
          .addWidget(message)
          .addWidget(authorizeAction);

        return CardService.newCardBuilder()
          .addSection(mainSection)
          .build();
      } catch ({ message }) {
        Logger.log('HomePage Else(API not working) Exception: ' + message);
      }
    }
  } catch ({ message }) {
    Logger.log('HomePage Start Exception: ' + message);
    return printMessage('<br><font color="#ff0000">Please try again later, &nbsp;Sorry for the Incovinience!</font>');
  }

}

function addOther(option) {
  let id = PropertiesService.getScriptProperties().getProperty('firstFileId');
  PropertiesService.getScriptProperties().setProperty('addFile', 'yes');
  return initialHome(option);
}

function onlyOneFile() {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('addFile', 'no');
  var card = homePage();
  var nav = CardService.newNavigation()
    .popToRoot().updateCard(card);

  return CardService.newActionResponseBuilder()
    .setNavigation(nav)
    .build();
}
function back(option) {
  var folderList = DriveApp.getFoldersByName(folderName);

  if (folderList.hasNext()) {
    var folder = folderList.next();
    folder.setTrashed(true);
  }
  var previous = initialHome(option);
  return previous;
}
function dialog(option) {

  const onYesBtnClicked = CardService.newAction()
    .setFunctionName("addOther")
    .setParameters({ "parameter": negate(option) })
    .setLoadIndicator(CardService.LoadIndicator.SPINNER);

  const yesBtn = CardService.newTextButton()
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setBackgroundColor('#28a745')
    .setText("Yes")
    .setOnClickAction(onYesBtnClicked);

  const onNoBtnClicked = CardService.newAction()
    .setFunctionName("homePage")
    .setLoadIndicator(CardService.LoadIndicator.SPINNER);

  const noBtn = CardService.newTextButton()
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setBackgroundColor('#dc3545')
    .setText("No")
    .setOnClickAction(onNoBtnClicked);

  const onBackBtnClicked = CardService.newAction()
    .setFunctionName("back")
    .setParameters({ "parameter": option })
    .setLoadIndicator(CardService.LoadIndicator.SPINNER);

  const backBtn = CardService.newTextButton()
    .setTextButtonStyle(CardService.TextButtonStyle.TEXT)
    .setText("Back")
    .setOnClickAction(onBackBtnClicked);

  let choice = negate(option).toUpperCase();
  if (choice == '892') {
    choice = 'PTO-892';
  }
  var mainSection = CardService.newCardSection()
    .addWidget(yesBtn)
    .addWidget(noBtn)
    .setHeader('Do you want to upload ' + choice + ' PDF too?');

  let cardFooter = CardService.newFixedFooter()
    .setPrimaryButton(homeBtn)
    .setSecondaryButton(backBtn);

  return CardService.newCardBuilder()
    .addSection(bannerSection)
    .addSection(mainSection)
    .setFixedFooter(cardFooter)
    .build();

}

function selectOption() {
  var footer = footersCardSection();
  if (typeof (footer[0]) == 'number' || typeof (footer[0]) != 'string') {

    if (PropertiesService.getScriptProperties().getProperty('addFile') == 'no') {
      PropertiesService.getScriptProperties().deleteProperty('addFile')
    }

    let onOASubmitAction = CardService.newAction()
      .setFunctionName("initialHome")
      .setParameters({ "parameter": "oa" })
      .setLoadIndicator(CardService.LoadIndicator.SPINNER);

    let on892SubmitAction = CardService.newAction()
      .setFunctionName("initialHome")
      .setParameters({ "parameter": "892" })
      .setLoadIndicator(CardService.LoadIndicator.SPINNER);

    let processOAButton = CardService.newTextButton()
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#187bcd')
      .setText("OA")
      .setOnClickAction(onOASubmitAction);

    let process892Button = CardService.newTextButton()
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#187bcd')
      .setText("892")
      .setOnClickAction(on892SubmitAction);

    let mainSectionOAText = CardService.newDecoratedText()
      .setText('<font color="#566573">Process Office Action(OA) file</font>')
      .setWrapText(true)
      .setIconUrl('https://icon-library.com/images/pdf-icon-transparent/pdf-icon-transparent-11.jpg')
      .setButton(processOAButton);

    let mainSection892Text = CardService.newDecoratedText()
      .setWrapText(true)
      .setText('<font color="#566573">Process PTO-892 file</font>')
      .setIconUrl('https://icon-library.com/images/pdf-icon-transparent/pdf-icon-transparent-11.jpg')
      .setButton(process892Button);

    let mainSection = CardService.newCardSection()
      .addWidget(mainSectionOAText)
      .addWidget(mainSection892Text);

    let headerSection = CardService.newCardSection()
      .addWidget(banner);

    return CardService.newCardBuilder()
      .setName("OA+")
      .addSection(headerSection)
      .addSection(mainSection)
      .addSection(footer[0])
      .addSection(footer[1])
      .build();
  }
  else {
    var message = CardService.newTextParagraph()
      .setText('<font color="#ff0000">Authorization failed</font><br><font color="#187bcd">Click on the following link to authorize</font>');

    var authorizeAction = CardService.newTextButton()
      .setText("Authorize")
      .setOpenLink(CardService.newOpenLink()
        .setUrl(footer[0].substring(7))
        .setOpenAs(CardService.OpenAs.OVERLAY)
        .setOnClose(CardService.OnClose.RELOAD));

    var mainSection = CardService.newCardSection()
      .addWidget(banner)
      .addWidget(message)
      .addWidget(authorizeAction);

    return CardService.newCardBuilder()
      .addSection(mainSection)
      .build();
  }
}

function initialHome(options) {
  var footer = footersCardSection();
  if (typeof (footer[0]) == 'number' || typeof (footer[0]) != 'string') {
    var scriptProperties = PropertiesService.getScriptProperties();
    var option = options.parameters.parameter;
    var text = "";
    if (option == "oa") {
      text = "Office Action";
      scriptProperties.setProperty('option', 'oa');
    }
    else {
      text = "PTO-892";
      scriptProperties.setProperty('option', '892');
    }

    let selectAction = CardService.newTextButton()
      .setText("Select")
      .setBackgroundColor('#187bcd')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOpenLink(CardService.newOpenLink()
        .setUrl(WebAppUrl + "?picker=select")
        .setOpenAs(CardService.OpenAs.OVERLAY)
        .setOnClose(CardService.OnClose.RELOAD));

    let uploadAction = CardService.newTextButton()
      .setText("Upload")
      .setBackgroundColor('#187bcd')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOpenLink(CardService.newOpenLink()
        .setUrl(WebAppUrl)
        .setOpenAs(CardService.OpenAs.OVERLAY)
        .setOnClose(CardService.OnClose.RELOAD));

    let mainSectionDriveText = CardService.newDecoratedText()
      .setText('<font color="#566573">Select ' + text + ' pdf from Google Drive</font>')
      .setWrapText(true)
      .setIconUrl('https://icon-library.com/images/google-drive-icon-png/google-drive-icon-png-27.jpg')
      .setButton(selectAction);

    let mainSectionUploadText = CardService.newDecoratedText()
      .setText('<font color="#566573">Upload ' + text + ' pdf from Computer</font>')
      .setWrapText(true)
      .setIconUrl('https://icon-library.com/images/upload-icon-image/upload-icon-image-19.jpg')
      .setButton(uploadAction);

    let mainSection = CardService.newCardSection()
      .addWidget(mainSectionDriveText)
      .addWidget(mainSectionUploadText);

    let backBtn = CardService.newTextButton()
      .setText("Back")
      .setOnClickAction(homeAction);

    return CardService.newCardBuilder()
      .addSection(mainSection)
      .addSection(footer[0])
      .addSection(footer[1])
      .setFixedFooter(CardService.newFixedFooter().setPrimaryButton(backBtn))
      .build();
  }
  else {
    var message = CardService.newTextParagraph()
      .setText('<font color="#ff0000">Authorization failed</font><br><font color="#187bcd">Click on the following link to authorize</font>');

    var authorizeAction = CardService.newTextButton()
      .setText("Authorize")
      .setOpenLink(CardService.newOpenLink()
        .setUrl(footer[0].substring(7))
        .setOpenAs(CardService.OpenAs.OVERLAY)
        .setOnClose(CardService.OnClose.RELOAD));

    var mainSection = CardService.newCardSection()
      .addWidget(banner)
      .addWidget(message)
      .addWidget(authorizeAction);

    return CardService.newCardBuilder()
      .addSection(mainSection)
      .build();
  }
}
function onDriveFilesSelected(event) {
  var file = event.drive.selectedItems[0];
  if (event.drive.selectedItems.length != 1) {
    var card = printMessage('<font color="#ff0000">Please select one Office Action PDF at a time.</font>');
    return [card];
  }
  else if (file['mimeType'] == 'application/pdf') {
    var card = selectedFilePage(file);
    return [card];
  }
  else {
    var card = printMessage('<font color="#ff0000">Please select PDF file type only</font>');
    return [card];
  }
}
function mediator(parameters) {
  var response = extraction(JSON.parse(parameters.parameters.file), parameters.parameters.type, 'single');
  if (response[0] == "printMessage") {
    var card = printMessage(response[1]);
    return [card];
  }
  else {
    var widget = extractionResultUi(response[1]);
    return CardService.newCardBuilder().addSection(widget).build();
  }

}
function selectedFilePage(file) {
  if (file.addonHasFileScopePermission == true) {
    var card = printMessage('<font color="#ff0000">Please select other Office Action PDF</font>');
    return card;
  }
  var title = CardService.newTextParagraph()
    .setText('<font color="#187bcd">Generate Google doc template data from Office Action PDF</font>');

  var fileName = CardService.newTextParagraph()
    .setText(file['title']);
  var onOAAction = CardService.newAction()
    .setFunctionName("mediator")
    .setParameters({ "file": JSON.stringify(file), "type": "oa" })
    .setLoadIndicator(CardService.LoadIndicator.SPINNER);

  var on892Action = CardService.newAction()
    .setFunctionName("mediator")
    .setParameters({ "file": JSON.stringify(file), "type": "892" })
    .setLoadIndicator(CardService.LoadIndicator.SPINNER);

  var aoButton = CardService.newTextButton()
    .setText("OA")
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setBackgroundColor('#187bcd')
    .setOnClickAction(onOAAction);

  var button892 = CardService.newTextButton()
    .setText("892")
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setBackgroundColor('#187bcd')
    .setOnClickAction(on892Action);

  var mainSection = CardService.newCardSection()
    .addWidget(banner)
    .addWidget(title)
    .addWidget(fileName)
    .addWidget(aoButton)
    .addWidget(button892);

  var footerSection = CardService.newCardSection()
    .addWidget(homeBtn);

  if (!footersCardSection()[0].includes("failed")) {
    return CardService.newCardBuilder()
      .addSection(mainSection)
      .addSection(footersCardSection()[0])
      .addSection(footerSection)
      .build();
  }
  else {
    var message = CardService.newTextParagraph()
      .setText('<font color="#ff0000">Authorization failed</font><br><font color="#187bcd">Click on the following link to authorize</font>');

    var authorizeAction = CardService.newTextButton()
      .setText("Authorize")
      .setOpenLink(CardService.newOpenLink()
        .setUrl(footer[0].substring(7))
        .setOpenAs(CardService.OpenAs.OVERLAY)
        .setOnClose(CardService.OnClose.RELOAD));

    var mainSection = CardService.newCardSection()
      .addWidget(banner)
      .addWidget(message)
      .addWidget(authorizeAction);

    return CardService.newCardBuilder()
      .addSection(mainSection)
      .build();
  }
}

function printMessage(message) {
  var footer = footersCardSection();
  var message = CardService.newDecoratedText()
    .setTopLabel('Something went wrong...')
    .setText(message)
    .setWrapText(true);
  var mainSection = CardService.newCardSection()
    .addWidget(banner)
    .addWidget(message);

  if (footer && typeof (footer[0]) != 'string') {
    return CardService.newCardBuilder()
      .addSection(mainSection)
      .addSection(footer[0])
      .addSection(footer[1])
      .setFixedFooter(CardService.newFixedFooter().setPrimaryButton(homeBtn))
      .build();
  }
  else {
    return CardService.newCardBuilder()
      .addSection(mainSection)
      .build();
  }
}

function continueWithDownloaded(downloaded) {
  try {
    var footer = footersCardSection();
    if (typeof (footer[0]) == 'number' || typeof (footer[0]) != 'string') {
      var remaining = JSON.parse(downloaded.parameters.parameter);
      var oaContinue = JSON.parse(downloaded.parameters.oaResponse);
      remaining.shift();
      remaining.shift();
      remaining.shift();

      Logger.log(remaining);


      let oaSection, oaFinal = true, final892 = true;
      let parallelExecution892 = PropertiesService.getScriptProperties().getProperty('parallelExtraction892');
      let card = CardService.newCardBuilder().setFixedFooter(CardService.newFixedFooter().setPrimaryButton(homeBtn));

      if (oaContinue[0] != 'one') {
        if (oaContinue[0] == "printMessage") {
          oaSection = CardService.newCardSection().addWidget(
            CardService.newDecoratedText()
              .setTopLabel('Something went wrong with OA...')
              .setText(oaContinue[1])
              .setWrapText(true));

        }
        else if (oaContinue[1][0] == 'timeout') {
          oaSection = OAexecutionTimeExceeded(oaContinue[1][1], oaContinue[1][2], remaining);
          oaFinal = false;
        }
        else {
          oaSection = extractionResultUi(oaContinue);
        }
        Logger.log(remaining);
        let widget = extractionResultUi(remaining);
        Logger.log(remaining);
        if (parallelExecution892 == 'first') {
          card.addSection(widget).addSection(oaSection);
        }
        else {
          card.addSection(oaSection).addSection(widget);
        }
      }
      else {
        card.addSection(widget);
      }
      updateCustomerDetails();
      if (oaFinal && final892) {
        footer = footersCardSection();
        card.addSection(footer[0]).addSection(footer[1]);
      }
      return card.build();
    }
    else {
      var message = CardService.newTextParagraph()
        .setText('<font color="#ff0000">Authorization failed</font><br><font color="#187bcd">Click on the following link to authorize</font>');

      var authorizeAction = CardService.newTextButton()
        .setText("Authorize")
        .setOpenLink(CardService.newOpenLink()
          .setUrl(footer[0].substring(7))
          .setOpenAs(CardService.OpenAs.OVERLAY)
          .setOnClose(CardService.OnClose.RELOAD));

      var mainSection = CardService.newCardSection()
        .addWidget(banner)
        .addWidget(message)
        .addWidget(authorizeAction);

      return CardService.newCardBuilder()
        .addSection(mainSection)
        .build();
    }
  } catch ({ message }) {
    Logger.log('Continue With Downloaded Exception: ' + message);
  }
}

function executeRemaining(remainingsParam) {
  try {
    var remainings = JSON.parse(remainingsParam.parameters.parameter);
    var oaContinue = JSON.parse(remainingsParam.parameters.oaResponse);
    remainings.shift();
    var remainingDocs = remainings.shift();
    var folderId = remainings.shift();
    var folder = DriveApp.getFolderById(folderId);
    var tempContainer = remainingDocs;
    const NOW = Date.now();
    for (let i = 0; i < remainingDocs.length && isTimeLeft(NOW); i++) {

      if (remainingDocs[i] != null) {
        let url = retrieveFromPatents(remainingDocs[i], folder);
        if (url != "notFound") {
          stat = saveToFile(folder, url);
        }
      }
      tempContainer = tempContainer.slice(1);
    }

    if (tempContainer.length > 0) {
      remainings = ['exceeded', tempContainer, folderId, "892", folderId];
    }

    let parallelExecution892 = PropertiesService.getScriptProperties().getProperty('parallelExtraction892');
    let oaSection, widget, oaFinal = true, final892 = true;

    if (oaContinue[0] != 'one') {
      if (oaContinue[0] == "printMessage") {
        oaSection = CardService.newCardSection().addWidget(
          CardService.newDecoratedText()
            .setTopLabel('Something went wrong with OA...')
            .setText(oaContinue[1])
            .setWrapText(true));
      }
      else if (oaContinue[1][0] == 'timeout') {
        oaSection = OAexecutionTimeExceeded(oaContinue[1][1], oaContinue[1][2], remainings);
        oaFinal = false;
      }
      else {
        oaSection = extractionResultUi(oaContinue);
      }
    }
    Logger.log(remainings);
    if (tempContainer.length > 0) {
      widget = executionTimeExceeded(remainings, oaContinue);
      final892 = false;
    }
    else {
      updateCustomerDetails();
      widget = extractionResultUi(remainings);
    }
    var card = CardService.newCardBuilder().addSection(bannerSection).setFixedFooter(CardService.newFixedFooter().setPrimaryButton(homeBtn));
    if (oaContinue[0] == 'one') {
      card.addSection(widget);
    }
    else {
      if (parallelExecution892 == 'first') {
        card.addSection(widget).addSection(oaSection);
      }
      else {
        card.addSection(oaSection).addSection(widget);
      }
    }

    if (oaFinal && final892) {
      let footer = footersCardSection();
      card.addSection(footer[0]).addSection(footer[1]);
    }
    return card.build();
  } catch ({ message }) {
    Logger.log('Execute Remaining Exception: ' + message);
    return printMessage('<br><font color="#ff0000">Please try again later, &nbsp;Sorry for the Incovinience!</font>');
  }
}
function executionTimeExceeded(remaining, oaResponse) {

  let onYesBtnClicked = CardService.newAction()
    .setFunctionName("executeRemaining")
    .setParameters({
      "parameter": JSON.stringify(remaining),
      "oaResponse": JSON.stringify(oaResponse)
    })
    .setLoadIndicator(CardService.LoadIndicator.SPINNER);

  let yesBtn = CardService.newTextButton()
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setBackgroundColor('#28a745')
    .setText("Yes")
    .setOnClickAction(onYesBtnClicked);

  let onNoBtnClicked = CardService.newAction()
    .setFunctionName("continueWithDownloaded")
    .setParameters({
      "parameter": JSON.stringify(remaining),
      "oaResponse": JSON.stringify(oaResponse)
    })
    .setLoadIndicator(CardService.LoadIndicator.SPINNER);

  let noBtn = CardService.newTextButton()
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setBackgroundColor('#dc3545')
    .setText("No")
    .setOnClickAction(onNoBtnClicked);

  let mainSection = CardService.newCardSection()
    .addWidget(yesBtn)
    .addWidget(noBtn)
    .setHeader('The PTO-892 execution is taking too long<br> Do you want to continue the process?</font>');

  return mainSection;
}

function OAContinueExecution(document) {
  try {
    var pdfFile = JSON.parse(document.parameters.pdfFile);
    let continue892 = JSON.parse(document.parameters.response892);
    var document = document.parameters.document;

    let startFunctionTime = Date.now() - 4800;
    let scriptProperties = PropertiesService.getScriptProperties();

    if (JSON.stringify(document)) {
      let noClicked = scriptProperties.getProperty('timeoutabort');
      let response = readResponse(document, pdfFile, startFunctionTime);

      scriptProperties.deleteProperty('pdfFile');

      let widget, section892, oaFinal = true, final892 = true;
      let parallelExecutionOA = PropertiesService.getScriptProperties().getProperty('parallelExtractionOA');
      let card = CardService.newCardBuilder().addSection(bannerSection).setFixedFooter(CardService.newFixedFooter().setPrimaryButton(homeBtn));

      let message = '<font color="#ff0000">The pdf file is not in correct format!</font>';
      if (noClicked) {
        message = '<font color="#ff0000">OA execution needed more time to get all field values!</font>';
      }
      if (typeof (response) == 'string') {
        if (response.includes('pdfFormatIssue')) {
          widget = CardService.newCardSection().addWidget(
            CardService.newDecoratedText()
              .setTopLabel('Something went wrong with OA...')
              .setText(message)
              .setWrapText(true));
          response = ['printMessage', message];
        }
        else if (response.includes('exceeded_limit')) {
          widget = CardService.newCardSection().addWidget(
            CardService.newDecoratedText()
              .setTopLabel('Something went wrong with OA...')
              .setText('<font color="#ff0000">Exception: Service invoked too many times for one day! You have finished your daily quota.</font><br><br>Please try again in the next day.</font>')
              .setWrapText(true));
          response = ['printMessage', '<font color="#ff0000">Exception: Service invoked too many times for one day! You have finished your daily quota.</font><br><br>Please try again in the next day.</font>'];
        }
        else {
          updateCustomerDetails();
          widget = extractionResultUi(response);
        }
      }
      else {
        if (response[0].includes('pdfFormatIssue')) {
          widget = CardService.newCardSection().addWidget(
            CardService.newDecoratedText()
              .setTopLabel('Something went wrong with OA...')
              .setText(message)
              .setWrapText(true));
          response = ['printMessage', message];
        }
        else if (response[0].includes('exceeded_limit')) {
          widget = CardService.newCardSection().addWidget(
            CardService.newDecoratedText()
              .setTopLabel('Something went wrong with OA...')
              .setText('<font color="#ff0000">Exception: Service invoked too many times for one day! You have finished your daily quota.</font><br><br>Please try again in the next day.</font>')
              .setWrapText(true));
          response = ['printMessage', '<font color="#ff0000">Exception: Service invoked too many times for one day! You have finished your daily quota.</font><br><br>Please try again in the next day.</font>'];
        }
        else if (response[0].includes('timeout')) {
          scriptProperties.setProperty('pdfFile', JSON.stringify(pdfFile));
          widget = OAexecutionTimeExceeded(response[1], continue892);
          oaFinal = false;
        }
        else {
          updateCustomerDetails();
          widget = extractionResultUi(response);
        }
      }


      Logger.log(continue892);
      if (continue892[0] != 'one') {
        if (continue892[0] == "printMessage") {
          section892 = CardService.newCardSection().addWidget(
            CardService.newDecoratedText()
              .setTopLabel('Something went wrong with PTO-892...')
              .setText(continue892[1])
              .setWrapText(true));
        }
        else if (continue892[1][0] == 'exceeded') {
          section892 = executionTimeExceeded(continue892[1], response);
          final892 = false;
        }
        else {
          section892 = extractionResultUi(continue892);
        }
      }



      if (continue892[0] == 'one') {
        card.addSection(widget);
      }
      else {
        if (parallelExecutionOA == 'first') {
          card.addSection(widget).addSection(section892);
        }
        else {
          card.addSection(section892).addSection(widget);
        }
      }

      if (final892 && oaFinal) {
        let footer = footersCardSection();
        card.addSection(footer[0]).addSection(footer[1]);
      }
      return card.build();

    }
    else {
      goToHome();
    }
  } catch ({ message }) {
    Logger.log('OA Continue Execution Exception: ' + message);
    return printMessage('<br><font color="#ff0000">Please try again later, &nbsp;Sorry for the Incovinience!</font>');
  }
}


function OAexecutionTimeExceeded(retrievedDocument, pdfFile, response892) {

  let onYesBtnClicked = CardService.newAction()
    .setFunctionName("OAContinueExecution")
    .setParameters({
      "document": JSON.stringify(retrievedDocument),
      "pdfFile": JSON.stringify(pdfFile),
      "response892": JSON.stringify(response892)
    })
    .setLoadIndicator(CardService.LoadIndicator.SPINNER);

  let yesBtn = CardService.newTextButton()
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setBackgroundColor('#28a745')
    .setText("Yes")
    .setOnClickAction(onYesBtnClicked);

  let onNoBtnClicked = CardService.newAction()
    .setFunctionName("OAcontinueWithDownloaded")
    .setParameters({
      "document": JSON.stringify(retrievedDocument),
      "pdfFile": JSON.stringify(pdfFile),
      "response892": JSON.stringify(response892)
    })
    .setLoadIndicator(CardService.LoadIndicator.SPINNER);

  let noBtn = CardService.newTextButton()
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setBackgroundColor('#dc3545')
    .setText("No")
    .setOnClickAction(onNoBtnClicked);

  let mainSection = CardService.newCardSection()
    .addWidget(yesBtn)
    .addWidget(noBtn)
    .setHeader('The OA execution is taking too long<br> Do you want to continue the process?</font>');

  return mainSection;
}

function OAcontinueWithDownloaded(parameter) {
  let scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('timeoutabort', 'true');
  return OAContinueExecution(parameter);
}

function extractionResultUi(documentId) {
  try {
    var labelText, folder;
    if (typeof (documentId) == 'string' || documentId[0] != "892") {

      var doc = '';
      if (typeof (documentId) == 'string') {
        doc = DriveApp.getFileById(documentId);
      }
      else {
        doc = DriveApp.getFileById(documentId.pop());
      }
      var folders = doc.getParents();
      if (folders.hasNext()) {
        folder = folders.next();
      }
      labelText = 'Click below to open the first draft of the Office Action Response Document';
    }
    else {
      documentId.shift();
      labelText = 'Click below to see all prior art patents cited by the examiner in Form PTO-892';
      folder = DriveApp.getFolderById(documentId.shift());
    }
    let folderName = folder.getName();
    let folderUrl = folder.getUrl();

    let mainSectionLabelText = CardService.newDecoratedText()
      .setText('<font color="#566573">' + labelText + '</font>')
      .setWrapText(true)

    let linkWidget = CardService
      .newKeyValue()
      .setIconUrl('https://icon-library.com/images/finished-icon/finished-icon-2.jpg')
      .setContent(HtmlService.createHtmlOutput('<font color="#187bcd"><a>' + folderName + '</a></font>').getContent())
      .setOpenLink(CardService.newOpenLink()
        .setUrl(folderUrl));


    let mainSection = CardService.newCardSection()
      .addWidget(mainSectionLabelText)
      .addWidget(linkWidget);

    return mainSection;
  } catch ({ message }) {
    Logger.log('Extraction Result UI Exception: ' + message);
    return printMessage('<br><font color="#ff0000">Please try again later, &nbsp;Sorry for the Incovinience!</font>');
  }
}

function goToHome() {
  var folderList = DriveApp.getFoldersByName(folderName);

  if (folderList.hasNext()) {
    var folder = folderList.next();
    folder.setTrashed(true);
  }
  var scriptProperties = PropertiesService.getScriptProperties();
  if (scriptProperties.getProperty('addFile')) {
    scriptProperties.deleteProperty('addFile');
  }


  var card = selectOption();
  var nav = CardService.newNavigation()
    .popToRoot().updateCard(card);

  return CardService.newActionResponseBuilder()
    .setNavigation(nav)
    .build();
}

function footersCardSection() {
  var usage = returnUsageStatus();
  try {
    if (typeof (usage[0]) == 'number') {

      let remainingCount = (usage[1] + FreeQuota) - usage[0];

      let usageSectionDecoratedTextProcessed = CardService.newDecoratedText()
        .setText('OAs and PTO-892s Processed:  <font color="#CD5C5C">' + usage[0] + '</font>')
        .setWrapText(true);

      let usageSectionDecoratedTextRemaining = CardService.newDecoratedText()
        .setText('Remaining OAs and PTO-892s quota:  <font color="#CD5C5C">' + remainingCount + '</font>')
        .setWrapText(true);

      let usageSection = CardService.newCardSection()
        .addWidget(usageSectionDecoratedTextProcessed)
        .addWidget(usageSectionDecoratedTextRemaining)
        .setHeader('Usage Status');

      let contactUsWidget = CardService
        .newKeyValue()
        .setIconUrl('https://www.gstatic.com/images/icons/material/system/1x/mail_googblue_48dp.png')
        .setContent(HtmlService.createHtmlOutput('<a href="mailto:support@hexon-ip.com">Contact us</a>').getContent());

      let feedbackWidget = CardService
        .newKeyValue()
        .setIconUrl('https://www.gstatic.com/images/icons/material/system/1x/feedback_googblue_48dp.png')
        .setContent(HtmlService.createHtmlOutput('<a href="https://forms.gle/egoxrkezB1n1RBBJ8"><i>Provide your Feedback here</a>').getContent());

      let refferalWidget = CardService
        .newKeyValue()
        .setIconUrl('https://icon-library.com/images/refer-a-friend-icon-png/refer-a-friend-icon-png-22.jpg')
        .setContent(HtmlService.createHtmlOutput('<a href="https://forms.gle/PUAgac5Gu8o3jErq5">Let your friends know about this add-on</a>').getContent())
        .setMultiline(true);

      let companyName = CardService.newKeyValue().setContent(HtmlService.createHtmlOutput('<a href="https://www.hexon-ip.com">Hexon-ip Infusing Technology into IP</a>').getContent()).setIconUrl('https://icon-library.com/images/icon-copyright/icon-copyright-4.jpg');

      let feedbackSection = CardService.newCardSection()
        .addWidget(contactUsWidget)
        .addWidget(feedbackWidget)
        .addWidget(refferalWidget)
        .addWidget(CardService.newDivider())
        .addWidget(companyName)
        .setHeader('For further information');

      return [usageSection, feedbackSection];
    }
    else {
      return usage;
    }
  }
  catch ({ message }) {
    Logger.log('Footers Card Section Exception:' + message);
  }
}

function negate(option) {
  if (option == 'oa') {
    return '892';
  }
  else if (option == '892') {
    return 'oa';
  }
}
