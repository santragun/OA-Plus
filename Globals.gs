var configSpreadsheetId = '1_aIs2imj8lTYAptNRO_YV_9uANzqY2RvLrTMhaShsL0';

var customerInformationSheetName = 'Customer Information';
var mailSheetName = 'Mail Content';
var folderName = ".OA+ config";
var previousResponse = '';

var WebAppUrl = "https://script.google.com/macros/s/AKfycbzoVoBsCvVyHcDZ7NM7b3u1eN9A9FyzrVoBy5UfDqgWRoDUolBDUhyQbRBQ2p1-GNlp/exec";
var APIExecutabledeploymentId = 'AKfycbxjGLgkmyHt37JyS1gqDvk3iuSizuQYXnSmOk4SmJg1ImolreH3e8VWroHWioqtXnm-4g';
var FreeQuota = 100;

try{
  var additionalValues = callSpreadsheetDatabaseAPI('GetAdditionalValues', '')[0];
  FreeQuota = additionalValues[0];
  WebAppUrl = additionalValues[1];
}
catch({message}){
  Logger.log('Call API GetAdditionalValues Exception: ' + message);
}

var fontStyle = {
  [DocumentApp.Attribute.FONT_FAMILY]: 'Arial',
  [DocumentApp.Attribute.FONT_SIZE]: 12,
};
var headingStyle = {
  [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.CENTER,
  [DocumentApp.Attribute.BOLD]: true,
  [DocumentApp.Attribute.FOREGROUND_COLOR]: '#000000',
  [DocumentApp.Attribute.FONT_SIZE]: 14
};
var noticeStyle = {
  [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.CENTER
};

var banner = CardService.newImage()
  .setImageUrl('https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcT0EuY4DEj4Mf_ASg7yTNyKPx0nrO_36f01EMl_zryq22jrejBJHY4_9rSfqy5IhjnVo7M&usqp=CAU');

var bannerSection = CardService.newCardSection()
  .addWidget(banner);

var homeAction = CardService.newAction()
  .setFunctionName('goToHome')
  .setLoadIndicator(CardService.LoadIndicator.SPINNER);

var homeBtn = CardService.newTextButton()
  .setText("Home")
  .setOnClickAction(homeAction);

const ONE_SECOND = 1000;
const ONE_MINUTE = ONE_SECOND * 60;
const MAX_EXECUTION_TIME = ONE_SECOND * 30;

const isTimeLeft = (functionStartedTime) => {
  return MAX_EXECUTION_TIME > Date.now() - functionStartedTime;
};

var emailAddress = Session.getEffectiveUser().getEmail();

function checkIfKeyword(array, string) {
  return array.some(function (element) {
    return string.includes(element);
  });
}

var parallelExecutionCard = CardService.newCardBuilder();


function callSpreadsheetDatabaseAPI(functionName, parameters) {
  try{
  var scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/script.send_mail'];
  var service = getOAuthServiceClientId(scopes);
  // service.reset();
  if (service.hasAccess()) {
    var token = service.getAccessToken();

    var url = 'https://script.googleapis.com/v1/scripts/' + APIExecutabledeploymentId + ':run';
    var options = {
      method: "POST",
      headers: { Authorization: "Bearer " + token },
      contentType: "application/json",
      payload: JSON.stringify({
        "function": functionName,
        "parameters": parameters,
        "devMode": false
      }),
      muteHttpExceptions: true
    };
    var response = UrlFetchApp.fetch(url, options);
    if (JSON.parse(response).response != undefined) {
      return JSON.parse(response).response.result;
    }
    else {

      Logger.log('Call API not working ' + response);
      if (JSON.parse(response).error && JSON.parse(response).error.code == '403' && JSON.parse(response).error.message.includes('Request had insufficient authentication scopes')) {
        var authorizationUrl = service.getAuthorizationUrl();
        return ['Not authorized' + authorizationUrl];
      }
      return '';
    }

  }
  else {
    var authorizationUrl = service.getAuthorizationUrl();
    return ['Not authorized' + authorizationUrl];
  }
  }catch({message}){
    Logger.log('Call API Exception: ' + message);
  }
}






