function checkAuthrization() {
  let isAuthorized = getUsageStatus();
  return isAuthorized;
}

function returnUsageStatus() {
  try {
    var customerInformationSheetData = callSpreadsheetDatabaseAPI('GetCustomerInformationData', '')[0];
    if (customerInformationSheetData) {
      if (!customerInformationSheetData.includes('Not authorized')) {
        var data = customerInformationSheetData;
        let row = 0, processed, payedFor;
        for (var i = 1; i < data.length; i++) {
          if (data[i][1] == emailAddress) {
            row = i;
            break;
          }
        }
        if (row == 0) {
          callSpreadsheetDatabaseAPI('InsertToCustomerInformationData', JSON.stringify([emailAddress, 0, 0]));
          processed = 0;
          payedFor = 0;
        }
        else {
          processed = data[i][2];
          payedFor = data[i][3];
        }
        return [processed, payedFor];

      }
      else {
        return ['failed ' + customerInformationSheetData.substring(14)];
      }
    }
    else {
      return ['failed'];
    }
  }
  catch ({ message }) {
    Logger.log('Return Usage Status Exception: ' + message);
  }
}

function getUsageStatus() {
  try{
  var data = callSpreadsheetDatabaseAPI('GetCustomerInformationData', '')[0];
  if (data) {
    let row = 0;
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] == emailAddress) {
        row = i;
        break;
      }
    }
    if (row == 0) {
      callSpreadsheetDatabaseAPI('InsertToCustomerInformationData', JSON.stringify([emailAddress, 0, 'non-payed']));
      return true;
    }
    else {

      var message = callSpreadsheetDatabaseAPI('GetMailSheetData', '')[0][1][1];
      let processed = data[i][2];
      let payedFor = data[i][3];
      if (!isNaN(processed) && !isNaN(payedFor) && processed >= payedFor + FreeQuota) {
        sendPaymentDetails(message);
        return false;
      }
    }
    return true;
  }
  else {
    return false;
  }
}
  catch ({ message }) {
    Logger.log('Get Usage Status Exception: ' + message);
  }
}

function sendPaymentDetails(message) {
  callSpreadsheetDatabaseAPI('SendMail', [emailAddress, 'OA+ Subscription', message]);
}

function updateCustomerDetails() {
  try{
  var data = callSpreadsheetDatabaseAPI('GetCustomerInformationData', '')[0];
  if (data) {
    let row = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] == emailAddress) {
        row = i;
        break;
      }
    }
    let usage = data[row][2];
    usage += 1;
    callSpreadsheetDatabaseAPI('UpdateCustomerInformationData', [row + 1, 3, usage]);
    let payedFor = data[row][3];

    let message = callSpreadsheetDatabaseAPI('GetMailSheetData', '')[0][1][0];
    if (usage == (payedFor + FreeQuota) - 1) {
      sendPaymentDetails(message);
    }
    else if (usage == (payedFor + FreeQuota)) {
      let message = callSpreadsheetDatabaseAPI('GetMailSheetData', '')[0][1][1];
      sendPaymentDetails(message);
    }
  }
  }
  catch ({ message }) {
    Logger.log('Update Usage Status Exception: ' + message);
  }
}
