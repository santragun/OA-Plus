function getFirstPageFields(blocks, document, functionStartedTime) {
  functionStartedTime -= 25000;
  var formFields = {};
  var scriptProperties = PropertiesService.getScriptProperties();
  let timeoutProperty = scriptProperties.getProperty('timeout');
  let timeoutabort = scriptProperties.getProperty('timeoutabort');

  if (timeoutProperty) {
    scriptProperties.deleteProperty('timeout');
    if (scriptProperties.getProperty('formFields')) {
      formFields = JSON.parse(scriptProperties.getProperty('formFields'));
    }
      if (timeoutabort) {
        defaultValues(formFields);
        return formFields;
      }
      var fullTextSegments = JSON.parse(scriptProperties.getProperty('fullTextSegments'));
      if (fullTextSegments) {
        try {
          getCommonFields(fullTextSegments.join('\n'), formFields);
          scriptProperties.deleteProperty('fullTextSegments');
        } catch ({ exception }) {
          Logger.log('getFirstPageFields Exception: timeout ' + exception);
        }
      }
      else {
        try {
          for (let i = 0; i < blocks.length; i++) {
            fullTextSegments.push(extractTextSegments(blocks[i].layout.textAnchor.textSegments, document));
          }

          for (let i = 0; i < fullTextSegments && isTimeLeft(functionStartedTime); i++) {
            getCommonFields(fullTextSegments[i], formFields);
          }
          if (!isTimeLeft(functionStartedTime)) {
            Logger.log('tt imeout');
            scriptProperties.setProperty('timeout', 'yes');
            scriptProperties.setProperty('formFields', JSON.stringify(formFields));
            scriptProperties.setProperty('fullTextSegments', JSON.stringify(fullTextSegments));
            return '';
          }
        } catch ({ exception }) {
          Logger.log('getFirstPageFields Exception timeout empty: ' + exception);
        }
      }

      scriptProperties.deleteProperty('formFields');
      scriptProperties.deleteProperty('timeout');
    
  }
  else {
    try {
      var fullTextSegments = [];
      for (let i = 0; i < blocks.length; i++) {
        fullTextSegments.push(extractTextSegments(blocks[i].layout.textAnchor.textSegments, document));
      }
      fullTextSegments = fullTextSegments.filter(element => element);
      if (fullTextSegments) {
          getCommonFields(fullTextSegments.join('\n'), formFields);
        }
      if (!formFields['applicationnumber'] || !formFields['firstnamedapplicant'] || !formFields['filingdate']) {
      for (let i = 0; i < fullTextSegments && isTimeLeft(functionStartedTime); i++) {
        getCommonFields(fullTextSegments[i], formFields);
      }
      }
      if (!isTimeLeft(functionStartedTime)) {
        Logger.log('tt imeout');
        scriptProperties.setProperty('timeout', 'yes');
        scriptProperties.setProperty('formFields', JSON.stringify(formFields));
        scriptProperties.setProperty('fullTextSegments', JSON.stringify(fullTextSegments));
        return '';
      }
      if (!formFields['applicationnumber'] || !formFields['firstnamedapplicant'] || !formFields['filingdate']) {
        if (fullTextSegments) {
          getCommonFields(fullTextSegments.join('\n'), formFields);
        }
      }
      if (!isTimeLeft(functionStartedTime)) {
        scriptProperties.setProperty('timeout', 'yes');
        scriptProperties.setProperty('formFields', JSON.stringify(formFields));
        scriptProperties.setProperty('fullTextSegments', JSON.stringify(fullTextSegments));
        return '';
      }
    } catch ({ exception }) {
      Logger.log('getFirstPageFields Exception: init' + exception);
    }
  }
  defaultValues(formFields);
  Logger.log(formFields);
  return formFields;
}

function getCommonFields(blockText, formFields) {
  let blockSplit = blockText.toLowerCase().split('\n').filter(item => item);
  var blockArray = blockText.split('\n').filter(item => item);
  Logger.log(blockArray);

  const keywords = ['attorney', 'application', 'filing', 'first name', 'replaced', 'atty. docket', 'confirmation',
    'art unit', 'paper', 'united states patent and trademark office', 'notification date', 'ptol-90a (rev',
    'pto-90c (', 'examiner', 'mail date', 'notification', 'united states department of commerce',
    'u.s. patent and trademark office', 'address', 'commissioner for'];

  for (let i = 0; i < blockSplit.length; i++) {

    // ADDRESS BLOCK
    if (!formFields.hasOwnProperty('address') && blockSplit[i].includes("address:")) {
      formFields['address'] = blockArray[i].substring(blockSplit[i].indexOf('address:') + 9) + '\n';

      let start = i;
      for (let j = 1; j <= 3; j++) {
        start += 1;
        if (blockSplit[start]) {
          while (checkIfKeyword(keywords, blockSplit[start])) {
            Logger.log('str  9 ' + blockSplit[start]);
            start++;
          }
        }
        else if (blockSplit[start + 1]) {
          start += 1;
          while (checkIfKeyword(keywords, blockSplit[start])) {
            Logger.log('str  9' + blockSplit[start + 1]);
            start += 2;
          }
        }
        formFields['address'] += blockArray[start] + '\n';
      }

      if (formFields['address'] == blockArray[i].substring(blockSplit[i].indexOf('address:') + 9) + '\n') {
        delete formFields['address'];
      }
    }

    // FIRST NAMED APPLICANT
    var key = 'firstnamedapplicant';
    var keyword = 'first named applicant';
    var keyword1 = 'first named inventor';

    if (!formFields.hasOwnProperty(key) && (blockSplit[i].includes(keyword) || blockSplit[i].includes(keyword1))) {
      let counter = i + 1;
      while (counter < blockSplit.length && (!blockArray[counter].match(/^[a-zA-ZàáâäãåąčćęèéêëėįìíîïłńòóôöõøùúûüųūÿýżźñçčšžÀÁÂÄÃÅĄĆČĖĘÈÉÊËÌÍÎÏĮŁŃÒÓÔÖÕØÙÚÛÜŲŪŸÝŻŹÑßÇŒÆČŠŽ∂ð ,.'-]+$/) ||
        checkIfKeyword(keywords, blockArray[counter].toLowerCase()))) {
        counter++;
      }
      if (blockArray[counter]) {
        formFields[key] = blockArray[counter];
      }
    }

    // APPLICATION NUMBER 
    key = 'applicationnumber';
    let key2 = 'confirmationno';
    keyword = 'application number';
    keyword1 = 'application no';

    if (!formFields.hasOwnProperty(key) && (blockSplit[i].includes(keyword) || blockSplit[i].includes(keyword1))) {
      let counter = i;
      Logger.log('app ' + blockArray[counter]);
      while (counter < blockSplit.length && !blockArray[counter].match(/[\d]{2}\/[\d]{3},[\d]{3}/)) {
        counter++;
      }
      if (blockArray[counter]) {
        formFields[key] = blockArray[counter];
        blockArray.splice(counter, 1, 'REPLACED');
      }
      Logger.log('app ' + blockArray[counter]);
    }

    // CONFIRMATION NUMBER
    key = 'confirmationno';
    keyword = 'confirmation no';

    if (!formFields.hasOwnProperty(key) && blockSplit[i].includes(keyword)) {
      let counter = i;
      if (blockSplit[i].match(/\d+$/) && blockArray[counter].match(/[\d]{1,2}\/[\d]{1,2}\/[\d]{2,4}|[\d]{1,2}\-[\d]{1,2}\-[\d]{2,4}/)) {
        formFields[key] = blockSplit[i].match(/[\d]{4}$/);
      }
      else {
        while (counter < blockSplit.length && (!blockArray[counter].match(/[\d]{4}$/) || blockArray[counter].match(/[\d]{1,2}\/[\d]{1,2}\/[\d]{2,4}|[\d]{1,2}\-[\d]{1,2}\-[\d]{2,4}/))) {
          counter++;
        }
        if (blockArray[counter]) {
          formFields[key] = blockArray[counter].match(/[\d]{4}$/);
          blockArray.splice(i, 1, 'REPLACED');
          blockArray.splice(counter, 1, 'REPLACED');
        }
      }
    }

    // FILING DATE
    key = 'filingdate';
    keyword = 'filing date';
    keyword1 = 'filing or 371(c) date';
    var keyword2 = 'filing or 371 (c) date';
    var keyword3 = 'filing date:';

    if (!formFields.hasOwnProperty(key) && (blockSplit[i] == keyword || blockSplit[i].includes(keyword1) || blockSplit[i].includes(keyword2) || blockSplit[i].includes(keyword3))) {
      let counter = i;
      while (counter < blockSplit.length && !blockArray[counter].match(/[\d]{1,2}\/[\d]{1,2}\/[\d]{2,4}|[\d]{1,2}\-[\d]{1,2}\-[\d]{2,4}/)) {
        counter++;
      }
      if (blockSplit[counter - 1].includes('confirmation')) {
        if (!formFields.hasOwnProperty(key2)) {
          if (blockArray[counter]) {
            formFields[key2] = blockArray[counter];
            blockArray.splice(counter - 1, 1, 'REPLACED');
            blockArray.splice(counter, 1, 'REPLACED');
          }
        }
        counter = i;
        while (counter < blockSplit.length && !blockArray[counter].match(/[\d]{1,2}\/[\d]{1,2}\/[\d]{2,4}|[\d]{1,2}\-[\d]{1,2}\-[\d]{2,4}/)) {
          counter++;
        }
      }
      if (counter < blockSplit.length && blockArray[counter].match(/[\d]{1,2}\/[\d]{1,2}\/[\d]{2,4}|[\d]{1,2}\-[\d]{1,2}\-[\d]{2,4}/)) {
        formFields['filingdate'] = blockArray[counter];
        blockArray.splice(counter, 1, 'REPLACED');
      }
    }

    // MAIL DATE
    key = 'datemailed';
    var key3 = 'noticename';
    keyword = 'date mailed';
    keyword1 = 'mail date';
    keyword2 = 'notification date';

    if (!formFields.hasOwnProperty(key) && (blockSplit[i].includes(keyword) || blockSplit[i].includes(keyword1) || blockSplit[i].includes(keyword2))) {
      let counter = i;
      if (blockSplit[i].match(/[\d]{1,2}\/[\d]{1,2}\/[\d]{2,4}|[\d]{1,2}\-[\d]{1,2}\-[\d]{2,4}/)) {
        formFields[key] = blockSplit[i].match(/[\d]{1,2}\/[\d]{1,2}\/[\d]{2,4}|[\d]{1,2}\-[\d]{1,2}\-[\d]{2,4}/); // date[date.length - 1];
      }
      else {
        while (counter < blockSplit.length && !blockArray[counter].match(/[\d]{1,2}\/[\d]{1,2}\/[\d]{2,4}|[\d]{1,2}\-[\d]{1,2}\-[\d]{2,4}/)) {
          counter++;
        }
        if (blockArray[counter]) {
          formFields[key] = blockArray[counter];
          blockArray.splice(counter, 1, 'REPLACED');
        }
      }
      if (!formFields.hasOwnProperty(key3)) {
        counter += 1;
        while (counter < blockSplit.length && !blockSplit[counter].startsWith('notice')) {
          counter++;
        }
        if (blockArray[counter]) {
          formFields[key3] = blockArray[counter];
        }
      }
      if (blockSplit[i].includes(keyword2) && !formFields.hasOwnProperty(key)) {
        counter = i;
        while (counter > 0 && !blockArray[counter].match(/[\d]{1,2}\/[\d]{1,2}\/[\d]{2,4}|[\d]{1,2}\-[\d]{1,2}\-[\d]{2,4}/)) {
          counter--;
        }
        if (blockArray[counter] && blockArray[counter].match(/[\d]{1,2}\/[\d]{1,2}\/[\d]{2,4}|[\d]{1,2}\-[\d]{1,2}\-[\d]{2,4}/)) {
          formFields[key] = blockArray[counter];
          blockArray.splice(counter, 1, 'REPLACED');
        }
      }
    }

    // NOTICE NAME
    if (!formFields.hasOwnProperty(key3) && blockSplit[i].startsWith("notice")) {
      formFields[key3] = blockArray[i];
    }

    // ATTORNEY DOCKET NUMBER
    if (!formFields.hasOwnProperty('attorneydocket') && (blockSplit[i].includes("attorney docket") || blockSplit[i].includes("atty. docket no"))) {
      let counter = i;
      while (counter < blockSplit.length && !/\d/.test(blockArray[counter])) {
        counter++;
      }
      if (blockSplit[counter]) {
        if (checkIfKeyword(keywords, blockSplit[counter]) || /^\d{4,5}$/.test(blockArray[counter])) {
          counter += 1;
          while (counter < blockSplit.length && !/\d/.test(blockArray[counter])) {
            counter++;
          }
        }
        if (blockSplit[counter]) {
          formFields['attorneydocket'] = blockSplit[counter];
          blockArray.splice(counter, 1, 'REPLACED');
        }
      }
    }

    // EXAMINER
    if (!formFields.hasOwnProperty('examiner') && blockSplit[i].includes("examiner")) {
      let counter = i + 1;
      while (counter < blockSplit.length && (checkIfKeyword(keywords, blockArray[counter].toLowerCase()) ||
        !blockArray[counter].match(/^[a-zA-ZàáâäãåąčćęèéêëėįìíîïłńòóôöõøùúûüųūÿýżźñçčšžÀÁÂÄÃÅĄẠĆČĖĘÈÉÊËÌÍÎÏĮŁŃÒÓÔÖÕØÙÚÛÜŲŪŸÝŻŹÑßÇŒÆČŠŽ∂ð ,.'-]+$/)
        || blockArray[counter].startsWith('www.'))) {
        counter++;
      }
      if (blockArray[counter] && !blockArray[counter].match(/\d/) &&
        !blockArray[counter].includes('DELIVERY MODE')) {
        formFields['examiner'] = blockArray[counter];
        Logger.log(checkIfKeyword(keywords, blockArray[counter].toLowerCase()) + ' ddress');
      }
      else {
        counter -= 1;
        while (counter > 0 && (checkIfKeyword(keywords, blockArray[counter].toLowerCase()) ||
          !blockArray[counter].match(/^[a-zA-ZàáâäãåąčćęèéêëėįìíîïłńòóôöõøùúûüųūÿýżźñçčšžÀÁÂÄÃÅĄẠĆČĖĘÈÉÊËÌÍÎÏĮŁŃÒÓÔÖÕØÙÚÛÜŲŪŸÝŻŹÑßÇŒÆČŠŽ∂ð ,.'-]+$/)
          || blockArray[counter].startsWith('www.'))) {
          counter--;
        }
        if (blockArray[counter] && !blockArray[counter].match(/\d/) && !blockArray[counter].includes('DOCKET')
          && !blockArray[counter].toLowerCase().includes('date')) {
          if (formFields['firstnamedapplicant'] && formFields['firstnamedapplicant'] != blockArray[counter]) {
            formFields['examiner'] = blockArray[counter];
          }
        }
      }
    }

    // GROUP ART UNIT
    if (!formFields.hasOwnProperty('groupartunit') && blockSplit[i].includes("art unit")) {
      let counter = i;
      while (counter < blockSplit.length && !blockArray[counter].match(/^\d/)) {
        counter++;
      }
      if (blockArray[counter]) {
        formFields['groupartunit'] = blockArray[counter];
        blockArray.splice(counter, 1, 'REPLACED');
      }
    }

    // CUSTOMER ADDRESS

    if (!formFields.hasOwnProperty('customerno') && blockArray[i + 3] != undefined) {
      var lastIndex;
      if (blockArray[i + 5] != undefined && (/^[\d]{5,6}/.test(blockArray[i]) || /^[\d]{3}/.test(blockArray[i])) &&
        /^\d/.test(blockArray[i + 1]) &&
        /^[a-zA-Z0-9&.,\/\-():'\s]*$/.test(blockArray[i + 2]) && blockArray[i + 2] != 'e' &&
        /^[a-zA-Z0-9&.,\/\-():'\s]*$/.test(blockArray[i + 3]) &&
        !/^\d+$/.test(blockArray[i + 3]) &&
        /^[a-zA-Z0-9&.,\/\-():'\s]*$/.test(blockArray[i + 4]) &&
        !/^\d+$/.test(blockArray[i + 4])) {

        formFields['customerno'] = blockArray[i];
        formFields['customeraddress'] = '\n' + blockArray[i + 3];

        if (!blockArray[i + 2].match(/[\d]{1,2}\/[\d]{1,2}\/[\d]{2,4}|[\d]{1,2}\-[\d]{1,2}\-[\d]{2,4}/)) {
          formFields['customeraddress'] = '\n' + blockArray[i + 2] + formFields['customeraddress'];
        }
        if (!checkIfKeyword(keywords, blockArray[i + 4].toLowerCase())) {
          formFields['customeraddress'] += '\n' + blockArray[i + 4];
        }
        Logger.log('second if ' + blockArray[i]);
        lastIndex = i + 5;
        if (checkIfKeyword(keywords, blockArray[i + 2].toLowerCase())) {
          let counter = i + 2;

          while (counter < blockSplit.length && (checkIfKeyword(keywords, blockArray[counter].toLowerCase()) ||
            blockArray[counter].match(/[\d]{1,2}\/[\d]{1,2}\/[\d]{2,4}|[\d]{1,2}\-[\d]{1,2}\-[\d]{2,4}/))) {
            counter++;
          }
          if (blockArray[counter].match(/^\d+$/)) {
            counter += 1;
          }
          formFields['customeraddress'] = '\n' + blockArray[counter] + '\n' + blockArray[counter + 1] +
            '\n' + blockArray[counter + 2];
          lastIndex = counter + 3;
        }
      }
      else if (blockArray[i + 6] != undefined && (/^[\d]{5,6}/.test(blockArray[i]) || /^[\d]{3}/.test(blockArray[i])) &&
        /^\d/.test(blockArray[i + 1]) &&
        /^[a-zA-Z0-9&.,\/\-():'\s]*$/.test(blockArray[i + 3]) &&
        !/^\d+$/.test(blockArray[i + 3]) &&
        /^[a-zA-Z0-9&.,\/\-():'\s]*$/.test(blockArray[i + 4]) &&
        !/^\d+$/.test(blockArray[i + 4]) &&
        /^[a-zA-Z0-9&.,\/\-():'\s]*$/.test(blockArray[i + 5]) &&
        !/^\d+$/.test(blockArray[i + 5])) {
        if (checkIfKeyword(keywords, blockArray[i + 4].toLowerCase())) {
          blockArray[i + 4] = blockArray[i + 3];
          blockArray[i + 3] = blockArray[i + 2];
        }
        Logger.log('first if ' + blockArray[i]);

        formFields['customerno'] = blockArray[i];
        formFields['customeraddress'] = '\n' + blockArray[i + 3] + '\n' + blockArray[i + 4] + '\n' + blockArray[i + 5];
        lastIndex = i + 6;

        if (checkIfKeyword(keywords, blockArray[i + 3].toLowerCase())) {
          let counter = i + 1;

          while (counter < blockSplit.length && (checkIfKeyword(keywords, blockArray[counter].toLowerCase()) ||
            blockArray[counter].match(/[\d]{1,2}\/[\d]{1,2}\/[\d]{2,4}|[\d]{1,2}\-[\d]{1,2}\-[\d]{2,4}/))) {
            counter++;
          }
          if (blockArray[counter].match(/^\d+$/)) {
            counter += 1;
          }
          formFields['customeraddress'] = '\n' + blockArray[counter] + '\n' + blockArray[counter + 1] +
            '\n' + blockArray[counter + 2] + '\n' + blockArray[counter + 3];
          lastIndex = counter + 4;
        }
      }
      else if ((/^[\d]{5,6}/.test(blockArray[i]) || /^[\d]{3}/.test(blockArray[i])) &&
        /^[a-zA-Z0-9.,&\/\-():'\s]*$/.test(blockArray[i + 1]) &&
        !/^\d+$/.test(blockArray[i + 1]) &&
        /^[a-zA-Z0-9&.,\/\-():'\s]*$/.test(blockArray[i + 2]) &&
        !/^\d+$/.test(blockArray[i + 2]) &&
        /^[a-zA-Z0-9&.,\/\-():'\s]*$/.test(blockArray[i + 3]) &&
        !/^\d+$/.test(blockArray[i + 3])) {


        formFields['customerno'] = blockArray[i];
        formFields['customeraddress'] = '\n' + blockArray[i + 2] + '\n' + blockArray[i + 3];
        if (blockArray[i + 1] != 'e' && !checkIfKeyword(keywords, blockArray[i + 1].toLowerCase())) {
          formFields['customeraddress'] = '\n' + blockArray[i + 1] + formFields['customeraddress'];
        }
        Logger.log('third if ' + blockArray[i]);

        lastIndex = i + 4;

        if (checkIfKeyword(keywords, blockArray[i + 1].toLowerCase())) {
          let counter = i + 1;

          while (counter < blockSplit.length && (checkIfKeyword(keywords, blockArray[counter].toLowerCase()) ||
            blockArray[counter].match(/[\d]{1,2}\/[\d]{1,2}\/[\d]{2,4}|[\d]{1,2}\-[\d]{1,2}\-[\d]{2,4}/))) {
            counter++;
          }
          if (blockArray[counter].match(/^\d+$/) || blockArray[counter] == 'e') {
            counter += 1;
            if (blockArray[counter].match(/^\d+$/) || blockArray[counter] == 'e' ||
              blockArray[counter].match(/[\d]{1,2}\/[\d]{1,2}\/[\d]{2,4}|[\d]{1,2}\-[\d]{1,2}\-[\d]{2,4}/)) {
              counter += 1;
            }
          }
          formFields['customeraddress'] = '\n' + blockArray[counter] + '\n' + blockArray[counter + 1] +
            '\n' + blockArray[counter + 2];
          if (!checkIfKeyword(keywords, blockArray[counter + 3].toLowerCase()) &&
            !blockArray[counter + 3].match(/[\d]{1,2}\/[\d]{1,2}\/[\d]{2,4}|[\d]{1,2}\-[\d]{1,2}\-[\d]{2,4}/)) {
            formFields['customeraddress'] += '\n' + blockArray[counter + 3];
          }

          lastIndex = counter + 4;
        }
      }
      var cities = ['los angeles', 'alabama', 'alaska', 'arizona', 'arkansas', 'california', 'chicago', 'colorado',
        'connecticut', 'delaware', 'florida', 'georgia', 'hawaii', 'idaho', 'illinois', 'indiana', 'iowa',
        'kansas', 'kentucky', 'louisiana', 'maine', 'maryland', 'massachusetts', 'michigan', 'minnesota',
        'mississippi', 'missouri', 'montana', 'nebraska', 'evada', 'new hampshire', 'new jersey', 'new mexico',
        'new york', 'north carolina', 'north dakota', 'ohio', 'oklahoma', 'oregon', 'pennsylvania',
        'rhode island', 'south carolina', 'south Dakota', 'tennessee', 'texas', 'utah', 'vermont',
        'virginia', 'washington', 'west virginia', 'wisconsin', 'wyoming', 'washington', 'seattle', 'palo alto',
        'san jose', 'troy', 'reston', 'bedminster', 'philadelphia', 'houston', 'louisville', 'san diego',
        'west orange', 'monroe', 'las vegas'];

      if (lastIndex) {
        if (/^[a-zA-Z0-9&.,\/\-():'\s]*$/.test(blockArray[lastIndex]) &&
          !/^\d+$/.test(blockArray[lastIndex]) && !checkIfKeyword(keywords, blockArray[lastIndex].toLowerCase()) &&
          !blockArray[lastIndex].match(/[\d]{1,2}\/[\d]{1,2}\/[\d]{2,4}|[\d]{1,2}\-[\d]{1,2}\-[\d]{2,4}/)
          && !blockArray[lastIndex].match(/^[a-zA-ZàáâäãåąčćęèéêëėįìíîïłńòóôöõøùúûüųūÿýżźñçčšžÀÁÂÄÃÅĄẠĆČĖĘÈÉÊËÌÍÎÏĮŁŃÒÓÔÖÕØÙÚÛÜŲŪŸÝŻŹÑßÇŒÆČŠŽ∂ð ,.'-]+$/)) {
          formFields['customeraddress'] += '\n' + blockArray[lastIndex];
        }

        if (blockArray[lastIndex + 1] != undefined && cities.some(city => blockArray[lastIndex + 1].toLowerCase().includes(city))) {
          formFields['customeraddress'] += '\n' + blockArray[lastIndex + 1];
        }
        if (blockArray[lastIndex + 2] != undefined && cities.some(city => blockArray[lastIndex + 2].toLowerCase().includes(city))) {
          formFields['customeraddress'] += '\n' + blockArray[lastIndex + 2];
        }
      }
    }
  }
  return formFields;
}



function defaultValues(formFields) {

  //GETTING DEFAULT VALUES
  var defaultValues = getDefaultValues();
  if (defaultValues != null) {
    formFields['phone'] = defaultValues.phone;
    formFields['fax'] = defaultValues.fax;
    formFields['email'] = defaultValues.email;
    formFields['attorneyname'] = defaultValues.attorneyName;
    formFields['regno'] = defaultValues.regNo;
    if (!formFields.hasOwnProperty('groupartunit')) {
      formFields['groupartunit'] = defaultValues.groupArtUnit;
    }
    formFields['fore'] = defaultValues['for'];
    if (!formFields.hasOwnProperty('examiner')) {
      formFields['examiner'] = defaultValues.examiner;
    }
  }
  var keys = ['phone', 'fax', 'email', 'attorneyname', 'regno', 'groupartunit', 'fore', 'examiner'];
  for (let i = 0; i < keys.length; i++) {
    let key = keys[i];
    if (formFields[key] == null || formFields[key] == undefined) {
      formFields[key] = '';
    }
  }

  if (formFields['fore'] == '') {
    formFields['fore'] = 'Application Title ';
  }
  if (formFields['examiner'] == '') {
    formFields['examiner'] = 'Not Yet Assigned ';
  }

  return formFields;
}



function getSecondPageFields(blocks, document, formFields) {
  formFields['noticename'] = 'Final Rejection';
  for (let i = 0; i < blocks.length; i++) {
    let textSegments = blocks[i].layout.textAnchor.textSegments;
    secondPageFormFields(extractTextSegments(textSegments, document), formFields);
  }
  return formFields;
}

function secondPageFormFields(blockText, formFields) {
  let blockTextLower = blockText.toLowerCase();

  if (blockTextLower.includes("advisory action") && blockTextLower.includes("before the filing of an appeal brief")) {
    formFields['noticename'] = "Advisory Action";
  }
  else if (blockTextLower.includes("office action summary")) {
    formFields['noticename'] = "Office Action Summary";
  }
}

