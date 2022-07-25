function generateOldFormatTemplate(formValues) {
  try {
    if ((formValues.applicationnumber == undefined || formValues.applicationnumber == null) && (formValues.firstnamedapplicant == undefined || formValues.firstnamedapplicant == null) && (formValues.filingdate == null || formValues.filingDate == undefined)) {
      return 'pdfFormatIssue';
    }

    var firstTableData = [['In re Application of: ', formValues.firstnamedapplicant + ''],
    ['Serial No.', formValues.applicationnumber + ''],
    ['Filed ', formValues.filingdate + ''],
    ['Confirmation No. ', formValues.confirmationno + ''],
    ['Group Art Unit ', formValues.groupartunit + ''],
    ['Examiner ', formValues.examiner + ''],
    ['For ', formValues.fore + ''],
    ['Customer No. ', formValues.customerno + '']];

    var lastTableR1C2 = ' Respectfully Submitted, \n\n Signature\n \nAttorney name: ' + formValues.attorneyname + '\nReg No. ' + formValues.regno;
    var lastTableR2C1 = 'Customer No. ' + formValues.customerno + ' ' + formValues.customeraddress + '\nPhone: ' + formValues.phone + '\nFax: ' + formValues.fax + '\nEmail: ' + formValues.email;
    var lastTableData = [['\n\nDated ' + Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy"), lastTableR1C2],
    [lastTableR2C1, '']];

    var docName, noticeName = formValues.noticename;
    if (formValues.hasOwnProperty('applicationnumber') && formValues.applicationnumber != null) {
      formValues['applicationnumber'] = formValues.applicationnumber.replace(/\\n/g, '').replace('\n', '');
      docName = formValues.applicationnumber;
    }
    else {
      docName = 'empty application number';
    }

    var doc = DocumentApp.create(docName);
    var body = doc.getBody();
    body.appendParagraph('IN THE UNITED STATES PATENT AND TRADEMARK OFFICE ')
      .setHeading(DocumentApp.ParagraphHeading.HEADING3)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('');
    body.appendParagraph('');
    body.appendParagraph('');
    body.appendParagraph('');
    body.appendParagraph('');
    var table1 = body.appendTable(firstTableData);
    table1.setBorderWidth(0);
    for (let i = 0; i < table1.getNumRows(); i++) {
      table1.getCell(i, 1).editAsText().setBold(true);
    }
    table1.setColumnWidth(0, 120);
    body.appendParagraph('');
    body.appendParagraph('Mail Stop Amendment');
    body.appendParagraph(formValues.address);
    body.appendParagraph('');
    var title = body.appendParagraph('RESPONSE TO: ' + noticeName).editAsText();
    title.setAttributes(noticeStyle);
    if (noticeName != undefined) {
      title.setBold(0, noticeName.length + 12, true);
      title.setUnderline(0, noticeName.length + 12, true);
    }
    body.appendParagraph('');
    body.appendParagraph('Dear Sir/Madam ');
    body.appendParagraph('');
    body.appendParagraph('In response to the ' + noticeName + ' dated ' + formValues.datemailed + ' we are submitting the following documents, ');
    body.appendParagraph('');
    body.appendListItem(' Response to ' + noticeName);
    body.appendListItem(' Transmittal Form ');
    body.appendListItem(' ');
    body.appendParagraph('');
    body.appendParagraph('');
    body.appendParagraph('If you have any questions or need anything further, please do not hesitate to contact the Applicant’s Attorney. ');
    body.appendParagraph('');
    body.appendParagraph('');
    body.appendParagraph('The Commissioner is hereby authorized to charge any additional fees associated with this paper or during the pendency of the application, or credit any overpayment, to Deposit Account No. ');
    body.appendParagraph('');
    var table2 = body.appendTable(lastTableData);
    table2.setBorderWidth(0);

    body.setAttributes(fontStyle);
    body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING3, headingStyle);
    return doc.getId();
  } catch ({ message }) {
    if (/Service invoked too many times/.test(message) || /Service using too much computer time for one day/.test(message)) {
      return 'exceeded_limit';
    }
  }
}

function generateClaimsTemplate(formValues) {
  try {
    if ((formValues.applicationnumber == undefined || formValues.applicationnumber == null) && (formValues.firstnamedapplicant == undefined || formValues.firstnamedapplicant == null) && (formValues.filingdate == null || formValues.filingDate == undefined)) {
      return 'pdfFormatIssue';
    }

    var doc = DocumentApp.create('Claims');
    var body = doc.getBody();
    var title = body.appendParagraph('AMENDMENTS TO THE CLAIMS')
      .setHeading(DocumentApp.ParagraphHeading.HEADING3)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    title = title.editAsText();
    title.setBold(0, 23, true);
    title.setUnderline(0, 23, true);

    body.setAttributes(fontStyle);
    body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING3, headingStyle);
    return doc.getId();
  } catch ({ message }) {
    if (/Service invoked too many times/.test(message) || /Service using too much computer time for one day/.test(message)) {
      return 'exceeded_limit';
    }
  }
}

function generateApplicantArgumentsTemplate(formValues) {
  try {
    var rowOneColumnTwo = ' Respectfully Submitted, \n\n Signature\n \nAttorney name: ' + formValues.attorneyname + '\nReg No. ' + formValues.regno;
    var rowTwoColumnOne = 'Customer No. ' + formValues.customerno + ' ' + formValues.customeraddress + '\nPhone: ' + formValues.phone + '\nFax: ' + formValues.fax + '\nEmail: ' + formValues.email;
    var tableData = [['\n\nDated ' + Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy"), rowOneColumnTwo],
    [rowTwoColumnOne, '']];

    var doc = DocumentApp.create('Applicant Arguments/Remarks Made in an Amendment');
    var body = doc.getBody();
    var header = body.appendParagraph('Remarks')
      .setHeading(DocumentApp.ParagraphHeading.HEADING3)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    header = header.editAsText();
    header.setUnderline(0, 6, true);
    body.appendListItem('\n');
    body.appendListItem('\n');
    body.appendListItem('\n');
    body.appendListItem('\n');
    body.appendListItem('\n');
    for (let i = 0; i < 27; i++) {
      body.appendParagraph('');
    }
    let title = body.appendParagraph('CONCLUSION').editAsText();
    title.setBold(0, 9, true);
    title.setUnderline(0, 9, true);
    body.appendParagraph("It is believed that all the stated grounds of rejection have been adequately addressed. In view of the foregoing, Applicant respectfully submits that the present application is in condition for allowance Applicant therefore respectfully requests that the examiner reconsider and withdraw all outstanding rejections. The examiner is respectfully requested to contact the undersigned by telephone at the below listed telephone number, in order to expedite resolution of any issues and to expedite the passage of present application to issue, if any comments, suggestions or issues arise in connection with the present application. ");

    body.appendParagraph('');
    var table = body.appendTable(tableData);
    table.setBorderWidth(0);

    body.setAttributes(fontStyle);
    body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING3, headingStyle);
    return doc.getId();
  } catch ({ message }) {
    if (/Service invoked too many times/.test(message) || /Service using too much computer time for one day/.test(message)) {
      return 'exceeded_limit';
    }
  }

}

function generateAmendmentTemplate(formValues) {
  try {
    var docName = 'Empty OA Name';
    if (formValues.noticename == 'Advisory Action') {
      docName = "Amendment Submitted/Entered with Filing of CPA/RCE";
    }
    else if (formValues.noticename == 'Office Action Summary') {
      docName = "Office Action Summary";
      formValues.noticename = "Office Action Summary ";
    }
    else if (formValues.noticename != null && formValues.noticename != undefined) {
      docName = formValues.noticename;
    }
    var doc = DocumentApp.create(docName);
    var tableData = [['In re Application of: ', formValues.firstnamedapplicant + ''],
    ['Serial No.', formValues.applicationnumber + ''],
    ['Filed ', formValues.filingdate + ''],
    ['Confirmation No. ', formValues.confirmationno + ''],
    ['Group Art Unit ', formValues.groupartunit + ''],
    ['Examiner ', formValues.examiner + ''],
    ['For ', formValues.fore + ''],
    ['Customer No. ', formValues.customerno + '']];

    var body = doc.getBody();
    body.appendParagraph('IN THE UNITED STATES PATENT AND TRADEMARK OFFICE')
      .setHeading(DocumentApp.ParagraphHeading.HEADING3)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('');
    body.appendParagraph('');
    body.appendParagraph('');
    body.appendParagraph('');
    body.appendParagraph('');
    var table = body.appendTable(tableData);
    table.setBorderWidth(0);
    for (let i = 0; i < table.getNumRows(); i++) {
      table.getCell(i, 1).editAsText().setBold(true);
    }
    table.setColumnWidth(0, 120);
    body.appendParagraph('');
    body.appendParagraph('Mail Stop Amendment');
    body.appendParagraph(formValues.address);
    body.appendParagraph('');
    var title = body.appendParagraph('RESPONSE TO: ' + formValues.noticename).editAsText();
    title.setAttributes(noticeStyle);
    if (formValues.noticename != null && formValues.noticename != undefined) {
      title.setBold(0, formValues.noticename.length + 12, true);
      title.setUnderline(0, formValues.noticename.length + 12, true);
    }
    body.appendParagraph('');
    body.appendParagraph('Dear Sir/Madam ');
    body.appendParagraph('');
    body.appendParagraph('In response to the ' + formValues.noticename + ' dated ' + formValues.datemailed + ' we request your reconsideration in view of the following amendments and remarks ');
    body.appendParagraph('');
    body.appendListItem("Amendments to the Specification begin on page  of this paper").setGlyphType(DocumentApp.GlyphType.SQUARE_BULLET);
    body.appendListItem("Amendments to the Claims are reflected in the listing of the claims which begin on page  of this paper").setGlyphType(DocumentApp.GlyphType.SQUARE_BULLET);

    body.appendListItem("Amendments to the Drawings begin on page  of this paper").setGlyphType(DocumentApp.GlyphType.SQUARE_BULLET);
    body.appendListItem("The Commissioner is hereby authorized to charge any additional fees associated with this paper or during the pendency of the application, or credit any overpayment, to Deposit Account No.").setGlyphType(DocumentApp.GlyphType.SQUARE_BULLET);

    body.setAttributes(fontStyle);
    body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING3, headingStyle);
    return doc.getId();
  } catch ({ message }) {
    Logger.log('Exception: generateAmendmentTemplate: ' + message);
    if (/Service invoked too many times/.test(message) || /Service using too much computer time for one day/.test(message)) {
      return 'exceeded_limit';
    }
  }

}
