// KEN M. HAGGERTY
// GOOGLE APPS SCRIPT – Code.gs
// GOOGLE FORM EMAIL (BEIGoogleFormEmail)
// VERSION : 0.1
// FOR     : DAVID HESTRIN
// CREATED : JUN 28 2017
// EDITED  : JUN 29 2017

// NOTES

// > This code is derived from BEIProduceFormAppScript and updated to make it more generic.
// > Make sure column A is blank for all non-response rows.

////////// CONSTANTS //////////

var responsesSheetName = "Form Responses 1";
var startingRow = 2;

////////// SETUP //////////

function onOpen() {
  Logger.log("[METHOD] onOpen");
  
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Email")
    .addItem("Email comments...", "emailComments")
    .addToUi();
}

////////// ACTIONS //////////

function emailComments() {
  Logger.log("[METHOD] emailComments");

  // Email column prompt UI

  var emailResponse = getResponseForPrompt("Please enter column for email address", "where column A = 1:");
  var emailColumn = convertTextToPositiveInteger(emailResponse, "Invalid column", "Could not get emails from column "+emailResponse+".");
  if (isNaN(emailColumn)) {
    return;
  }

  // Comment column prompt UI

  var commentResponse = getResponseForPrompt("Please enter column for email body", "where column A = 1:");
  var commentColumn = convertTextToPositiveInteger(commentResponse, "Invalid column", "Could not get comments from column "+commentResponse+".");
  if (isNaN(commentColumn)) {
    return;
  }

  // Email subject prompt UI

  var subject = getResponseForPrompt("Please enter subject for email", "");

  // Generate and send emails
  
  var responses = getResponseRows();
  var row, email, comment;
  var count = 0;
  for (i in responses) {
    row = responses[i];
    email = row[emailColumn-1];
    comment = row[commentColumn-1];
    sendEmail(email, subject, comment);
    count += 1;
  }

  // Done
  
  var ui = SpreadsheetApp.getUi();
  ui.alert("Emails successfully sent", count+" emails sent.", ui.ButtonSet.OK);

  Logger.log("[DONE] emailComments");
}

////////// USER INTERFACE //////////

function getResponseForPrompt(title, body) {
  Logger.log("[METHOD] getResponseForPrompt");
  
  var ui = SpreadsheetApp.getUi();
  
  var prompt = ui.prompt(title, body, ui.ButtonSet.OK_CANCEL);
  
  var button = prompt.getSelectedButton();
  if ((button == ui.Button.CANCEL) || (button == ui.Button.CLOSE)) {
    return null;
  }
  
  return prompt.getResponseText();
}

function convertTextToPositiveInteger(text, errorTitle, errorMessage) {
  Logger.log("[METHOD] convertTextToPositiveInteger");

  if (isNaN(text) || (Number(text) % 1 != 0) || (Number(text) <= 0)) {
    var ui = SpreadsheetApp.getUi();
    ui.alert(errorTitle, errorMessage, ui.ButtonSet.OK);
    return NaN;
  }

  return Number(text);
}

////////// EMAIL //////////

function sendEmail(recipient, subject, htmlBody) {
  Logger.log("[METHOD] sendEmail");

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: htmlBody
  });
}

////////// GOOGLE SHEET //////////

function getResponseRows() {
  Logger.log("[METHOD] getResponseRows");
  
  var rows = SpreadsheetApp.getActiveSheet().getDataRange().getValues();
  for (i = startingRow-1; i < rows.length; i++) {
    if (rows[i][0] == "") {
      break;
    }
  }
  rows = rows.splice(startingRow-1, i+1-startingRow);
  return rows;
}
