// KEN M. HAGGERTY
// GOOGLE APPS SCRIPT
// GOOGLE FORM EMAIL (BEIGoogleFormEmail)
// VERSION : 0.1
// FOR     : DAVID HESTRIN
// CREATED : JUN 28 2017
// EDITED  : JUN 28 2017

// NOTES

// > This code is derived from BEIProduceFormAppScript and updated to make it more generic.
// > Make sure column A is blank for all non-response rows.

// CONSTANTS

var testRow = 4;

var responsesSheetName = "Form Responses 1";
var startingRow = 3;

// MENU ITEMS

function onOpen() {
  Logger.log("[METHOD] onOpen");
  
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Email')
      .addItem('Compose email...', 'showEmailComposer')
      .addToUi();
}

function showEmailComposer() {
  Logger.log("[METHOD] showEmailComposer");
  
  // Generate Pane UI
  
  var ui = SpreadsheetApp.getUi();
  
  Logger.log("[SCRIPT COMPLETE]");
}

// EMAIL GENERATION

function sendEmail() {
  Logger.log("[METHOD] sendEmail");
  
  // Send Email
  
  emailAddress = "kenmhaggerty@gmail.com"; // temp
  emailSubject = "subject"; // temp
  emailBody = "body"; // temp
  
  MailApp.sendEmail({
    to: emailAddress,
    subject: emailSubject,
    htmlBody: emailBody
  });
}

// ADDITIONAL FUNCTIONS

function getRow(index) {
  Logger.log("[METHOD] getRow");
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(responsesSheetName);
  var numOfColumns = sheet.getLastColumn();
  var range = sheet.getRange(index, 1, 1, numOfColumns);
  var values = range.getValues();
  return values[0];
}

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

function getColumnForHeader(header) {
  Logger.log("[METHOD] getColumnForHeader");
  
  var headerRow = getRow(1);
  var column = headerRow.indexOf(header);
  return column;
}

function getHTMLTable(array, headers, cellStyle, rowStyle, tableStyle) {
  Logger.log("[METHOD] getHTMLTable");
  
  var htmlTable = "";
  
  var row, htmlRow, item;
  for (var i = 0; i < array.length; i++) {
    row = array[i];
    htmlRow = getHTMLRow(row, cellStyle, rowStyle);
    htmlTable += htmlRow;
  }
  
  if (Object.prototype.toString.call(headers) === '[object Array]') {
    htmlRow = getHTMLRow(headers, cellStyle, rowStyle);
    htmlTable = htmlRow + htmlTable;
  }
  
  var tags = "";
  if (Object.prototype.toString.call(tableStyle) === '[object String]') {
    tags = " style='"+tableStyle+"'";
  }
  
  htmlTable = "<table"+tags+">"+htmlTable+"</table>";
  return htmlTable;
}

function getHTMLRow(array, cellStyle, rowStyle) {
  Logger.log("[METHOD] getHTMLRow");
  
  var tags = "";
  if (Object.prototype.toString.call(cellStyle) === '[object String]') {
    tags = " style='"+cellStyle+"'";
  }
  
  htmlRow = "";
  var item;
  for (var i = 0; i < array.length; i++) {
    item = array[i];
    htmlRow += "<td"+tags+">"+item+"</td>";
  }
  
  tags = "";
  if (Object.prototype.toString.call(rowStyle) === '[object String]') {
    tags = " style='"+rowStyle+"'";
  }
  
  htmlRow = "<tr"+tags+">"+htmlRow+"</tr>";
  return htmlRow;
}
