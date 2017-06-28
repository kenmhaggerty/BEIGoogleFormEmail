// KEN M. HAGGERTY
// GOOGLE APPS SCRIPT – Code.gs
// GOOGLE FORM EMAIL (BEIGoogleFormEmail)
// VERSION : 0.1
// FOR     : DAVID HESTRIN
// CREATED : JUN 28 2017
// EDITED  : JUN 28 2017

// NOTES

// > This code is derived from BEIProduceFormAppScript and updated to make it more generic.
// > Make sure column A is blank for all non-response rows.

////////// CONSTANTS //////////

var selectorColumn = 1; // where column A = 1
var testRow = 2;

var responsesSheetName = "Form Responses 1";
var startingRow = 2;

////////// SETUP //////////

function onOpen() {
  Logger.log("[METHOD] onOpen");
  
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Email")
      .addItem("Test", "test")
      .addToUi();
}

////////// DEBUGGING //////////

function test() {
  Logger.log("[METHOD] test");
  
  var options = getSelectorOptions();
  Logger.log(options);
  
  Logger.log("[DONE] test");
}

////////// DATA //////////

function getSelectorOptions() {
  Logger.log("[METHOD] getSelectorOptions");
  
  var rows = getResponseRows();
  var options = [];
  for (i = 0; i < rows.length; i++) {
    options.push(rows[i][selectorColumn-1]);
  }
  return options;
}

////////// GOOGLE SHEET //////////

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
