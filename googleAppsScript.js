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

var paypalUsername = "betterearthinstitute";
var responsesSheetName = "Form Responses 1";
var startingRow = 3;
var testRow = 4;

var cellStyle = null;
var rowStyle = null;
var tableStyle = "border-spacing: 10px 0px;";

var emailKey = "Email Address";
var timestampKey = "Timestamp";
var deliveryKey = "What is your address for delivery?";
var commentsKey = "If you want anything not listed or have any comments, suggestions, or requests please let me know here:";

var sectionKey = "section";
var productKey = "product";
var priceKey = "price";
var unitsKey = "unit";
var quantityKey = "quantity";
var descriptionKey = "description";

// MENU ITEMS

function onOpen() {
  Logger.log("[METHOD] onOpen");
  
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Email')
      .addItem('Email row...', 'showSendEmailPrompt')
      .addItem('Email all respondents', 'showEmailAllPrompt')
      .addItem('Email orders total', 'showEmailOrdersTotalPrompt')
      .addToUi();
}

function showSendEmailPrompt() {
  Logger.log("[METHOD] showSendEmailInput");
  
  // Generate Prompt UI
  
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt(
    "Send Manual Email Receipt",
    "Please enter row number:",
    ui.ButtonSet.OK_CANCEL
  );

  // Process Response
  
  var button = result.getSelectedButton();
  
  if ((button == ui.Button.CANCEL) || (button == ui.Button.CLOSE)) {
    return;
  }
  
  var response = result.getResponseText();
  if (isNaN(response) || (Number(response) % 1 != 0) || (Number(response) < startingRow)) {
    ui.alert(
      "Invalid Row",
      "Could not send email for row "+response+".",
      ui.ButtonSet.OK
    );
    return;
  }
  
  var index = Number(response);
  var row = getRow(index);
  
  // Verify email
  
  index = getColumnForHeader(emailKey);
  var email = row[index];
  
  button = ui.alert("Is this the right email?", email, ui.ButtonSet.YES_NO);

  if (button == ui.Button.CLOSE) {
    return;
  }
  
  if (button == ui.Button.NO) {
    showSendEmailPrompt();
    return;
  }
  
  sendEmailForRow(row);
  
  // Done
  
  ui.alert("Email sent successfully.");
  
  Logger.log("[SCRIPT COMPLETE]");
}

function showEmailAllPrompt() {
  Logger.log("[METHOD] showEmailAllPrompt");
  
  // Extract Rows
  
  var rows = getResponseRows();
  var count = rows.length;
  
  // Generate Prompt UI
  
  var ui = SpreadsheetApp.getUi();
  
  if (count <= 0) {
    ui.alert(
      "0 Emails Sent",
      "Could not find respondents to email.",
      ui.ButtonSet.OK
    );
    return;
  }

  var button = ui.alert(
    "Email all respondents?",
    "This will send "+count+" emails.",
    ui.ButtonSet.OK_CANCEL
  );

  // Process Response
  
  if ((button == ui.Button.CANCEL) || (button == ui.Button.CLOSE)) {
    return;
  }
  
  for (var i = 0; i < count; i++) {
    sendEmailForRow(rows[i]);
  }
  
  // Done
  
  ui.alert(count+" emails sent.");
  
  Logger.log("[SCRIPT COMPLETE]");
}

function showEmailOrdersTotalPrompt() {
  Logger.log("[METHOD] showEmailOrdersTotalPrompt");
  
  // Extract Rows
  
  var rows = getResponseRows();
  var count = rows.length;
  
  // Generate Prompt UI
  
  var ui = SpreadsheetApp.getUi();
  
  if (count <= 0) {
    ui.alert(
      "0 Orders Found",
      "No email was generated because no orders could be found.",
      ui.ButtonSet.OK
    );
    return;
  }

  var result = ui.prompt(
    "Please Enter Email",
    "This will send an email summary of "+count+" orders to the provided address:",
    ui.ButtonSet.OK_CANCEL
  );
  
  // Process Response
  
  var button = result.getSelectedButton();
  
  if ((button == ui.Button.CANCEL) || (button == ui.Button.CLOSE)) {
    return;
  }
  
  var email = result.getResponseText();
  
  // Sum Orders
  
  var totals = rows[0];
  var row, value;
  for (var i = 1; i < count; i++) {
    row = rows[i];
    for (j in row) {
      value = Number(row[j]);
      if (isNaN(value)) {
        continue;
      }
      totals[j] = Number(totals[j])+value;
    }
  }
  
  var index = getColumnForHeader(timestampKey);
  totals[index] = Date();
  
  index = getColumnForHeader(emailKey);
  totals[index] = email;
  
  index = getColumnForHeader(commentsKey);
  totals[index] = "";
  
  index = getColumnForHeader(deliveryKey);
  totals[index] = "";
  
  // Send Email
  
  sendEmailForRow(totals);
  
  // Done
  
  ui.alert("Email sent successfully.");
  
  Logger.log("[SCRIPT COMPLETE]");
}

// TESTING

function test() {
  Logger.log("[METHOD] test");
  
  // Extract Range
  
  var row = getRow(testRow);
  
  // Generate Email
  
  sendEmailForRow(row);
  
  // Done
  
  Logger.log("[SCRIPT COMPLETE]");
}

// FORM SUBMISSION

function onFormSubmit(e) {
  Logger.log("[METHOD] onFormSubmit");
  
  // Extract Range
  
  var range = e.range;
  var values = range.getValues();
  var row = values[0];
  
  // Generate Email
  
  sendEmailForRow(row);
  
  // Complete
  
  Logger.log("[SCRIPT COMPLETE]");
}

// EMAIL GENERATION

function sendEmailForRow(row) {
  Logger.log("[METHOD] sendEmailForRow");
  
  // Generate Line Items
  
  var lineItems = getLineItemsForRow(row);
  
  // Extract Additional Info
  
  var headerRow = getRow(1);
  var index;
  
  index = headerRow.indexOf(emailKey);
  var emailAddress = row[index];
  
  index = headerRow.indexOf(timestampKey);
  var orderDate = row[index];
  
  index = headerRow.indexOf(deliveryKey);
  var deliveryAddress = row[index];
  
  index = headerRow.indexOf(commentsKey);
  var comments = row[index];
  
  // Generate HTML Itemized Receipt
  
  var array = [];
  var subtotal = 0;
  var lineItem, product, quantity, row, unit, quantityString, unitCost, costString, cost;
  for (i in lineItems) {
    lineItem = lineItems[i];
    if (!(quantityKey in lineItem) || (lineItem[quantityKey] == undefined)) {
      continue;
    }
    
    product = lineItem[productKey];
    quantity = lineItem[quantityKey];
    
    var quantityString, unit, unitCost, cost, costString;
    if (quantity == "case") {
      quantityString = "case";
      costString = "varies";
    }
    else {
      quantityString = quantity.toString();
      unit = lineItem[unitsKey];
      if (unit) {
        quantityString += (" "+unit);
      }
      unitCost = lineItem[priceKey];
      cost = quantity*unitCost;
      costString = currency(cost/100.0, "$");
      subtotal += cost;
    }
    
    row = [product, quantityString, costString];
    array.push(row);
  }
  subtotal = subtotal/100.0;
  var subtotalString = currency(subtotal, "$");
  array.push(["", "subtotal", subtotalString]);
  var headers = ["Item", "Quantity", "Price"];
  var htmlTable = getHTMLTable(array, headers, null, null, tableStyle);
  
  var total = subtotal; // temporary
  
  // Generate PayPal Link
  
  var htmlPayPal = "<p><a href='https://www.paypal.me/"+paypalUsername+"/"+total+"'>Click here to pay "+currency(total, "$")+" via PayPal</a></p>";
  
  // Generate HTML Comments / Suggestions
  
  var htmlComments = "<p>Comments / Suggestions / Requests:<br />"+comments+"</p>";
  
  // Generate HTML Delivery Address
  
  var htmlDelivery = "<p>Delivery Address:<br />"+deliveryAddress+"</p>";
  
  // Send Email
  
  emailSubject = "Your Produce Receipt";
  emailBody = htmlTable+htmlPayPal+htmlComments+htmlDelivery;
  
  MailApp.sendEmail({
    to: emailAddress,
    subject: emailSubject,
    htmlBody: emailBody
  });
}

// PRODUCE PRICE GENERATION

var regex = /[ \t]*(.*\w)?[ \t]*\[([^\$]*\w)[\s\W]*(\$(\d*(\.\d+)?))\s*(each|per\s(\w+))([ \t]+(.*))?\]/

// 0 = (value)
// 1 = Section
// 2 = Product
// 3 = Price (must include $)
// 4 = Price (w/o $ sign)
// 5 = Cents (unused)
// 6 = Allocation (“each” or “per ____”)
// 7 = Unit (if “per ____”)
// 8 = (unused)
// 9 = Description

function getLineItemsForRow(row) {
  Logger.log("[METHOD] getLineItemsForRow");
  
  // Get Header Row
  
  var headerRow = getRow(1);
  
  // Generate Dictionary
  
  var lineItems = [];
  var skipRow = getRow(2);
  var header, components, quantity, section, product, price, units, description, dictionary;
  for (column in headerRow) { // 0-indexed
    if (skipRow[column] == "SKIP") {
      continue;
    }
    
    header = headerRow[column];
    components = header.match(regex);
    quantity = row[column];
    if ((components == null) || (row[column] <= 0)) {
      continue;
    }
    
    section = components[1];
    product = components[2];
    price = Number(components[4])*100;
    units = components[7];
    description = components[9];
    
    dictionary = {};
    dictionary[sectionKey] = section;
    dictionary[productKey] = product;
    dictionary[priceKey] = price;
    dictionary[unitsKey] = units;
    dictionary[quantityKey] = quantity;
    dictionary[descriptionKey] = description;
    
    lineItems.push(dictionary);
  }
  
  return lineItems;
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

function currency(value, symbol) {
  return symbol+parseFloat(Math.round(value*100)/100).toFixed(2).toString();
}
