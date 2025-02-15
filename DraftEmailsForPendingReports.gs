var TrackingPersonCol;
var ListedDateCol;
var BranchCol;
var LoadingDateCol;
var ExpectedDeliveryDateCol;
var TxnNoCol;
var TruckNoCol;
var DrivermobilenumberCol;
var FromCityCol;
var ToCityCol;
var StatusCol;
var SubStatusCol;
var RemarksCol;
var NearUnloadingDateCol;
var NearUnloadingTimeCol;
var ReportingDateCol;
var ReportingTimeCol;
var UnloadingDateCol;
var UnloadingTimeCol;
var OrderReceivedByCol;
var TruckPlacedByCol;
var TransporterNameCol;
var TruckerNameCol;
var TruckTypeCol;
var InformedSupplyTeamCol;
var MailSentOnCol;
var PriorityVehicleCol;
var GPSConsentedcol;
var ReportEmailsFilterCol;
var TATCol;

function NewAssignHeaderColumnNumbers1() {
  var headerRow = 2;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() == "Tracking Data") {
    headerRow = 6;
  }
  var headerData = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  TrackingPersonCol = headerData.indexOf("Tracking Person") + 1;
  ListedDateCol = headerData.indexOf("Listed Date") + 1;
  BranchCol = headerData.indexOf("Branch") + 1;
  LoadingDateCol = headerData.indexOf("Loading Date") + 1;
  ExpectedDeliveryDateCol = headerData.indexOf("Expected Delivery Date") + 1;
  TxnNoCol = headerData.indexOf("Txn No") + 1;
  TruckNoCol = headerData.indexOf("Truck No") + 1;
  DrivermobilenumberCol = headerData.indexOf("Driver mobile number") + 1;
  FromCityCol = headerData.indexOf("From City") + 1;
  ToCityCol = headerData.indexOf("To City") + 1;
  StatusCol = headerData.indexOf("Status") + 1;
  SubStatusCol = headerData.indexOf("Sub Status") + 1;
  RemarksCol = headerData.indexOf("Remarks") + 1;
  NearUnloadingDateCol = headerData.indexOf("Near Unloading Date") + 1;
  NearUnloadingTimeCol = headerData.indexOf("Near Unloading Time") + 1;
  ReportingDateCol = headerData.indexOf("Reporting Date") + 1;
  ReportingTimeCol = headerData.indexOf("Reporting Time") + 1;
  UnloadingDateCol = headerData.indexOf("Unloading Date") + 1;
  UnloadingTimeCol = headerData.indexOf("Unloading Time") + 1;
  OrderReceivedByCol = headerData.indexOf("Order Received By") + 1;
  TruckPlacedByCol = headerData.indexOf("Truck Placed By") + 1;
  TransporterNameCol = headerData.indexOf("Transporter Name") + 1;
  TruckerNameCol = headerData.indexOf("Trucker Name") + 1;
  TruckTypeCol = headerData.indexOf("Truck Type") + 1;
  InformedSupplyTeamCol = headerData.indexOf("Informed Supply Team") + 1;
  MailSentOnCol = headerData.indexOf("Mail Sent On") + 1;
  PriorityVehicleCol = headerData.indexOf("Priority Vehicle") + 1;
  GPSConsentedcol = headerData.indexOf("GPS Consented") + 1;
  ReportEmailsFilterCol = headerData.indexOf("Report emails filter") + 1; // Column AC
  TATCol = headerData.indexOf("TAT") + 1; // Column AF
  
}

function DraftEmailsForPendingReports() {
  // Assign column numbers to global variables
  NewAssignHeaderColumnNumbers1();

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var userName = Session.getActiveUser().getEmail();
  userName = userName.replace("@lobb.in", "");


//|| reportEmailsFilter === "issue report pending" added to check new condition
  for (var row = 6; row <= lastRow; row++) {
    var reportEmailsFilter = sheet.getRange(row, ReportEmailsFilterCol).getValue();
    if (reportEmailsFilter === "delayreport pending"|| reportEmailsFilter === "issue report pending" ) {
      var mailSentOn = sheet.getRange(row, MailSentOnCol).getValue();
      if (mailSentOn !== "") {
        var response = responseYesNoDialog("Mail already sent", "Do you want to send again?");
        if (!response) {
          continue;
        }
      }

      SendEmailold(row, sheet, userName);

      // Update sheet with sent data/time and user name
      var sentDateTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
      sheet.getRange(row, MailSentOnCol).setValue(mailSentOn + " ; " + userName + ": " + sentDateTime);
    }
  }
}

function SendEmailold(activeRow, sheet, userName) {
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");

  // Get tracking data values
  var branchName = sheet.getRange(activeRow, BranchCol).getValue();
  if (branchName === null) {
    showMessageBoxOK("Warning", "Branch Name missing; Select a row that has tracking data");
    return;
  }
  var subStatus = sheet.getRange(activeRow, SubStatusCol).getValue();
  var status = sheet.getRange(activeRow, StatusCol).getValue();
  if (status == "In Transit" && subStatus == "Language Issue") {
    showMessageBoxOK("Warning", "You can't send mail for this remarks: " + subStatus);
    return;
  }

  var loadingDate = sheet.getRange(activeRow, LoadingDateCol).getValue();
  var transaction = sheet.getRange(activeRow, TxnNoCol).getValue();
  var truckNo = sheet.getRange(activeRow, TruckNoCol).getValue();
  var driverNo = sheet.getRange(activeRow, DrivermobilenumberCol).getValue();
  var truckerName = sheet.getRange(activeRow, TruckerNameCol).getValue();
  var fromCity = sheet.getRange(activeRow, FromCityCol).getValue();
  var toCity = sheet.getRange(activeRow, ToCityCol).getValue();
  var orderReceivedBy = sheet.getRange(activeRow, OrderReceivedByCol).getValue();
  var truckPlacedBy = sheet.getRange(activeRow, TruckPlacedByCol).getValue();
  var gpsconsented = sheet.getRange(activeRow, GPSConsentedcol).getValue();
  var ExpectedDeliveryDate = sheet.getRange(activeRow, ExpectedDeliveryDateCol).getValue();
  var TAT = sheet.getRange(activeRow, TATCol).getValue();

  var ccEmail = worksheetFunctionQuery(settingsSheet, "A:H", "A", "H", branchName);
  if (!ccEmail || ccEmail === "") {
    showMessageBoxOK("Warning", "Branch Name is invalid or CC Mail is blank; Inform your Team Lead or Manager to fix the data");
    return;
  }

  var orderReceivedByEmail = worksheetFunctionQuery(settingsSheet, "A:J", "I", "J", orderReceivedBy);
  if (!orderReceivedByEmail || orderReceivedByEmail === "") {
    showMessageBoxOK("Warning", "Order Received By Name is invalid, sent email to branch manager");
    orderReceivedByEmail = ccEmail;
  }

  var truckPlacedByEmail = worksheetFunctionQuery(settingsSheet, "A:J", "I", "J", truckPlacedBy);
  if (!truckPlacedByEmail || truckPlacedByEmail === "") {
    showMessageBoxOK("Warning", "Truck Placed By Name is invalid, sent email to branch manager");
    truckPlacedByEmail = ccEmail;
  }

  var recipientEmail = orderReceivedByEmail == truckPlacedByEmail ? ccEmail : orderReceivedByEmail + ";" + truckPlacedByEmail;

  var subject = branchName + " - Tracking -" + transaction + " Status : " + status;
  if (TAT.toLowerCase().includes("delay")) {
    subject += ": "+TAT ;
  }
  var body = "Truck Details\n\n";
  if (loadingDate instanceof Date) {
    loadingDate = Utilities.formatDate(loadingDate, Session.getScriptTimeZone(), "dd-MM-yyyy");
  }
  if (ExpectedDeliveryDate instanceof Date) {
    ExpectedDeliveryDate = Utilities.formatDate(ExpectedDeliveryDate, Session.getScriptTimeZone(), "dd-MM-yyyy");
  }

  body += "TAT Reporting Date at destination  : " + TAT + "\n";
  body += "Loading Date: " + loadingDate + "\n";
  body += "Expected Delivery Date: " + ExpectedDeliveryDate + "\n";
  body += "Trucker Name: " + truckerName + "\n";
  body += "Truck Number: " + truckNo + "\n";
  body += "Truck Driver Number: " + driverNo + "\n";
  body += "From City: " + fromCity + "\n";
  body += "To City: " + toCity + "\n";
  body += "Order Received By: " + orderReceivedBy + "\n";
  body += "Truck Placed By: " + truckPlacedBy + "\n";
  body += "GPS Consent status: " + gpsconsented + "\n";
  body += "Status: " + subStatus + ".\n\nProvide needed information and/or take suitable action at your side. \n\nThis is a manually triggered email.";

  // Create the draft email
  GmailApp.createDraft(
    recipientEmail,
    subject,
    body, {
      cc: ccEmail + ";riyaz.blr@lobb.in;alfiz.sha@lobb.in;mylari.gupta@lobb.in;naveenkumar.m@lobb.in"
    }
  );
}

function worksheetFunctionQuery(sheet, range, searchCol, returnCol, searchKey) {
  var dataRange = sheet.getRange(range);
  var data = dataRange.getValues();
  var searchColIndex = searchCol.charCodeAt(0) - 65; // Convert column letter to index
  var returnColIndex = returnCol.charCodeAt(0) - 65; // Convert column letter to index

  for (var i = 0; i < data.length; i++) {
    var checkSearchKey = data[i][searchColIndex]
    if (data[i][searchColIndex] == searchKey) {
      return data[i][returnColIndex];
    }
  }
  return null;
}

function showMessageBoxOK(title, message) {
  SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function responseYesNoDialog(title, message) {
  var response = SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.YES_NO);
  return response == SpreadsheetApp.getUi().Button.YES;
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Bulk Draft_mail')
    .addItem('Draft Emails for Delay Report Pending', 'DraftEmailsForPendingReports')
    .addToUi();
}
