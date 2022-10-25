function moveValuesonNightAuditOffenders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hotelpaste = ss.getRange("Hotelpaste").getValue();
  var source = ss.getRange("MoveColumn");
  var destSheet = ss.getSheetByName("Night Audit Offenders");
  //Day before needs to be copied and saved
  if (hotelpaste != 1) {
    var destRange = destSheet.getRange("A9").offset(0, hotelpaste - 1);
    source.copyTo(destRange, { contentsOnly: true });
  }
  //Now forumlas here
  var destRange = destSheet.getRange("A9").offset(0, hotelpaste);
  source.copyTo(destRange);
  //clears source sheet

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName("List"));
  ss.getRange("List!E4:E150").clearContent();

  if (hotelpaste == 1) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.setActiveSheet(ss.getSheetByName("Night Audit Offenders"));
    ss.getRange("NAOClear").clearContent();
  }
}

/* Code adopted from https://code.google.com/p/google-apps-script-issues/issues/detail?id=3579#c25
Updated since oAuthConfig is deprecated
http://ctrlq.org/code/19869-email-google-spreadsheets-pdf
*/

/* Send Spreadsheet in an email as PDF, automatically */
function emailNightAuditReportasPDF() {
  // Get the currently active spreadsheet URL (link)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName("Night Audit Offenders"));
  var url = ss.getUrl();
  url = url.replace(/edit$/, "");

  // Subject of email message
  var hotelsMissing = ss.getRange("hotelsMissing").getValue();

  var runDate = ss.getRange("runDate").getValue();

  if (hotelsMissing > 0) {
    var subject =
      "Night Audit Reports - " +
      runDate +
      " - " +
      hotelsMissing +
      " Hotels Missing";
  } else {
    var subject = "Night Audit Reports - " + runDate + " - All Hotels Reported";
  }
  /* Specify PDF export parameters



exportFormat = pdf / csv / xls / xlsx
gridlines = true / false
printtitle = true (1) / false (0)
size = legal / letter/ A4
fzr (repeat frozen rows) = true / false
portrait = true (1) / false (0)
fitw (fit to page width) = true (1) / false (0)
add gid if to export a particular sheet - 0, 1, 2,..



*/

  var url_ext =
    "export?exportFormat=pdf&format=pdf" + // export as pdf
    "&size=letter" + // paper size
    "&portrait=false" + // orientation, false for landscape
    "&fitw=true" + // fit to width, false for actual size
    "&sheetnames=false&printtitle=false" + // hide optional headers and footers
    "&pagenumbers=true&gridlines=false" + // show page numbers and hide gridlines
    "&fzr=true" + // do not repeat row headers (frozen rows) on each page
    "&gid="; // the sheet's Id

  var token = ScriptApp.getOAuthToken();

  // Convert individual worksheets to PDF
  var response = UrlFetchApp.fetch(url + url_ext + ss.getSheetId(), {
    headers: {
      Authorization: "Bearer " + token,
    },
  });

  //convert the response to a blob and store in our array
  pdfBlob = response
    .getBlob()
    .setName(
      ss.getRange("runDate").getValue() +
        "-Night Audit Report Files Received.pdf"
    );

  // Email subject and message body
  //var email = "zachary.ruben@hawkeyehotels.com, angie.patel@hawkeyehotels.com, rdos@hawkeyehotels.com, bob.patel@hawkeyehotels.com, cassandra.rule@hawkeyehotels.com"; // Send the PDFs to this email address
  var email = "corpstaff@hawkeyehotels.com";

  //var email = "zachary.ruben@hawkeyehotels.com";

  var message =
    "Please see attached for Night Audit Report Files Received on " +
    runDate +
    ". ";

  if ((hotelsMissing = 0)) {
    message = message + "All hotels reported today.";
  } else {
    message = message + "These hotels did not send any reports today: <br><br>";
    var thenumber = 0;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.setActiveSheet(ss.getSheetByName("List"));
    var hotelcount = ss.getRange("G2").getValue();
    for (
      var theincrementnumber = 4;
      theincrementnumber <= hotelcount + 3;
      theincrementnumber++
    ) {
      if (ss.getRange("F" + theincrementnumber).getValue() != "No") {
        if (
          ss.getRange("E" + theincrementnumber).getValue() == 0 ||
          ss.getRange("E" + theincrementnumber).getValue() == ""
        ) {
          message =
            message +
            thenumber +
            ". " +
            ss.getRange("B" + theincrementnumber).getValue() +
            "<br><br>";
          thenumber = thenumber + 1;
        }
      }
    }
  }
  // save the file to the root folder of Google Drive

  // If allowed to send emails, send the email with the PDF attachment

  //GmailApp.sendEmail(email, subject, message, {attachments:[pdfBlob]});

  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: message,
    name: "Hawkeye Reports",
    attachments: [pdfBlob],
  });
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName("Night Audit Offenders"));
  var corporatefolder = "Corporate";
  var reporting = "Reports";
  var reportfolder = "Night Audit Report Files Received Report";
  var yearfolder = ss.getRange("yearOffender").getValue();
  var monthfolder = ss.getRange("monthdigit").getValue();

  var baseFolder =
    "//" +
    corporatefolder +
    "/" +
    reporting +
    "/" +
    reportfolder +
    "/" +
    yearfolder +
    "/" +
    monthfolder;

  var folderOrig = getDriveFolder(baseFolder);
  folderOrig.createFile(pdfBlob);

  moveValuesonNightAuditOffenders();
}

function clearCount() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName("List"));
  ss.getRange("List!E4:E149").clearContent();
}
