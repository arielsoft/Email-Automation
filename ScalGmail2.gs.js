function ScalGmail2() {
  // Default Drive folder where to archive messages
  var sheet = SpreadsheetApp.getActiveSheet();
  var values = SpreadsheetApp.getActiveSheet()
    .getRange(1, 1, 70, 17)
    .getValues();
  //SpreadsheetApp.getActiveSheet().getRange integer row, integer column, num rows, num columns

  //var hotelcount = sheet.getRange("G1").getValue();
  var hotelcount = values[1][6];
  var reportfolder = "Night Audits";
  //var yearfolder = sheet.getRange("B1").getValue();
  var yearfolder = values[0][1];
  //var monthfolder = sheet.getRange("Q1").getValue();
  var monthfolder = values[0][16];
  var mainfolder = "Hotels";
  var corporatefolder = "Corporate";
  var reporting = "Reports";

  //for(var theincrementnumber = 4; theincrementnumber <= hotelcount + 3; theincrementnumber++){
  for (
    var theincrementnumber = 3;
    theincrementnumber <= hotelcount + 2;
    theincrementnumber++
  ) {
    //var gmailLabels = sheet.getRange("A" + theincrementnumber).getValue();
    var gmailLabels = values[theincrementnumber][0];
    //var hotelfolder = sheet.getRange("B" + theincrementnumber).getValue();
    var hotelfolder = values[theincrementnumber][1];
    //var filetypes = "all";
    //var brandfolder = sheet.getRange("C" + theincrementnumber).getValue();
    var brandfolder = values[theincrementnumber][2];

    //var QuantityofFiles = Number(sheet.getRange("E" + theincrementnumber).getValue());
    var QuantityofFiles = Number(values[theincrementnumber][4]);
    //---------------------------

    // Get the label
    var label = GmailApp.getUserLabelByName(gmailLabels);
    var threadsArr = getThreadsForLabel(label);

    /*if(threadsArr.length == 0){
Browser.msgBox("a");
}*/

    for (var j = 0; j < threadsArr.length; j++) {
      var messagesArr = getMessagesforThread(threadsArr[j]);
      for (var k = 0; k < messagesArr.length; k++) {
        var messageId = messagesArr[k].getId();
        var messageDate = Utilities.formatDate(
          messagesArr[k].getDate(),
          Session.getScriptTimeZone(),
          "yyyy-MM-dd"
        );
        var messageFrom = messagesArr[k].getFrom();
        var messageSubject = messagesArr[k].getSubject();
        var messageBody = messagesArr[k].getBody();
        var messageAttachments = messagesArr[k].getAttachments();
        //Browser.msgBox(messageDate);

        // creating folders

        var baseFolder =
          "//" +
          mainfolder +
          "/" +
          brandfolder +
          "/" +
          hotelfolder +
          "/" +
          reporting +
          "/" +
          reportfolder +
          "/" +
          yearfolder +
          "/" +
          monthfolder +
          "/" +
          messageDate;
        var baseFolder2 =
          "//" +
          corporatefolder +
          "/" +
          reporting +
          "/" +
          reportfolder +
          "/" +
          yearfolder +
          "/" +
          monthfolder +
          "/" +
          messageDate +
          "/" +
          hotelfolder;

        // get a the system route folder (if it deosn't existing make it

        var folderOrig = getDriveFolder(baseFolder);
        var folder2 = getDriveFolder(baseFolder2);

        var messageBody2 = messageBody.substr(0, 400);
        var messageBody3 = messageBody.substr(0, 81);

        if (
          messageBody2 ===
            "The delivery of the attached reports in clear text via internet e-mail is provided at your request.&nbsp; Marriott assumes no liability for (i) the accuracy or propriety of email addresses that the properties use for distributing reports, (ii) the propriety of individuals you have allowed to access this email, or (iii) the unauthorized use, disclosure, loss or modification of these reports during " ||
          messageBody2 === "" ||
          messageBody2 ===
            "Requested report enclosed. Please see attachment." ||
          messageBody2 === "l" ||
          messageBody2 ===
            "Requested report enclosed.&nbsp; Please see attachment." ||
          messageBody2 ===
            '<div lang="EN-US" link="#0563C1" vlink="#954F72"><div><p class="MsoNormal"><u></u> <u></u></p></div></div>' ||
          messageBody2 === "NiteVision Report: Daily Account Summary" ||
          messageBody2 === "NiteVision Report: ARAging-LaQuinta.rpx" ||
          messageBody2 === "NiteVision Report: FolioAuditTrail.rpx" ||
          messageBody2 === "NiteVision Report: LedgerActivity-LaQuinta.rpx" ||
          messageBody2 === "NiteVision Report: Occupancy Summary" ||
          messageBody2 === "NiteVision Report: Sales Forecast" ||
          messageBody2 === "Have a great day!" ||
          messageBody2 === "Have a great day!!" ||
          messageBody2 === "Have a great day." ||
          messageBody2 ===
            "This email and attached report was sent to you from the Red Roof Inn Cedar Rapids<br /><br />" ||
          messageBody2 === '<div dir="ltr"><br></div>'
        ) {
        } else {
          // var folderLocation = baseFolder + "/" + messageDate;
          // var folderLocation2 = baseFolder2 + "/" + hotelfolder;
          if (messageSubject === "") {
            var messageSubject2 = "Message from " + messageFrom;
            folderOrig.createFile(messageSubject2, messageBody);
            folder2.createFile(messageSubject2, messageBody);
          } else {
            folderOrig.createFile(messageSubject, messageBody);
            folder2.createFile(messageSubject, messageBody);
          }
        }
        // Save attachments
        for (var i = 0; i < messageAttachments.length; i++) {
          Utilities.sleep(3000);
          var attachmentName = messageAttachments[i].getName();
          var attachmentContentType = messageAttachments[i].getContentType();
          var attachmentBlob = messageAttachments[i].copyBlob();
          attachmentBlob.setName(messageSubject + "-" + attachmentName);
          folderOrig.createFile(attachmentBlob);
          folder2.createFile(attachmentBlob);
          QuantityofFiles = QuantityofFiles + 1;
        }

        sheet
          .getRange("E" + (theincrementnumber + 1))
          .setValue(QuantityofFiles);
      }
      // Remove Gmail label from archived thread
      label.removeFromThread(threadsArr[j]);
      GmailApp.moveThreadToArchive(threadsArr[j]);
    }

    // Browser.msgBox("Gmail messages successfully archived to Google Drive");
  }
}

/**
 * Find all user's Gmail labels that represent mail message
 * movement requests es: moveto->xx@yyyy.com
 *
 * @return {GmailLabel[]} Array of GmailLabel objects
 */
function scanLabels() {
  // logs all of the names of your labels
  var labels = GmailApp.getUserLabels();
  var results = new Array();
  for (var i = 0; i < labels.length; i++) {
    if (labels[i].getName() == "Archive to Drive") {
      results.push(labels[i]);
    }
  }
  return results;
}

/**
 * Get all Gmail threads for the specified label
 *
 * @param {GmailLabel} label GmailLabel object to get threads for
 * @return {GmailThread[]} an array of threads marked with this label
 */
function getThreadsForLabel(label) {
  var threads = label.getThreads();
  return threads;
}

/**
 * Get all Gmail messages for the specified Gmail thread
 *
 * @param {GmailThread} thread object to get messages for
 * @return {GmailMessage[]} an array of messages contained in the specified thread
 */
function getMessagesforThread(thread) {
  var messages = thread.getMessages();
  return messages;
}

/**
 * Get methods of an object
 * @param {Object} object to scan
 * @return {Array} object's methods
 */
function getMethods(obj) {
  var result = [];
  for (var id in obj) {
    try {
      if (typeof obj[id] == "function") {
        result.push(id + ": " + obj[id].toString());
      }
    } catch (err) {
      result.push(id + ": inaccessible");
    }
  }
  return result;
}

function getDriveFolder(path) {
  var name, folder, search, fullpath;

  // Remove extra slashes and trim the path
  fullpath = path
    .replace(/^\/*|\/*$/g, "")
    .replace(/^\s*|\s*$/g, "")
    .split("/");
  Logger.log(fullpath);

  // Always start with the main Drive folder
  folder = DriveApp.getRootFolder();

  for (var subfolder in fullpath) {
    name = fullpath[subfolder];
    search = folder.getFoldersByName(name);

    // If folder does not exit, create it in the current level
    folder = search.hasNext() ? search.next() : folder.createFolder(name);
  }

  return folder;
}
