function ScalGmail2() {
  
  // Gets the active sheet.
  var sheet = SpreadsheetApp.getActiveSheet();
  // Default Drive folder where to archive messages
  var allData = SpreadsheetApp.getActiveSheet()
    .getRange(1, 1, 70, 17) //startRow, startColumn, numRows, numColumns
    .getValues();

  // Set the variable
  var nhotel = allData[1][6],
      reportFolder = "Night Audits",
      yearFolder = allData[0][1],
      monthFolder = allData[0][16],
      mainFolder = "Hotels",
      corporateFolder = "Corporate",
      reporting = "Reports";

  for ( var iterator = 3; iterator <= nhotel + 2; iterator++ ) 
  {
    var gmailLabels = allData[iterator][0],
        hotelFolder = allData[iterator][1],
        brandFolder = allData[iterator][2];

    var QuantityofFiles = Number(allData[iterator][4]); // E-iterator
    //---------------------------

    // Get the gmail label
    var label = GmailApp.getUserLabelByName(gmailLabels);

    // Get the threads for the label
    var threadsArr = getThreadsForLabel(label);

    for (var j = 0; j < threadsArr.length; j++) {
      var messagesArr = getMessagesforThread(threadsArr[j]);
      for (var k = 0; k < messagesArr.length; k++) {
        //var messageId = messagesArr[k].getId();
        var messageDate = Utilities.formatDate(
          messagesArr[k].getDate(),
          Session.getScriptTimeZone(),
          "yyyy-MM-dd"
        );

        // path of the folders

        var baseFolder =
          "//" +
          mainFolder +
          "/" +
          brandFolder +
          "/" +
          hotelFolder +
          "/" +
          reporting +
          "/" +
          reportFolder +
          "/" +
          yearFolder +
          "/" +
          monthFolder +
          "/" +
          messageDate;
        var baseFolder2 =
          "//" +
          corporateFolder +
          "/" +
          reporting +
          "/" +
          reportFolder +
          "/" +
          yearFolder +
          "/" +
          monthFolder +
          "/" +
          messageDate +
          "/" +
          hotelFolder;

        // Get message items
        var messageFrom = messagesArr[k].getFrom();
        var messageSubject = messagesArr[k].getSubject();
        var messageBody = messagesArr[k].getBody();
        var messageAttachments = messagesArr[k].getAttachments();


        // creating folders
        // get a the system route folder (if it deosn't existing make it)

        var folderOrig = getDriveFolder(baseFolder);
        var folder2 = getDriveFolder(baseFolder2);

        var messageBody2 = messageBody.substr(0, 400);

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
        } else 
        {
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
          //var attachmentContentType = messageAttachments[i].getContentType();
          var attachmentBlob = messageAttachments[i].copyBlob();
          attachmentBlob.setName(messageSubject + "-" + attachmentName);
          folderOrig.createFile(attachmentBlob);
          folder2.createFile(attachmentBlob);
          QuantityofFiles = QuantityofFiles + 1;
        }

        sheet
          .getRange("E" + (iterator + 1))
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

/**♦
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
