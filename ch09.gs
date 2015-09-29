/*
All GAS code for chapter 9 in book:
"Google Sheets Programming With Google Apps Script",
(https://leanpub.com/googlespreadsheetprogramming)
Michael Maguire
2015-09-29
*/

// Code Example 9.1
// Use a named range (name = "EmailContacts")
// to return an array of email addresses.
// All the code techniques used here have been
// covered in earlier examples.
// To use: Create a list of email addresses
// in a single column, give that list the range
// name "EmailContacts". Execute function
// "run_getEmailList()" to ensure it is working.
function getEmailList() {
  var ss =
      SpreadsheetApp.getActiveSpreadsheet(),
      rngName = 'EmailContacts',
      emailRng = ss.getRangeByName(rngName),
      rngValues = emailRng.getValues(),
      emails = [];
  emails = rngValues.map(
    function (row) {
      return row[0];
    });
  return emails;
}
// Check that function "getEmailList()"
//  is working.
function run_getEmailList() {
  Browser.msgBox(getEmailList().join(','));
}


// Code Example 9.2
// Need to autorize!
// Get an array of email addresses using function
// "getEmailList()" in code example 9.1 and send
// a test email to them.
function sendEmail (){
  var emailList = getEmailList(),
      subject = 'Testing MailApp',
      body = 'Hi all,\n' +
        'Just checking that our email distribution list is working!\n' +
          'Cheers\n' +
            'Mick';
      MailApp.sendEmail(emailList.join(','),
                        subject,
                        body);
}
// Display remaining daily quota.
function showDailyQuota() {
  Browser.msgBox(MailApp.getRemainingDailyQuota());
}

// Code Example 9.3
//  Uses function "getEmailList()* from 
// code example 9.1.
// Creates a Google document and converts it to
// a PDF that it then sends to the email list
// as an attachment.
function sendAttachment() {
  var doc
      = DocumentApp.create('ToSendAsAttachment'),
      emailList = getEmailList(),
      file,
      pdf,
      pdfName = 'Test.pdf',
      fileId = doc.getId(),
      subject = 'Test attachment',
      body = 'See attached PDF',
      attachment,
      paraText = 
      'This is text that will be written\n' +
      'to a document that will then be saved\n' +
      'as a PDF and sent as an attachment. ';
  doc.appendParagraph(paraText);
  doc.saveAndClose();
  file = DriveApp.getFileById(fileId);
  pdf = file.getAs('application/pdf').getBytes();
  attachment = {fileName: pdfName, 
                content:pdf, 
                mimeType:'application/pdf'};
  MailApp.sendEmail(emailList.join(','), 
                    subject,
                    body,
                   {attachments:[attachment]});
}


// Code Example 9.4
// Return a list of objects that contain
//  selected details of the user's email
//  threads.
// This will be slow for a large mail box
//  so there is an optional argument to limit
//  the returned array to a certain number.
// The most recent threads are returned first.
function getThreadSummaries(threadCount) {
  var threads = 
        GmailApp.getInboxThreads(),
      threadSummaries = [],
      threadSummary = {},
      recCount = 0;
  threads.forEach(
    function (thread) {
      recCount += 1;
      if (recCount > threadCount) {
        return;
      }
      threadSummary = {};
      threadSummary['MessageCount'] =
        thread.getMessageCount();
      threadSummary['Subject'] =
        thread.getFirstMessageSubject();
      threadSummary['ThreadId'] =
        thread.getId();
      threadSummary['LastUpdate'] =
        thread.getLastMessageDate();
      threadSummary['URL'] = 
        thread.getPermalink();
      threadSummaries
         .push(threadSummary);
    });
  return threadSummaries;
}
// Run "getThreadSummaries()" for 50 threads.
// Check the log after running.
function run_getThreadSummaries() {
  Logger.log(getThreadSummaries(50));
}

// Code Example 9.5
// Write some thread details returned by
//  function "getThreadSummaries()" to
//  a newly inserted sheet named
//  "EmailThreadSummary"
// Make sure the sheet is deleted before
//  re-running or an error will be thrown
//  when it attempts to insert a new sheet
// with the same name.
function writeThreadSummary() {
  var threadCount = 10,
      threadSummaries = 
      getThreadSummaries(threadCount),
      ss = 
      SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.insertSheet(),
      headerRow;
  sh.setName('EmailThreadSummary');
  sh.appendRow(['Subject',
                'MessageCount',
                'LastUpdate',
                'ThreadId',
                'URL']);
  headerRow = sh.getRange(1,1,1,5);
  headerRow.setFontWeight('bold');
  threadSummaries.forEach(
    function (threadSummary) {
      sh.appendRow([threadSummary['Subject'],
                   threadSummary['MessageCount'],
                   threadSummary['LastUpdate'],
                   threadSummary['ThreadId'],
                   threadSummary['URL']]);
    });
}


// Code Example 9.6 - A
// Create an array of objects containing
// information extracted from each GmailMessage
// object for the given a thread ID.
// Each object property is populated by
// the return value of a GmailMessage
//  method.
// Return 'undefined' if there are no
// messages.
function getMsgsForThread(threadId) {
  var thread = 
      GmailApp.getThreadById(threadId),
      messages,
      msgSummaries = [],
      messageSummary = {};
  if (!thread) {
    return;
  }
  messages = thread.getMessages();
  messages.forEach(
    function (message) {
      messageSummary = {};
      messageSummary['From'] =
        message.getFrom();
      messageSummary['Cced'] =
        message.getCc();
      messageSummary['Date'] =
        message.getDate();
      messageSummary['Body'] =
        message.getPlainBody();
      messageSummary['MsgId'] =
        message.getId();
      msgSummaries.push(messageSummary);
    });
  return msgSummaries;
}
// Code Example 9.6 - B
// Given a message summary object as returned
//  in an array by "getMsgsForThread" and
//  a Google Document ID, open the document,
// extract the object values and write them
// to the document.
// Argument checks could/should be added to
// ensure that the first argument is an
// object and that the second is a valid
// ID for a document.
function writeMsgBodyToDoc(msgSummsForThread,
                           docId) {
  var doc = DocumentApp.openById(docId),
      header,
      from = 
      msgSummsForThread['From'],
      msgDate =
      msgSummsForThread['Date'],
      msgBody = 
      msgSummsForThread['Body'],
      msgId = 
      msgSummsForThread['MsgId'],
      docBody = doc.getBody(),
      paraTitle;
  docBody.appendParagraph('From: ' +
                          from + 
                          '\rDate: ' +
                          msgDate + 
                         '\rMessage ID: ' +
                         msgId);
  docBody.appendParagraph(msgBody);
  docBody
  .appendParagraph('   ####################   ');
  doc.saveAndClose();
}
// Code Example 9.6 -  C
// Use the Thread IDs generated earlier
//  to extract the message for each thread
// Create a new Google Document
// Write summary message data to the Google
//  Document.
// Contains a nested forEach structure,
//  see explanation in text.
function writeMsgsForThreads() {
  var ss =
      SpreadsheetApp.getActiveSpreadsheet(),
      shName = 'EmailThreadSummary',
      sh = ss.getSheetByName(shName),
      docName = 'ThreadMessages',
      doc =
      DocumentApp.create(docName),
      docId = doc.getId(),
      threadIdCount =
      sh.getLastRow() -1,
      rngThreadIds = 
        sh.getRange(2, 
                    4, 
                    threadIdCount,
                    1),
      threadIds = 
        rngThreadIds.getValues();
  threadIds.forEach(
      function(row) {
        var threadId
        = row[0],
        msgsForThread = 
        getMsgsForThread(threadId);
        msgsForThread.forEach(
          function (msg) {
            writeMsgBodyToDoc(
            msg,
            docId);
          });
      });
  
}
//Code Example 9.7
// Add a label text as given in the first argument
// to all email threads where
//  the first message in the thread has
// subject text that matches the first 
// argument in its subject section.
// To do this:
//  1: Create a label object using the
//     GmailApp createLabel() method.
//  2: Retrieve all the inbox threads.
//  3: Filter the resulting array of
//     threads based on the specified
//     subject text.
//  4: Add the label to each of the
//     threads in the filtered array
//  5: Call the thread refresh method
//     to display the label.
function labelAdd(subjectText, labelText) {
  var label = 
      GmailApp.createLabel(labelText),
      threadsAll = 
      GmailApp.getInboxThreads(),
      threadsToLabel;
  threadsToLabel = threadsAll.filter(
    function (thread) {
      return (thread.getFirstMessageSubject()
              === 
              subjectText);
    });
  threadsToLabel.forEach(
    function (thread) {
      thread.addLabel(label);
      thread.refresh();
    });
}

// Code Example 9.8
// Using a hard-code thread ID taken from
// the email URL, extract the first message
// from the thread and extract the first
// attachment from that message.
// Copy the attachment into a Blob.
// Use the DriveApp object to create a File
// object from this blob.
function putAttachmentInGoogleDrive() {
  var threadId = '150195c16476de2a',
      thread = 
      GmailApp.getThreadById(threadId),
      firstMsg = 
      thread.getMessages()[0],
      firstAttachment =
      firstMsg.getAttachments()[0],
      blob = 
      firstAttachment.copyBlob();
  DriveApp.createFile(blob);
  Logger.log('File named ' + 
             firstAttachment.getName() +
             ' has been saved to Google Drive');
}

// Code example 9.9
// Write some basic information for all
//  calendars available to the user to
//  the log.
function calendarSummary() {
  var cals = 
      CalendarApp.getAllOwnedCalendars();
  Logger.log('Number of Calendars: '
             + cals.length);
  cals.forEach(
    function(cal) {
      Logger.log('Calendar Name: ' + 
                 cal.getName());
      Logger.log('Is primary calendar? ' +
                 cal.isMyPrimaryCalendar());
      Logger.log('Calendar description: ' +
                 cal.getDescription());
                 
    });
}

// Code Example 9.10
// Take a range of dates from 29th September 
// 2015 to 7th October 2015 (inclusive) from a 
//  spreadsheet and create calendar 
//  all-day events for each of these dates.
// The title is set as "Holidays" and the
//  description to "Forget about work".
function calAddEvents() {
  var cal = CalendarApp.getDefaultCalendar(),
      ss = 
      SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.getSheetByName('Holidays'),
      holDates = 
      sh.getRange('A1:A9').getValues().map(
        function (row) {
          return row[0];
        });
  holDates.forEach(
    function (holDate) {
      var calEvent;
        calEvent  = 
        cal.createAllDayEvent('Holiday',
                              holDate);
        calEvent.setDescription(
                     'Forget about work')
    });
}

// Code Example 9.11
// Remove all events from the default calendar
//  for the dates between 29th September 
// 2015 to 7th October 2015.
//  inclusive where the event title equals
//  "Holiday".
// First get all events for the date range
//  as an array.
// Then filter this array to
//  get only those with the indicated title.
// Remove those filtered events from the
//  calendar.
function calRemoveEvents() {
  var cal = CalendarApp.getDefaultCalendar(),
      calEvents,
      eventTitleToCancel = 'Holiday',
      toCancelEvents = [];
  calEvents = 
    cal.getEvents(new Date('September 29, 2015'),
                  new Date('October 7, 2015'));
  toCancelEvents = calEvents.filter(
    function (calEvent) {
      return (calEvent.getTitle() 
              ===
             eventTitleToCancel);
    });
  toCancelEvents.forEach(
    function (eventToCancel) {
      eventToCancel.deleteEvent();
    });
}

// Code Example 9.12
// Get the default calendar object.
// Insert a new sheet into the active
//  spreadsheet.
// Retrieve all CalendarEvent objects
//  for a specified date interval
//  (month of September 2015).
// Append a header row to the new sheet
//  and populate it.
// Make the header row font bold.
// Loop over the array of CalendarEvent objects
//  and use their methods to extract property
//  values of interest.
// Write the property values to the new sheet.
function writeCalEventsToSheet() {
  var cal = CalendarApp.getDefaultCalendar(),
      ss = 
      SpreadsheetApp.getActiveSpreadsheet(),
      newShName = 'CalendarEvents',
      newSh = ss.insertSheet(newShName),
      startDate = new Date('September 01, 2015'),
      endDate = new Date('October 01, 2015'),
      calEvents = cal.getEvents(startDate, endDate),
      colHeaders = [],
      colCount,
      colHeaderRng;
  colHeaders = ['Title',
               'Description',
               'EventId',
               'DateCreated',
               'IsAllDay',
               'MyStatus',
               'StartTime',
               'EndTime',
               'Location'];
  colCount = colHeaders.length;
  newSh.appendRow(colHeaders);
  colHeaderRng = newSh.getRange(1, 
                 1, 
                 1, 
                 colCount);
  colHeaderRng.setFontWeight('bold');
  calEvents.forEach(
    function (calEvent) {
      newSh.appendRow(
       [calEvent.getTitle(),
       calEvent.getDescription(),
       calEvent.getId(),
       calEvent.getDateCreated(),
       calEvent.isAllDayEvent(),
       calEvent.getMyStatus(),
       calEvent.getStartTime(),
       calEvent.getEndTime(),
       calEvent.getLocation()
       ]);
    });
}

