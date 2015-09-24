/*
All GAS code for chapter 8 in book:
"Google Sheets Programming With Google Apps Script",
(https://leanpub.com/googlespreadsheetprogramming)
Michael Maguire
2015-09-24
*/

// Code Example 8.1
function listDriveFolders() {
  var it = DriveApp.getFolders(),
      folder;
  Logger.log(it);
  while(it.hasNext()) {
    folder = it.next();
    Logger.log(folder.getName());
  }
}

// Code Example 8.2
// Write the names of files in the user's
// Google Drive to the logger.
function listDriveFiles() {
  var fileIt = DriveApp.getFiles(),
      file;
  Logger.log("File Names:");
  while (fileIt.hasNext()) {
    file = fileIt.next()
    Logger.log(file.getName());
  }
}

// Code Example 8.3
// Display the name of the top-level
// Google Drive folder.
function showRoot() {
  var root = DriveApp.getRootFolder();
  Browser.msgBox(root.getName());
}

// Code Example 8.4
// Demonstration code only to show how 
// Google Drive allows duplicate file
// names and duplicate folder names
// within the same folder.
function makeDuplicateFilesAndFolders() {
  SpreadsheetApp.create('duplicate spreadsheet');
  SpreadsheetApp.create('duplicate spreadsheet');
  DriveApp.createFolder('duplicate folder');
  DriveApp.createFolder('duplicate folder');
}


// Code Example 8.5
// Demonstrates how folders can have the 
//  same name and parent folder 
//  (Root in this instance) but yet
// have different IDs.
function writeFolderNamesAndIds() {
  var folderIt = DriveApp.getFolders(),
      folder; 
  while(folderIt.hasNext()) {
    folder = folderIt.next();
    Logger.log('Folder Name: ' + 
               folder.getName() + 
               ', ' + 
               'Folder ID: ' +
               folder.getId());
  }
}


// Code Example 8.6
// Remove any folders with the given argument name.
// Caution: Make sure no required
function removeFolder(folderName) {
  var folderIt =
      DriveApp.getFoldersByName(folderName),
      folder;
  while(folderIt.hasNext()) {
    folder = folderIt.next();
    folder.setTrashed(true);
  }
}
// Call "removeFolder()" passing the name of the
// dummy folders created in example 8.4 above.
function removeTestDuplicateFolders() {
  var folderName = 'duplicate folder';
  removeFolder(folderName);
}
// Remove all files called 'duplicate spreadsheet'
// created in example 8.4 above.
function removeDummyFiles() {
  var fileIt = 
      DriveApp.getFilesByName('duplicate spreadsheet'),
      file;
  while(fileIt.hasNext()) {
    file = fileIt.next();
    file.setTrashed(true);
  }
}

// Code Example 8.7
// Create a test folder and a test file.
// Add the test file to the test folder.
function addNewFileToNewFolder() {
  var newFolder = DriveApp.createFolder('Test Folder'),
      newSpreadsheet = SpreadsheetApp.create('Test File'),
      newSpreadsheetId = newSpreadsheet.getId(),
      newFile = DriveApp.getFileById(newSpreadsheetId);
  newFolder.addFile(newFile)
}

// Code Example 8.8
// Remove the file from root folder.
// Does not delete it!
// ID was taken from the URL.
// Warning: Replace the ID below with your own
// or you will get an error.
function removeTestFileFromRootFolder() {
  var root = DriveApp.getRootFolder(),
      id = '1xIeidH1-Em-xDlm_84e70FrORfkttXFfVqYRzEv58ZA',
      file = DriveApp.getFileById(id);
  root.removeFile(file);
}

// Code Example 8.9
// Demonstrate how a file can have multiple parents.
//
// Create a folder and a spreadsheet.
// Add the newly created spredsheet to the new folder
// and return the ID of the new spreadsheet.
// Used by function "getFileParents()" below.
function createSheetAndAddToFolder() {
  var newFolder = 
      DriveApp.createFolder('TestParent'),
      gSheet = SpreadsheetApp.create('DummySheet'),
      gSheetId = gSheet.getId(),
      newFile = DriveApp.getFileById(gSheetId);
  newFolder.addFile(newFile);
  return gSheetId;
}
// Call the function above to create a folder and 
// spreadsheet and add the spreadsheet to the file.
// Get the file object usin the returned ID and
// call its "getParents()" method. Traverse the
// returned iterator to print out the parent names.
function getFileParents() {
  var gSheetId = createSheetAndAddToFolder(),
      file = DriveApp.getFileById(gSheetId), 
      gSheetParentsIt = file.getParents(),
      parentFolder;
  while(gSheetParentsIt.hasNext()) {
    parentFolder = gSheetParentsIt.next();
    Logger.log(parentFolder.getName());
  }
}

// Code Example 8.10

function createFiles() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.getActiveSheet(),
      newSsName = 'myspreadsheet',
      newSs = SpreadsheetApp.create(newSsName),
      newDocName = 'mydocument',
      newDoc = DocumentApp.create(newDocName);
  
  sh.appendRow(['File Name', 
                'File ID', 
                'File URL']);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold');
  sh.appendRow([newSsName, 
                newSs.getId(), 
                newSs.getUrl()]);
  sh.appendRow([newDocName, 
                newDoc.getId(), 
                newDoc.getUrl()]);
  ss.setNamedRange('FileDetails',
                   sh.getRange(2, 1, 2, 3));
}


// Code Example 8.11
// Adds a viewer to the document created above.
// Uses the range name to access the file ID.
// Requires authorisation.
function addViewer() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      rngName = 'FileDetails',
      inputRng = ss.getRangeByName(rngName),
      docId = inputRng.getCell(2, 2).getValue(),
      doc = DriveApp.getFileById(docId),
      newViewer =
        'michaelfmaguire@gmail.com';
  doc.addEditor(newViewer);
}

// Code Example 8.12
// Create a new folder and two files:
// a sheet and a document.
// Add them to the new folder and 
// remove them from the root folder.
function newFilesToShareInNewFolder() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      rootFolder = DriveApp.getRootFolder(),
      newFolder = 
        DriveApp.createFolder('ToShareWithMick'),
      sh = SpreadsheetApp.create('mysheet'),
      shId = sh.getId(),
      shFile =
        DriveApp.getFileById(shId),
      doc = DocumentApp.create('mydocument'),
      docId = doc.getId(),
      docFile = DriveApp.getFileById(docId);
  newFolder.addFile(shFile);
  newFolder.addFile(docFile);
  rootFolder.removeFile(shFile);
  rootFolder.removeFile(docFile);
}

// Code Example 8.13
// Add a viewer to a folder.
// Replace ID with one from a folder URL and
// the email address in the "addViewer()* method
// call.
function shareFolder() {
  var folderId = 'ID From Url',
      folder = DriveApp.getFolderById(folderId);
  folder.addViewer('emailAddress');
}

// Code Example 8.14
// Given a Folder object and an array,
// recursively search for all folders
// and return an array of Folder objects.
function getAllSubFolders(parent, folders) {
  var folderIt = parent.getFolders(),
      subFolder;
  while (folderIt.hasNext()) {
    subFolder = folderIt.next();
    getAllSubFolders(subFolder, folders);
    folders.push(subFolder);
  }
  return folders;
}
// Execute function "getAllSubFolders()" passing
// it the "root" folder.
function run_getAllSubFolders() {
  var root = DriveApp.getRootFolder(),
      folders = [],
      folders = getAllSubFolders(root, folders);
  folders.forEach(function(folder) {
                    Logger.log(folder.getName());
                  });
}


// Code example 8.15
// Given a parent Folder object and a folder name,
// call the function "getAllSubFolders()" to return
// an array of Folder objects. The array "map" methos
// is used to return the folder names as an array.
// This array is checked to determine if the folder name
// is a member.
function folderNameExists(parentFolder, 
                          subFolderName) {
  var folderNames = [],
      folders = getAllSubFolders(parentFolder, []);
    folderNames = folders.map(
                    function (folder) {
                      return folder.getName();
                    });
  if (folderNames
      .indexOf(subFolderName) > -1) {
    return true;
  } else {
    return false;
  }   
}
// Code to test "folderNameExists()"
// Change the value of folderName to existing
//  and non-existent folder names to see
// how the tested function operates!
function test_folderNameExists() {
  var folder = DriveApp.getRootFolder(),
      folderName = 'JavaScript';
  if (folderNameExists(folder, folderName)) {
    Logger.log("yup");
  } else {
    Logger.log("Nope");
  }
}


// Code example 8.16
// Filter all files based on a given cut-off date
//  and return an array of files older than the
//  cut-off date.
function getFilesOlderThan(cutoffDate) {
  var fileIt = DriveApp.getFiles(),
      filesOlderThan = [],
      file;
  while(fileIt.hasNext()) {
    file = fileIt.next();
    Logger.log(file);
    if(file.getDateCreated() < cutoffDate) {
      filesOlderThan.push(file);
    }
  }
  return filesOlderThan;
}

// Write some file details for files older than
//  a specified date to a new sheet.
function test_getFilesOlderThan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      shName = 'OldFiles',
      sh = ss.insertSheet(),
      // 'July 28, 2014', months are 0-11 in JS
      testDate = new Date(2014, 8, 23);
      oldFiles = getFilesOlderThan(testDate);
  sh.setName(shName);
  oldFiles.forEach(
    function (file) {
      sh.appendRow([file.getName(), 
                    file.getSize(), 
                    file.getDateCreated(),
                    file.getId()]);
    });
}


// Code Example 8.17
// Given a File object, return number
// of days since its creation.
function getFileAgeInDays(file) {
var today = new Date(),
      createdDate = file.getDateCreated(),
      msecPerDay = 1000 * 60 * 60 * 24,
      fileAgeInDays =
     (today - createdDate)/msecPerDay;
  return Math.round(
    fileAgeInDays).toFixed(0);
}
// Get a test file ID from the URL
// and assign it to variable fileID
// Run and check the log
function test_getFileAgeInDays() {
  var fileId = '0B2dsdq7IKB9ydXNBZXg2czJpYTg',
      file = 
        DriveApp.getFileById(fileId),
      fileAgeInDays =
       getFileAgeInDays(file);
  Logger.log('File ' + 
             file.getName() + 
             ' is ' + 
             fileAgeInDays + 
             ' days old.');
}

// Code Example 8.18
// Return an array of File objects
// for a given Folder argument.
function getFilesForFolder(folder) {
  var fileIt = folder.getFiles(),
      file,
      files = [];
  while(fileIt.hasNext()) {
    file = fileIt.next();
    files.push(file);
  }
  return files;
}
// Return an array of empty folder objects 
// in the user's Google drive.
// An empty folder is one with no associated 
//  files or sub-folders.
// NB: Requires function "getAllSubFolders()"
// from Code Example 8.14.
function getEmptyFolders() {
  var folderIt = DriveApp.getFolders(),
      folder,
      emptyFolders = [],
      files = [],
      folders = [];
  while(folderIt.hasNext()){
    folder = folderIt.next();
    files = getFilesForFolder(folder);
    folders = getAllSubFolders(folder, []);
    if(files.length === 0 && folders.length === 0) {
      emptyFolders.push(folder);
    }
  }
  return emptyFolders;
}
// Write the IDs and names of
//  all empty folders to the 
//  log.
function processEmptyFolders() {
  var emptyFolders =
      getEmptyFolders();
  emptyFolders.forEach(
    function (folder) {
      Logger.log(folder.getName() +
                 ': ' + 
                 folder.getId());
    });
}

// Code Example 8.19 - A
// Return object mapping each
//  file name to an array of file IDs.
// If file names are not duplicated,
//  the arrays they map to will
// have a single element (file ID).
function getFileNameIdMap() {
  var fileIt = DriveApp.getFiles(),
      fileNameIdMap = {},
      file,
      fileName,
      fileId;
  while(fileIt.hasNext()) {
    file = fileIt.next();   
    fileName = file.getName();
    fileId = file.getId();
    if (fileName in fileNameIdMap) {
      fileNameIdMap[fileName]
        .push(fileId);
    } else {
      fileNameIdMap[fileName] = [];
      fileNameIdMap[fileName]
        .push(fileId);
    }
  }
  return fileNameIdMap;
}
// Execute to test "getFileNameIdMap()".
function run_getFileNameIdMap() {
  var fileNameIdMap = getFileNameIdMap();
  Logger.log(fileNameIdMap);
}
// Code Example 8.19 - B
// Return an array of file IDs for files
// with duplicate names.
// Loops over the object returned by
// getFileNameIdMap() and returns
// only those file IDs for duplicate
// file names.
function getDuplicateFileNameIds() {
  var fileNameIdMap = getFileNameIdMap(),
      fileName,
      duplicateFileNameIds = [];
  for (fileName in fileNameIdMap) {
    if (fileNameIdMap[fileName].length > 1) {
      duplicateFileNameIds =
        duplicateFileNameIds
          .concat(fileNameIdMap[fileName]);
    }
  }
  return duplicateFileNameIds;
}
// Run function getDuplicateFileNameIds().
function run_getDuplicateFileNameIds() {
  var duplicateFileNameIds =
        getDuplicateFileNameIds();
  Logger.log(duplicateFileNameIds);
}
// Code Example 8.19 - C
// Add files with duplicated names to
//  a newly created folder.
function addDuplicateFilesToFolder() {
  var duplicateFileNameIds = 
      getDuplicateFileNameIds(),
      folderName = 'DUPLICATE FILES',
      folder =
       DriveApp.createFolder(folderName);
  duplicateFileNameIds.forEach(
    function (fileId) {
      var file = 
        DriveApp.getFileById(fileId);
      folder.addFile(file);
    });
}
