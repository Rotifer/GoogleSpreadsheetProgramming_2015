// Code Example 4.1
// Function to demonstrate the Spreadsheet 
//  object hierarchy.
// All the variables are gathered in a 
//  JavaScript array.
// At each iteration of the for loop the 
//  "toString()" method 
//  is called for each variable and its
//   output is printed to the log.
function showGoogleSpreadsheetHierarchy() {
 var ss = SpreadsheetApp.getActiveSpreadsheet(),
     sh = ss.getActiveSheet(),
     rng = ss.getRange('A1:C10'),
     innerRng = rng.getCell(3, 3),
     innerRngAddress = innerRng.getA1Notation(),
     column = innerRngAddress.slice(0,1),
     googleObjects = [ss, sh, rng, innerRng, 
                       innerRngAddress, column],
     i;
 for (i = 0; i < googleObjects.length; i += 1) {
   Logger.log(googleObjects[i].toString());
 }
}

// Code Example 4.2
// Print the column letter of the third row and 
//   third column of the range "A1:C10"
//  of the active sheet in the active 
//   spreadsheet.
// This is for demonstration purposes only!
function getColumnLetter () {
  Logger.log(
    SpreadsheetApp.getActiveSpreadsheet()
      .getActiveSheet().getRange('A1:C10')
      .getCell(3, 3).getA1Notation()
        .slice(0,1));
}

// Code Example 4.3
// Extract an array of all the property names 
//  defined for Spreadsheet and write them to
//  column A of the active sheet in the active
//   spreadsheet.
function testSpreadsheet () {
  var ss = 
     SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.getActiveSheet(),
      i,
      spreadsheetProperties = [],
      outputRngStart = sh.getRange('A2');
  sh.getRange('A1')
    .setValue('spreadsheet_properties');
  sh.getRange('A1')
    .setFontWeight('bold');
  spreadsheetProperties = 
    Object.keys(ss).sort();
  for (i = 0; 
       i < spreadsheetProperties.length;
       i += 1) {
    outputRngStart.offset(i, 0)
       .setValue(spreadsheetProperties[i]);
  }
}

// Code Example 4.4
//  Extract, an array of properties from a
//   Sheet object.
// Sort the array alphabetically using the
//  Array sort() method.
// Use the Array join() method to a create
//   a string of all the Sheet properties
//   separated by a new line.
function printSheetProperties () {
  var ss = 
    SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.getActiveSheet();
  Logger.log(Object.keys(sh)
             .sort().join('\n'));
} 

// Code Example 4.5
// Call function listSheets() passing it the 
//  Spreadsheet object for the active 
//    spreadsheet.
// The try - catch construct handles the 
//  error thrown by listSheets() if the given
// argument is absent or something
//    other than a Spreadsheet object.
function test_listSheets () {
  var ss = 
      SpreadsheetApp.getActiveSpreadsheet();
  try {
    listSheets(ss);
  } catch (error) {
    Logger.log(error.message);
  }
}
// Given a Spreadsheet object, 
//  print the names of its sheets
//   to the logger.
// Throw an error if the argument
//   is missing or if it is not
//  of type Spreadsheet.
function listSheets (spreadsheet) {
  var sheets,
      i;
  if (spreadsheet.toString()
      !== 'Spreadsheet') {
    throw TypeError(
      'Function "listSheets" expects ' +
      'argument of type "Spreadsheet"');
  }
  sheets = spreadsheet.getSheets();
  for (i = 0; 
       i < sheets.length; i += 1) {
    Logger.log(sheets[i].getName());
  }
}

// Code Example 4.6
// Create a Spreadsheet object and call 
//  "sheetExists()" for an array of sheet 
//  names to see if they exist in 
//  the given spreadsheet.
// Print the output to the log.
function test_sheetExists () {
  var ss = 
      SpreadsheetApp.getActiveSpreadsheet(),
      sheetNames = ['Sheet1', 
                    'sheet1', 
                    'Sheet2',
                    'SheetX'],
      i;
  for (i = 0; 
       i < sheetNames.length; 
       i +=1) {
    Logger.log('Sheet Name ' + 
                sheetNames[i] + 
               ' exists: ' + 
                sheetExists(ss, 
                      sheetNames[i]));
  }
}
// Given a Spreadsheet object and a sheet name, 
//  return "true" if the given sheet name exists
//  in the given Spreadsheet, 
//  else return "false".
function sheetExists(spreadsheet, sheetName) {
  if (spreadsheet.getSheetByName(sheetName)) {
    return true;
  } else {
    return false;
  }
}

// Code Example 4.7
// Copy the first sheet of the active
//  spreadsheet to a newly created 
//  spreadsheet.
function copySheetToSpreadsheet () {
  var ssSource = 
     SpreadsheetApp.getActiveSpreadsheet(),
      ssTarget = 
      SpreadsheetApp.create(
        'CopySheetTest'),
      sourceSpreadsheetName =
        ssSource.getName(),
      targetSpreadsheetName = 
        ssTarget.getName();
  Logger.log(
     'Copying the first sheet from ' + 
            sourceSpreadsheetName + 
            ' to ' + targetSpreadsheetName);
  // [0] extracts the first Sheet object 
  //   from the array created by
  //   method call "getSheets()"
  ssSource.getSheets()[0].copyTo(ssTarget);
}

// Code Example 4.8
// Create a Sheet object and pass it 
// as an argument to getSheetSummary().
// Print the return value to the log.
function test_getSheetSummary () {
  var sheet = SpreadsheetApp
             .getActiveSpreadsheet()
             .getSheets()[0];
  Logger.log(getSheetSummary(sheet));
}
// Given a Sheet object as an argument, 
//  use Sheet methods to extract 
//  information about it.
// Collect this information into an object
// literal and return the object literal.
function getSheetSummary (sheet) {
  var sheetReport = {};
  sheetReport['Sheet Name'] = 
     sheet.getName();
  sheetReport['Used Row Count'] =
     sheet.getLastRow();
  sheetReport['Used Column count'] = 
    sheet.getLastColumn();
  sheetReport['Used Range Address'] = 
   sheet.getDataRange().getA1Notation();
  return sheetReport;
}

