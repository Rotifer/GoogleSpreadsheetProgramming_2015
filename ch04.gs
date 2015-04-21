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
