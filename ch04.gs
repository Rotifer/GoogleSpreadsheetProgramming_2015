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
