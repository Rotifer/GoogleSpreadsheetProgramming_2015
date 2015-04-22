// Code Example 5.1
// Select a number of cells in a spreadsheet and 
//  then execute the following function.
// The address of the selected range, that is the
//   active range, is written to the log.
function activeRangeFromSpreadsheetApp () {
 var activeRange = 
    SpreadsheetApp.getActiveRange();
 Logger.log(activeRange.getA1Notation());
}

// Code Example 5.2
// Get the active cell and print its containing
//  sheet name and address to the log.
// Try re-running after adding a new sheet
//  and selecting a cell at random.
function activeCellFromSheet () {
 var activeSpreadsheet = 
     SpreadsheetApp.getActiveSpreadsheet(),
     activeCell = 
       activeSpreadsheet.getActiveCell(),
     activeCellSheetName = 
       activeCell.getSheet().getSheetName(),
     activeCellAddress = 
      activeCell.getA1Notation();
 Logger.log('The active cell is in sheet: ' + 
              activeCellSheetName);
 Logger.log('The active cell address is: ' + 
              activeCellAddress);
}

// Code Example 5.3
// Print Range object properties 
// (all are methods) to log.
function printRangeMethods () {
 var rng = 
   SpreadsheetApp.getActiveRange();
 Logger.log(Object.keys(rng)
   .sort().join('\n'));
}

// Code Example 5.4
// Creating a Range object using two different 
//  overloaded versions of the Sheet 
//  "getRange()" method.
// "getSheets()[0]" gets the first sheet of the 
//   array of Sheet objects returned by 
//  "getSheets()".
// Both calls to "getRange()" return a Range
// object representing the same range address
//   (A1:B10).
function getRangeObject () {
  var ss = 
    SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.getSheets()[0],
      rngByAddress = sh.getRange('A1:B10'),
      rngByRowColNums = 
        sh.getRange(1, 1, 10, 2);
  Logger.log(rngByAddress.getA1Notation());
  Logger.log(
    rngByRowColNums.getA1Notation());
}

// Code Example 5.5
// Set a number of properties for a range.
// Add a new sheet.
// Set various properties for the cell 
//  A1 of the new sheet.
function setRangeA1Properties() {
  var ss = 
   SpreadsheetApp.getActiveSpreadsheet(),
      newSheet,
     rngA1;
 newSheet = ss.insertSheet();
 newSheet.setName('RangeTest');
 rngA1 = newSheet.getRange('A1');
 rngA1.setNote(
   'Hold The date returned by spreadsheet ' 
    + ' function "TODAY()"');
 rngA1.setFormula('=TODAY()');
 rngA1.setBackground('black');
 rngA1.setFontColor('white');
 rngA1.setFontWeight('bold');
}

// Code Example 5.6
// Demonstrate get methods for 'Range' 
//  properties.
// Assumes function "setRangeA1Properties()
//   has been run.
// Prints the properties to the log.
// Demo purposes only!
function printA1PropertiesToLog () {
  var rngA1 = 
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName('RangeTest').getRange('A1');
  Logger.log(rngA1.getNote());
  Logger.log(rngA1.getFormula());
  Logger.log(rngA1.getBackground());
  Logger.log(rngA1.getFontColor());
  Logger.log(rngA1.getFontWeight());
}

// Code Example 5.7
// Starting with cell C10 of the active sheet,
// add comments to each of its adjoining cells
//  stating the row and column offsets needed
//  to reference the commented cell 
// from cell C10.
function rangeOffsetDemo () {
  var rng = 
    SpreadsheetApp.getActiveSheet()
     .getRange('C10');
  rng.setBackground('red');
  rng.setValue('Method offset()');
  rng.offset(-1,-1)
   .setNote('Offset -1, -1 from cell '
    + rng.getA1Notation());
  rng.offset(-1,0)
    .setNote('Offset -1, 0 from cell '
    + rng.getA1Notation());
  rng.offset(-1,1)
    .setNote('Offset -1, 1 from cell '
    + rng.getA1Notation());
  rng.offset(0,1)
  .setNote('Offset 0, 1 from cell '
  + rng.getA1Notation());
  rng.offset(1,0)
    .setNote('Offset 1, 0 from cell '
    + rng.getA1Notation());
  rng.offset(0,1)
    .setNote('Offset 0, 1 from cell '
    + rng.getA1Notation());
  rng.offset(1,1)
    .setNote('Offset 1, 1 from cell '
    + rng.getA1Notation());
  rng.offset(0,-1)
    .setNote('Offset 0, -1 from cell '
    + rng.getA1Notation());
  rng.offset(1,-1)
    .setNote('Offset -1, -1 from cell '
    + rng.getA1Notation());
}

// Code Example 5.8
// See the Sheet method getDataRange() in action.
// Print the range address of the used range for
//  a sheet to the log.
// Generates an error if the sheet name
// "english_premier_league" does not exist.
function getDataRange () {
  var ss = 
    SpreadsheetApp.getActiveSpreadsheet(),
      sheetName = 'english_premier_league',
      sh = ss.getSheetByName(sheetName),
      dataRange = sh.getDataRange();
  Logger.log(dataRange.getA1Notation());
}

// Code Example 5.9
// Read the entire data range of a sheet 
// into a JavaScript array.
// Uses the JavaScript Array.isArray()
//  method twice to verify that method
// getValues()returns an array-of-arrays. 
// Print the number of array elements 
// corresponding to the number of data 
//  range rows.
// Extract and print the first 10
//  elements of the array using the 
//  array slice() method.
function dataRangeToArray () {
  var ss = 
    SpreadsheetApp.getActiveSpreadsheet(),
      sheetName = 'english_premier_league',
      sh = ss.getSheetByName(sheetName),
      dataRange = sh.getDataRange(),
      dataRangeValues = dataRange.getValues();
  Logger.log(Array.isArray(dataRangeValues));
  Logger.log(Array.isArray(dataRangeValues[0]));
  Logger.log(dataRangeValues.length);
  Logger.log(dataRangeValues.slice(0, 10));
}

// Code Example 5.10
// Loop over the array returned by 
//  getRange() and a CSV-type output
// to the log using array join() method.
function loopOverArray () {
  var ss =
    SpreadsheetApp.getActiveSpreadsheet(),
      sheetName = 'english_premier_league',
      sh = ss.getSheetByName(sheetName),
      dataRange = sh.getDataRange(),
      dataRangeValues = 
        dataRange.getValues(),
      i;
  for ( i = 0;
        i < dataRangeValues.length; 
        i += 1 ) {
    Logger.log(
      dataRangeValues[i].join(','));
  }
}

// Code Example 5.11
// In production code, this function would be 
//   re-factored into smaller functions.
// Read the data range into a JavaScript array.
// Remove and store the header line using the
//    array shift() method.
// Use the array filter() method with an anonymous
//   function as a callback to implement the 
//   filtering logic.
// Determine the element count of the 
//  filter() output array.
// Add a new sheet and store a reference to it.
// Create a Range object from the new
//   Sheet object using the getRange() method.
// The four arguments given to getRange() are:
//   (1) first column, (2) first row,
//   (3) last row, and (4) last column. 
// This creates a range corresponding to 
//  range address "A1:C5".
// Write the values of the filtered array to the 
//  newly created sheet.
function writeFilteredArrayToRange () {
  var ss = 
     SpreadsheetApp.getActiveSpreadsheet(),
      sheetName = 'english_premier_league',
      sh = ss.getSheetByName(sheetName),
      dataRange = sh.getDataRange(),
      dataRangeValues = dataRange.getValues(),
      filteredArray,
      header = dataRangeValues.shift(),
      filteredArray,
      filteredArrayColCount = 3,
      filteredArrayRowCount,
      newSheet,
      outputRange;
  filteredArray = dataRangeValues.filter( 
    function (innerArray) { 
      if (innerArray[2] >= 40) {
        return innerArray;
      }});
  filteredArray.unshift(header);
  filteredArrayRowCount = filteredArray.length;
  newSheet = ss.insertSheet();
  outputRange = newSheet
                 .getRange(1, 
                           1, 
                           filteredArrayRowCount, 
                           filteredArrayColCount);
  outputRange.setValues(filteredArray);
}

// Code Example 5.12
// Add a name to a range.
function setRangeName () {
  var ss = 
    SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.getActiveSheet(),
      rng = sh.getRange('A1:B10'),
      rngName = 'MyData';
  ss.setNamedRange(rngName, rng);
}

// Code Example 5.13
// Create a range object using the 
//  getDataRange() method.
// Pass the range and a colour name 
//  to function "setAlternateRowsColor()".
// Try changing the 'color' variable to
//   something like:
//   'red', 'green', 'yellow', 'gray', etc.
function call_setAlternateRowsColor  () {
  var ss = 
    SpreadsheetApp.getActiveSpreadsheet(),
      sheetName = 'english_premier_league',
      sh = ss.getSheetByName(sheetName),
      dataRange = sh.getDataRange(),
      color = 'grey';
  setAlternateRowsColor(dataRange, color);
}
// Set every second row in a given range to
//   the given colour.
// Check for two arguments:
//   1: Range, 2: string for colour.
// Throw a type error if either argument
//  is missing or of the wrong type.
// Use the Range offset() to loop 
//   over the range rows.
// the for loop counter starts at 0.
// It is tested in each iteration with the
//   modulus operator (%).
// If i is an odd number, the if condition 
// evaluates to true and the colour 
//  change is applied.
// WARNING: If a non-existent colour is given, 
//  then the "color" is set to undefined
//  (no color). NO error is thrown!
function setAlternateRowsColor (range, 
                                color) {
  var i,
      startCell = range.getCell(1,1),
      columnCount = range.getLastColumn(),
      lastRow = range.getLastRow();
  for (i = 0; i < lastRow; i += 1) {
    if (i % 2) {
      startCell.offset(i, 0, 1, columnCount)
                 .setBackground(color);
    }
  } 
}

// Example 5.14
// Test function for 
//  "deleteLeadingTrailingSpaces()".
// Creates a Range object and passes
//   it to this function.
function call_deleteLeadingTrailingSpaces() {
  var ss = 
    SpreadsheetApp.getActiveSpreadsheet(),
      sheetName = 'english_premier_league',
      sh = ss.getSheetByName(sheetName),
      dataRange = sh.getDataRange();
  deleteLeadingTrailingSpaces(dataRange);
}
// Process each cell in the given range.
// If the cell is of type text 
//  (typeof === 'string') then
//  remove leading and trailing white space. 
//  Else ignore it.
// Code note: The Range getCell() method
//   takes two 1-based indexes 
//  (row and column).
//   This is in contrast to the offset() method. 
//  rng.getCell(0,0) will throw an error!
function deleteLeadingTrailingSpaces(range) {
  var i,
      j,
      lastCol = range.getLastColumn(),
      lastRow = range.getLastRow(),
      cell,
      cellValue;
  for (i = 1; i <= lastRow; i += 1) {
    for (j = 1; j <= lastCol; j += 1) {
      cell = range.getCell(i,j);
      cellValue = cell.getValue();
      if (typeof cellValue === 'string') {
        cellValue = cellValue.trim();
        cell.setValue(cellValue);
      }
    }
  } 
}

// Code Example 5.15
// Create a Sheet object for the active sheet.
// Pass the sheet object to 
//   "getAllDataRangeFormulas()"
// Create an array of the keys in the returned 
//  object in default "sort()".
// Loop over the array of sorted keys and 
//  extract the  values they keys map to.
// Write the output to the log.
function call_getAllDataRangeFormulas() {
  var ss = 
    SpreadsheetApp.getActiveSpreadsheet(),
      sheet = ss.getActiveSheet(),
      sheetFormulas = getAllDataRangeFormulas(sheet),
      formulaLocations = 
        Object.keys(sheetFormulas).sort(),
      formulaCount = formulaLocations.length,
      i;
  for (i = 0; i < formulaCount; i += 1) {
    Logger.log(formulaLocations[i] + 
          ' contains ' + 
          sheetFormulas[formulaLocations[i]]);
  }
}
// Take a Sheet object as an argument.
// Return an object literal where formula 
//  locations map to formulas for all formulas
//   in the input sheet data range.
// Loop through every cell in the data range.
// If a cell has a formula, 
//   store that cells location as 
//  the key and its formula as the value
//   in the object literal.
// Return the populated object literal.
function getAllDataRangeFormulas(sheet) {
  var dataRange = sheet.getDataRange(),
      i,
      j,
      lastCol = dataRange.getLastColumn(),
      lastRow = dataRange.getLastRow(),
      cell,
      cellFormula,
      formulasLocations = {},
      sheetName = sheet.getSheetName(),
      cellAddress;
  for (i = 1; i <= lastRow; i += 1) {
    for (j = 1; j <= lastCol; j += 1) {
      cell = dataRange.getCell(i,j);
      cellFormula = cell.getFormula();
      if  (cellFormula) {
        cellAddress = sheetName + '!' + 
            cell.getA1Notation();
        formulasLocations[cellAddress] =
          cellFormula;
      }
    }
  }
  return formulasLocations;
}

// Code Example 5.16
// Call copyColumns() function passing it:
//  1: The active sheet
//  2: A newly inserted sheet
//  3: An array of column indexes to copy
//     to the new sheet
// The output in the newly inserted sheet 
//  contains the columns for the indexes
//   given in the array in the 
// order specified in the array.
function run_copyColumns() {
  var ss = 
    SpreadsheetApp.getActiveSpreadsheet(),
      inputSheet = ss.getActiveSheet(),
      outputSheet = ss.insertSheet(),
      columnIndexes = [4,3,2,1];
  copyColumns(inputSheet, 
              outputSheet, 
              columnIndexes);
}
// Given an input sheet, an output sheet,
// and an array:
// Use the numeric column indexes in 
//  the array to copy those columns from 
//  the input sheet to the output sheet.
// The function checks its input arguments
//   and throws an error
// if they are not Sheet, Sheet, Array.
// The array is expected to be an array of
//  integers but it does 
//   not check the array element types
function copyColumns(inputSheet,
                     outputSheet,
                     columnIndexes) {
  var dataRangeRowCount = 
       inputSheet.getDataRange()
                 .getNumRows(),
      columnsToCopyCount = 
         columnIndexes.length,
      i,
      columnIndexesCount,
      valuesToCopy = [];
  for (i = 0; 
       i < columnsToCopyCount;
       i += 1) {
    valuesToCopy = 
      inputSheet
        .getRange(1, 
                  columnIndexes[i], 
                  dataRangeRowCount,
                  1).getValues();
    outputSheet
     .getRange(1, 
      i+1, 
      dataRangeRowCount, 
      1).setValues(valuesToCopy);
  } 
}
