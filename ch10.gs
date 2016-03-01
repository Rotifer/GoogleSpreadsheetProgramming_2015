// CHAPTER NOT YET RELEASED!


// ####### Code Example 10.1 #######

// The first version of a GAS application to select a random
// sample of rows from a sheet and write them to a new sheet.
// Full explanation is given in chapter 10 of the book.
/*
Given an array of values, return a sample.
The size of the sample to return is the howMany"
argument.
This function uses the same trick discussed in
chapter 3, code example 3.13.
*/
function getRandomArrayElements(array, howMany) {
  var sampleElements = [],
      elementCount = array.length,
      idx,
      sampleElement;
  for(var i = 1; i <= howMany; i += 1) {
    idx = 1 + Math.floor(Math.random() * elementCount);
    sampleElement = array[idx];
    sampleElements.push(sampleElement);
    delete array[idx];
    elementCount = array.length;
  }
  return sampleElements;
}

/*
Given a Sheet object, a count of the row sample size and the header row count,
return an object containing two arrays:
1: The header row or rows.
2: The sampled data rows.
*/
function getRandomRowsFromSheet(sheetObj, sampleRowCount, headerRowCount) {
  var rows = sheetObj.getDataRange().getValues(),
      headerRows = [],
      sampleRows = [],
      results = {};
  headerRows = headerRowCount > 0 ? rows.splice(0, headerRowCount) : [];
  sampleRows = getRandomArrayElements(rows, 10);
  results['headerRows'] = headerRows;
  results['sampleRows'] = sampleRows;
  return results;
}
/*
Adds a new sheet to the active spreadsheet and assigns it the name
specifies by the argument "newSheetName".
It then appends the header and sampled rows to the new sheet.
*/
function writeSampleRowsToNewSheet(newSheetName, headerRows, sampleRows) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().insertSheet(newSheetName),
      i;
  for(i = 0; i < headerRows.length; i += 1) {
    sh.appendRow(headerRows[i]);
  }
  for(i = 0; i < sampleRows.length; i += 1) {
    sh.appendRow(sampleRows[i]);
  }
}
/*
Execute this function to get a sample of 10 rows from a spreadsheet
named "countries".
The data used was generated using the spreadsheet function call:
=IMPORTHTML("http://www.w3schools.com/tags/ref_country_codes.asp", "table", 1)
The sheet was then renamed "countries".
This function uses a date stamp to ensure that each new sheet gets a unique name.
See book text for details!
*/
function runMain() {
  var sheetName = 'countries',
      sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName),
      sampleRowCount = 10,
      headerRowCount = 1,
      rows = getRandomRowsFromSheet(sh, sampleRowCount, headerRowCount),
      headerRows = rows['headerRows'],
      sampleRows = rows['sampleRows'],
      sec = Math.floor(new Date().getTime() / 1000),
      newSheetName = 'sample_' + sec.toString();
  writeSampleRowsToNewSheet(newSheetName, headerRows, sampleRows);
}
