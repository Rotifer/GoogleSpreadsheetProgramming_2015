// This is code for a new chapter on object-oriented GAS that will be added in 2016,
// The code examples are not yet finalized or properly documented.

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
function test() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('cttv009_snp2gene_20160218');
  Logger.log(getRandomRowsFromSheet(sh, 10, 1));
}
