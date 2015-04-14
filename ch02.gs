function sayHelloAlert() {
  // Declare a string literal variable.
  var greeting = 'Hello world!',
      ui = SpreadsheetApp.getUi();
  // Display a message dialog with the greeting 
  //(visible from the containing spreadsheet).
  // Older versions of Sheets used the Browser.msgBox()
  ui.alert(greeting);
}

function helloDocument() {
  var greeting = 'Hello world!';
  // Create DocumentApp instance.
  var doc = 
    DocumentApp.create('test_DocumentApp');
  // Write the greeting to a Google document.
  doc.setText(greeting);
  // Close the newly created document
  doc.saveAndClose();  
}

function helloLogger() {
  var greeting = 'Hello world!';
  //Write the greeting to a logging window.
  // This is visible from the script editor
  //   window menu "View->Logs...".
  Logger.log(greeting);  
}


function helloSpreadsheet() {
  var greeting = 'Hello world!',
      sheet = SpreadsheetApp.getActiveSheet();
  // Post the greeting variable value to cell A1
  // of the active sheet in the containing 
  //  spreadsheet.
  sheet.getRange('A1').setValue(greeting);
  // Using the LanguageApp write the 
  //  greeting to cell:
  // A2 in Spanish, 
  //  cell A3 in German, 
  //  and cell A4 in French.
  sheet.getRange('A2')
        .setValue(LanguageApp.translate(
                  greeting, 'en', 'es'));
  sheet.getRange('A3')
        .setValue(LanguageApp.translate(
                  greeting, 'en', 'de'));
  sheet.getRange('A4')
         .setValue(LanguageApp.translate(
                   greeting, 'en', 'fr'));
}
