// Code Example 7.1
/**
* Function "onOpen()" is executed when a Sheet is opened.
* The "onOpen()" function given here adds a menu to the menu bar
* and adds two items to the new menu. Each of these items
* execute a function when selected. The functions simply display a
* message box dialog.
*/
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      menuItems = [{name: "Item1", functionName: "function1"}, 
                   {name: "Item2", functionName: "function2"}];
  ss.addMenu("My Menu", menuItems);
}

function function1() {
  Browser.msgBox('You selected menu item 1');
}

function function2() {
  Browser.msgBox('You selected menu item 2');
}

// Code Example 7.2
/**
* Create a simple data entry form that collects
* two text values and writes them to the active sheet.
* See also "index.html" in code example ch07.html
*/
/**
* Trigger function to build a menu and add it
* to the menu bar when the spreadsheet is opened.
*/
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      menuItems = [{name: "Show Form", functionName: "showForm"}];
  ss.addMenu("Data Entry", menuItems);
}
/**
* Read in the Html from file "index.html"
* and create and display the form.
*/
function showForm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      html = HtmlService.createHtmlOutputFromFile('index')
             .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  ss.show(html);
}
/**
* Take a form as an argument and use the input element name
* attribute to reference the text values entered by the user
* and write these values to the active spreadsheet.
*
* @param {form}
* @return null
*/
function getValuesFromForm(form){
  var firstName = form.firstName,
      lastName = form.lastName,
      sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.appendRow([firstName, lastName]);
}

// GAS Code: Example 7-4
// Show the web data entry form
function showForm() {
     var ss = SpreadsheetApp.getActiveSpreadsheet(),
         html = HtmlService.createHtmlOutputFromFile('index')
                .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  html.setWidth(500);
  html.setHeight(250);
  ss.show(html);
}
// Get values from submitted HTML form
function getValuesFromForm(form){
  var firstName = form.firstName,
      lastName = form.lastName,
      userAddress = form.userAddress,
      zipCode = form.zipCode,
      userPhone = form.userPhone,
      chosenDate = form.chosenDate,
      userEmail = form.userEmail,
      sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.appendRow([firstName, lastName, 
                   userAddress, zipCode, 
                   userPhone, chosenDate, userEmail]);
}

// Code Example 7-5
// Add some dummy data to the active sheet.
// Data from "http://www.w3schools.com/html/html_tables.asp"
function addRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      sheet = ss.getActiveSheet(),
      rows = [['Number', 'First Name', 'Last Name', 'Points'],
             [1, 'Eve', 'Jackson', 94],
             [2, 'John', 'Doe', 80],
             [3, 'Adam', 'Johnson', 67],
             [4, 'Jill', 'Smith', 50]],
      rng,
      rngName = 'Input';
  rows.forEach(function(row) {
    sheet.appendRow(row);
  });
  rng = sheet.getDataRange();
  ss.setNamedRange(rngName, rng);
}

// Take the template file and fill in the data
// in range named "Input".
// Display the data as a sidebar.
function displayDataAsSidebar() {
  var html = HtmlService.createTemplateFromFile('DummyData'); 
  SpreadsheetApp.getUi() 
      .showSidebar(html.evaluate());
}
// Return the data in the range named "Input"
// as an array-of arrays.
function getData(){
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      rng = ss.getRangeByName('Input'),
  data = rng.getValues(); 
  return data;
}
