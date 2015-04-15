// Code Example 2.1
/**
* Simple function that cannot be called from the spreadsheet as a user-defined function
* because it sets a spreadsheet property.
*
* @param {String} rangeAddress
* @return {undefined}
*/
function setRangeFontBold (rangeAddress) {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(rangeAddress).setFontWeight('bold');
}

// Code Example 2.2
/**
* A function that demonstrates that function "setRangeFontBold()
* is valid although it cannot be called as a user-defined function.
*
* @return {undefined}
*/
function call_setCellFontBold () {
  var ui = SpreadsheetApp.getUi(),
      response = ui.prompt(
                    'Set Range Font Bold', 
                    'Provide a range address',
                    ui.ButtonSet.OK_CANCEL),
      rangeAddress = response.getResponseText();
  setRangeFontBold(rangeAddress);
}

// Code Example 2.3
/**
 * Function to demonstrate how to check
 * the number and types of passed arguments.
 * 
 * 
 * @return {undefined}
 */
function testFunc(arg1, arg2) {
  var i;
  Logger.log('Number of arguments given: ' + 
              arguments.length);
  Logger.log('Number of arguments expected: ' + 
              testFunc.length);
  for (i = 0; i< arguments.length; i += 1) {
    Logger.log('The type of argument number ' + 
               (i + 1) + ' is ' + 
                      typeof arguments[i]);
  }
}

// Code Example 2.4
/**
 * Function that calls "testFunc()"
 * twice with different argument types
 * and different argument counts.
 * 
 * @return {undefined}
 */
 function call_testFunc() {
  Logger.log('First Invocation:');
  testFunc('arg1', 2);
  Logger.log('Second Invocation:');
  testFunc('arg1', 2, 'arg3', false, 
            new Date(), null, undefined);
}

// Code Example 2.5
// Function that is expected
// to add numbers but will also
// "add" strings.
function adder(a, b) {
  return a + b;
}
// Test "adder()" with numeric arguments
// and with one numeric and one string
// argument.
function run_adder() {
  Logger.log(adder(1, 2));
  Logger.log(adder('cat', 1));
}
