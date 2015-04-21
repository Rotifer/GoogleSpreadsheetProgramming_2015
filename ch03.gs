// Code Example 3.1
/**
* Simple function that cannot be called from
* the spreadsheet as a user-defined function
* because it sets a spreadsheet property.
*
* @param {String} rangeAddress
* @return {undefined}
*/
function setRangeFontBold (rangeAddress) {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(rangeAddress).setFontWeight('bold');
}

// Code Example 3.2
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

// Code Example 3.3
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
  testFunc('arg1', 2, 'arg3', false, new Date(), null, undefined);
}

// Code Example 3.4
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

// Code Example 3.5
// Function that checks that
// both arguments are of type number.
// Throws an error if this is not true.
function adder(a, b) {
  if (!(typeof a === 'number' && 
        typeof b === 'number')) {
    throw TypeError(
          'TypeError: Both arguments must be numeric!');
  }
  return a + b;
}

// Test "adder()" with numeric arguments
// Thrown error is caught, see logger.
function run_adder() {
  Logger.log(adder(1, 2));
  try {
    Logger.log(adder('cat', 1));
  } catch (error) {
    Logger.log(error.message);
  }
}

// Code Example 3.6
/*
RSD Data (Paste this into a Google Sheet to test):
19.81
18.29
21.47
22.54
20.17
20.1
17.61
20.91
21.62
19.17
*/
/**
 * Given the man and standard deviation
 * return the relative standard deviation.
 * 
 * @param {number} stdev 
 * @param {number} mean
 * @return {number}
 * @customfunction
*/
function RSD (stdev, mean) {
  return 100 * (stdev/mean);
}

// Code Example 3.7
/**
 * Given a temperature in Celsius, return Fahrenheit value.
 * @param {number} celsius
 * @return {number}
 * @customfunction
 */
function CELSIUSTOFAHRENHEIT(celsius) {
  if (typeof celsius !== 'number') {
    throw TypeError('Celsius value must be a number');
  }
  return ((celsius * 9) / 5) + 32;
}

// Code Example 3.8
/**
 * Given the radius, return the area of the circle.
 * @param {number} radius
 * @return {number}
 * @customfunction
 */
function AREAOFCIRCLE (radius)  {
  if (typeof radius !== 'number' || radius < 0){
    throw Error('Radius must be a positive number');
  }
  return Math.PI * (radius * radius);
}

// Code Example 3.9
/**
 * Given a date, return the name of the day for that date.
 *
 * @param {Date} date
 * @return {String}
 * @customfunction
 */
function DAYNAME(date) {
  var dayNumberNameMap = {
    0: 'Sunday',
    1: 'Monday',
    2: 'Tuesday',
    3: 'Wednesday',
    4: 'Thursday',
    5: 'Friday',
    6: 'Saturday'},
      dayName,
      dayNumber;
  if(! date.getDay ) {
    throw TypeError('TypeError: Argument is not of type "Date"!');
  }
  dayNumber = date.getDay();
  dayName = dayNumberNameMap[dayNumber];
  return dayName;
}


// Code Example 3.10
/**
* Add a given number of days to the given date
* and return the new date.
*
* @param {Date} date
* @param {number} days
* @return {Date}
*/
function addDays(date, days) {
  // Taken from Stackoverflow:
  // http://stackoverflow.com/questions/563406/add-days-to-datetime
    var result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
}

/**
* Given a start date, an end date and a day name.
* return an array of all dates that fall (inclusive)
* between those two dates for the given day name.
*
* @param {Date} startDate
* @param {Date} endDate
* @param {String} dayName
* @return {Date[]}
* @customfunction
*/
function DATESOFDAY(startDate, endDate, dayName) {
  var dayNameDates = [],
      testDate = startDate,
      testDayName;
  while(testDate <= endDate) {
    testDayName = DAYNAME(testDate);
    if(testDayName.toLowerCase() === dayName.toLowerCase()) {
      dayNameDates.push(testDate);
    }
    testDate = addDays(testDate, 1);
  }
  return dayNameDates;
}

// Code Example 3.11
/**
 * Print all string properties to the 
 *  Script Editor Logger
 * @return {undefined}
 */
function printStringMethods() {
  var strMethods = 
    Object.getOwnPropertyNames(String.prototype);
  Logger.log('String has ' +
              strMethods.length +
             ' properties.');
  Logger.log(strMethods.sort().join('\n'));
}

// Code Example 3.12
/**
 * Given a string, return a new string
 * with the characters in reverse order.
 * 
 * @param {String} str
 * @return {String}
 * @customfunction
 */
function REVERSESTRING(str) {
  var strReversed = '',
      lastCharIndex = str.length - 1,
      i;
  for (i = lastCharIndex; i >= 0; i -= 1) {
    strReversed += str[i];
  }
  return strReversed;
}

// Code Example 3.13
/**
 * Simulate a throw of a die
 * by returning a number between
 * and 6.
 * 
 * @return {number}
 * @customfunction
 */
function THROWDIE() {
  return 1 + Math.floor(Math.random() * 6);
}

// Code Example 3.14
/** Concatenate cell values from
* an input range.
* Single quotes around concatenated 
* elements are optional.
* 
* @param {String[]} inputFromRng
* @param {String} concatStr
* @param {Boolean} addSingleQuotes
* @return {String}
* @customfunction
*/
function CONCATRANGE(inputFromRng, concatStr,
                    addSingleQuotes) {
  var cellValues;
  if (addSingleQuotes) {
    cellValues = 
      inputFromRng.map(
        function (element) {
          return "'" + element + "'";
        });
    return cellValues.join(concatStr);
 }
   return inputFromRng.join(concatStr);
}

// Code Example 3.15
/**
 * Return the ID of the active
 *  spreadsheet.
 * 
 * @return {String}
 * @customfunction
 */
function GETSPREADSHEETID() {
  return SpreadsheetApp
    .getActiveSpreadsheet().getId();
}
/**
 Return the URL of the active
 *  spreadsheet.
 * 
 * @return {String}
 * @customfunction
 */
function GETSPREADSHEETURL() {
  return SpreadsheetApp
    .getActiveSpreadsheet().getUrl();
}
/**
  Return the owner of the active
 *  spreadsheet.
 * 
 * @return {String}
 * @customfunction
 */
function GETSPREADSHEETOWNER() {
  return SpreadsheetApp
    .getActiveSpreadsheet().getOwner();
}
/**
 Return the viewers of the active
 *  spreadsheet.
 * 
 * @return {String}
 * @customfunction
 */
function GETSPREADSHEETVIEWERS() {
  var ss = 
    SpreadsheetApp.getActiveSpreadsheet();
  return ss.getViewers().join(', ');
}
/**
 Return the locale of the active
 *  spreadsheet.
 * 
 * @return {String}
 * @customfunction
 */
function GETSPREADSHEETLOCALE() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSpreadsheetLocale();
}

// Code Example 3.16
/**
 * Return French version
 * of English input.
 * 
 * @param {String} input
 * @return {String}
 * @customfunction
 */
function ENGLISHTOFRENCH(input) {
  return LanguageApp
    .translate(input, 'en', 'fr');
}
