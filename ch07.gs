// code Example 7.1
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
