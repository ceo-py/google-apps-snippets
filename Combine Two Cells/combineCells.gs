var sheetName = "New Data";
var itemNameColumn = "J:J";
var itemNameColumnToChange = "J";
var itemLvlColumn = "A:A";

function changeName(combinedData, sheetName, cell, i) {
  SpreadsheetApp.getActive()
    .getSheetByName(sheetName)
    .getRange(cell + (i + 1))
    .setValue(combinedData);
}

function getColumnData(sheet, column) {
  var columnDRange = sheet.getRange(column);
  return columnDRange.getValues();
}

function addItemLvlToName() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  var names = getColumnData(sheet, itemNameColumn);
  var itemLvls = getColumnData(sheet, itemLvlColumn);

  names.forEach((x, i) => {
    let name = x[0];
    let iLvl = itemLvls[i];
    if (name.length > 0 && i > 0) {
      changeName(`${name}  (${iLvl})`, sheetName, itemNameColumnToChange, i);
    }
  });
}