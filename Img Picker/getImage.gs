var sheetName = "New Data";
var column = "D:D";
var imageBaseUrl = "http://cdn.warmane.com/wotlk/icons/large/"
var sheetToPuImg = "items pictures"

function GenerateUrl(cell, number, url) {
  let image = SpreadsheetApp.newCellImage()
    .setSourceUrl(url)
    .setAltTextDescription("TestImage")
    .toBuilder()
    .build();
  SpreadsheetApp.getActive()
    .getSheetByName(sheetToPuImg)
    .getRange(cell + number)
    .setValue(image);
}

function onEdit() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  var columnDRange = sheet.getRange(column);
  var values = columnDRange.getValues();
  var bb = 1;
  values.forEach((cell, i) => {
    var imageUrl = cell[0];
    if (cell.length > 0 && i > 0) {
      GenerateUrl("A", i + 1, `${imageBaseUrl}${imageUrl}.jpg`);
      bb++;
    }
  });
}