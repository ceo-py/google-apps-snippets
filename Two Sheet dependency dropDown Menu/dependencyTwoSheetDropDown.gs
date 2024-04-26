var talentsSheet = "Talents";
var characterSettingsSheet = "CharacterSettings";
var activeRow = [31, 39];
var activeCol = [5, 7, 21];
var autoSpecCells = ["AL5", "AL6"];
var talents = {
  holyShock: {
    talent: "G31",
    dropDownLocation: "P41",
    dropDownOptions: 10,
    glyph: {
      talent: "U34",
      dropDownOptions: 12,
    },
  },
  divineIllumination: {
    talent: "E39",
    dropDownLocation: "AD54",
    dropDownOptions: 6,
  },
  glyphOfHolyLight: {
    talent: "U31",
    dropDownLocation: "AD52",
    dropDownOptions: 5,
  },
};

function talentDynamicDropDownMenus(e) {
  const spreadSheet = e.source;
  const activeRange = e.source.getActiveRange();
  const sheetName = spreadSheet.getActiveSheet().getName();

  if (sheetName !== talentsSheet) return;
  const activeSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const dropDownSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    characterSettingsSheet
  );
  // resetDropDownMenus(dropDownSheet);

  const row = activeRange.getRow();
  const column = activeRange.getColumn();
  const cellPosition = activeRange.getA1Notation();

  if (
    !activeCol.includes(column) &&
    !activeRow.includes(row) &&
    !autoSpecCells.includes(cellPosition)
  ) {
    return;
  }

  // const cellPosition = activeRange.getA1Notation();
  // const currentCellValue = activeRange.getValue();

  if (activeSheet.getRange(talents.holyShock.talent).getValue() === 1) {
    generateDropDownMenu(
      dropDownSheet,
      activeSheet.getRange(talents.holyShock.glyph.talent).getValue() === true
        ? talents.holyShock.glyph.dropDownOptions
        : talents.holyShock.dropDownOptions,
      talents.holyShock.dropDownLocation
    );
  } else {
    resetDropDownMenus(dropDownSheet, talents.holyShock.dropDownLocation);
  }
  if (
    activeSheet.getRange(talents.divineIllumination.talent).getValue() === 1
  ) {
    generateDropDownMenu(
      dropDownSheet,
      talents.divineIllumination.dropDownOptions,
      talents.divineIllumination.dropDownLocation
    );
  } else {
    resetDropDownMenus(
      dropDownSheet,
      talents.divineIllumination.dropDownLocation
    );
  }
  if (
    activeSheet.getRange(talents.glyphOfHolyLight.talent).getValue() === true
  ) {
    generateDropDownMenu(
      dropDownSheet,
      talents.glyphOfHolyLight.dropDownOptions,
      talents.glyphOfHolyLight.dropDownLocation
    );
  } else {
    resetDropDownMenus(
      dropDownSheet,
      talents.glyphOfHolyLight.dropDownLocation
    );
  }
}

function generateDropDownMenu(sheet, options, location) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      Array.from({ length: options }, (_, index) => index + 1)
    )
    .build();
  const cell = sheet.getRange(location);
  cell.setDataValidation(rule);
  if (cell.getValue() > options) cell.setValue(options);
}

function resetDropDownMenus(sheet, location) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([0])
    .build();
  const cell = sheet.getRange(location);
  cell.setValue("").setDataValidation(rule);
}
