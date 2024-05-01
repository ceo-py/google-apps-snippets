var sheetNameDropOptions = "DataCalc";
var sheetNameTalens = "Talents";
var sheetNameCharacter = "CharacterSettings";
var dropDownMenusBaseOnFightTimer = {
  activation: "P39",

  dropDownMenus: {
    "Divine Plea": { position: "F46", optionsPos: "Q23" },
    "Divine Illumination": {
      position: "F47",
      optionsPos: "Q22",
      talentPos: "E39",
      activationCells: ["AL5", "AL6", "E39"],
    },
    "Avenging Wrath": { position: "Y39", optionsPos: "Q21" },
  },
};

function combatTimeDropDownMenus(e) {
  const spreadSheet = e.source;
  const activeRange = e.source.getActiveRange();
  const sheetName = spreadSheet.getActiveSheet().getName();
  if (![sheetNameTalens, sheetNameCharacter].includes(sheetName)) return;

  const cell = activeRange.getA1Notation();
  if (
    cell !== dropDownMenusBaseOnFightTimer.activation &&
    !dropDownMenusBaseOnFightTimer.dropDownMenus[
      "Divine Illumination"
    ].activationCells.includes(cell)
  )
    return;
  const activeSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNameCharacter);
  const rangeDropDown =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNameDropOptions);
  const talentSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNameTalens);

  getDropDownMenusOptions(activeSheet, rangeDropDown, talentSheet);
}

function getDropDownMenusOptions(activeSheet, rangeDropDown, talentRange) {
  Object.values(dropDownMenusBaseOnFightTimer.dropDownMenus).forEach(
    (menu, i) => {
      const dropDownOptionsDynamic = rangeDropDown
        .getRange(menu.optionsPos)
        .getValue();

      if (i === 0 || i === 2) {
        generateDropDownMenu(
          activeSheet,
          dropDownOptionsDynamic,
          menu.position
        );
      } else if (
        i === 1 &&
        talentRange.getRange(menu.talentPos).getValue() === 1
      ) {
        generateDropDownMenu(
          activeSheet,
          dropDownOptionsDynamic,
          menu.position
        );
      } else {
        resetDropDownMenus(
          activeSheet,
          dropDownMenusBaseOnFightTimer.dropDownMenus["Divine Illumination"]
            .position
        );
      }
    }
  );
}

function generateDropDownMenu(sheet, options, location) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      Array.from({ length: options + 1 }, (_, index) => index)
    )
    .build();
  const cell = sheet.getRange(location);
  cell.setDataValidation(rule);
  if (cell.getValue() > options) cell.setValue(options);
}

function resetDropDownMenus(sheet, location) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["0"])
    .build();
  const cell = sheet.getRange(location);
  cell.setDataValidation(rule);
  cell.setValue("0");
}
