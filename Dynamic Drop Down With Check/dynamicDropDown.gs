var workingSheet = "CharacterSettings";
var dataSheetName = "DataItems";
var colors = {
  0: "L",
  1: "Q",
  2: "V",
};
var backgroundColors = {
  Meta: "#bdbdbd",
  Blue: "#0096ff",
  Red: "Red",
  Yellow: "Yellow",
};

var backgroundColorsBonus = {
  "#bdbdbd": ["#bdbdbd"],
  "#0096ff": [
    "Dazzling Eye of Zul",
    "Energized Eye of Zul",
    "Dazzling Forest Emerald",
    "Energized Forest Emerald",
    "Perfect Dazzling Dark Jade",
    "Perfect Energized Dark Jade",
    "Dazzling Dark Jade",
    "Energized Dark Jade",
    "Dazzling Eye of Zul",
    "Energized Eye of Zul",
    "Dazzling Forest Emerald",
    "Energized Forest Emerald",
    "Perfect Dazzling Dark Jade",
    "Perfect Energized Dark Jade",
    "Dazzling Dark Jade",
    "Energized Dark Jade",
    "Royal Dreadstone",
    "Royal Twilight Opal",
    "Perfect Royal Shadow Crystal",
    "Royal Shadow Crystal",
  ],
  Red: [
    "Luminous Ametrine",
    "Reckless Ametrine",
    "Luminous Monarch Topaz",
    "Reckless Monarch Topaz",
    "Perfect Luminous Huge Citrine",
    "Perfect Reckless Huge Citrine",
    "Luminous Huge Citrine",
    "Reckless Huge Citrine",
    "Royal Dreadstone",
    "Royal Twilight Opal",
    "Perfect Royal Shadow Crystal",
    "Royal Shadow Crystal",
    "Runed Dragon's Eye",
    "Runed Cardinal Ruby",
    "Runed Scarlet Ruby",
    "Perfect Runed Bloodstone",
    "Runed Bloodstone",
  ],
  Yellow: [
    "Dazzling Eye of Zul",
    "Energized Eye of Zul",
    "Dazzling Forest Emerald",
    "Energized Forest Emerald",
    "Perfect Dazzling Dark Jade",
    "Perfect Energized Dark Jade",
    "Dazzling Dark Jade",
    "Energized Dark Jade",
    "Luminous Ametrine",
    "Reckless Ametrine",
    "Luminous Monarch Topaz",
    "Reckless Monarch Topaz",
    "Perfect Luminous Huge Citrine",
    "Perfect Reckless Huge Citrine",
    "Luminous Huge Citrine",
    "Reckless Huge Citrine",
    "Brilliant Dragon's Eye",
    "Quick Dragon's Eye",
    "Brilliant King's Amber",
    "Quick King's Amber",
    "Brilliant Autumn's Glow",
    "Quick Autumn's Glow",
    "Perfect Brilliant Sun Crystal",
    "Perfect Quick Sun Crystal",
    "Brilliant Sun Crystal",
    "Quick Sun Crystal",
  ],
};
var bsAndSocketBelt = {
  1133: "AG11",
  1233: "AG12",
  1333: "AG13",
  115: "AG11",
  125: "AG12",
  135: "AG13",

  1112: "AG11",
  1117: "AG11",
  1122: "AG11",

  1212: "AG12",
  1217: "AG12",
  1222: "AG12",

  1312: "AG13",
  1317: "AG13",
  1322: "AG13",
};
var correctGemColors = {
  Yellow: [
    "Brilliant Dragon's Eye",
    "Quick Dragon's Eye",
    "Brilliant King's Amber",
    "Quick King's Amber",
    "Brilliant Autumn's Glow",
    "Quick Autumn's Glow",
    "Perfect Brilliant Sun Crystal",
    "Perfect Quick Sun Crystal",
    "Brilliant Sun Crystal",
    "Quick Sun Crystal",
  ],
  Red: [
    "Runed Dragon's Eye",
    "Runed Cardinal Ruby",
    "Runed Scarlet Ruby",
    "Perfect Runed Bloodstone",
    "Runed Bloodstone",
  ],
  Orange: [
    "Luminous Ametrine",
    "Reckless Ametrine",
    "Luminous Monarch Topaz",
    "Reckless Monarch Topaz",
    "Perfect Luminous Huge Citrine",
    "Perfect Reckless Huge Citrine",
    "Luminous Huge Citrine",
    "Reckless Huge Citrine",
  ],
  Purple: [
    "Royal Dreadstone",
    "Royal Twilight Opal",
    "Perfect Royal Shadow Crystal",
    "Royal Shadow Crystal",
  ],
  Green: [
    "Dazzling Eye of Zul",
    "Energized Eye of Zul",
    "Dazzling Forest Emerald",
    "Energized Forest Emerald",
    "Perfect Dazzling Dark Jade",
    "Perfect Energized Dark Jade",
    "Dazzling Dark Jade",
    "Energized Dark Jade",
  ],
};

var dropdownOptions = [
  "Prismatic",
  "Nightmare Tear",
  "Enchanted Tear",
  "Yellow",
  "Brilliant Dragon's Eye",
  "Quick Dragon's Eye",
  "Brilliant King's Amber",
  "Quick King's Amber",
  "Brilliant Autumn's Glow",
  "Quick Autumn's Glow",
  "Perfect Brilliant Sun Crystal",
  "Perfect Quick Sun Crystal",
  "Brilliant Sun Crystal",
  "Quick Sun Crystal",
  "Red",
  "Runed Dragon's Eye",
  "Runed Cardinal Ruby",
  "Runed Scarlet Ruby",
  "Perfect Runed Bloodstone",
  "Runed Bloodstone",
  "Orange",
  "Luminous Ametrine",
  "Reckless Ametrine",
  "Luminous Monarch Topaz",
  "Reckless Monarch Topaz",
  "Perfect Luminous Huge Citrine",
  "Perfect Reckless Huge Citrine",
  "Luminous Huge Citrine",
  "Reckless Huge Citrine",
  "Purple",
  "Royal Dreadstone",
  "Royal Twilight Opal",
  "Perfect Royal Shadow Crystal",
  "Royal Shadow Crystal",
  "Green",
  "Dazzling Eye of Zul",
  "Energized Eye of Zul",
  "Dazzling Forest Emerald",
  "Energized Forest Emerald",
  "Perfect Dazzling Dark Jade",
  "Perfect Energized Dark Jade",
  "Dazzling Dark Jade",
  "Energized Dark Jade",
];

function applyColorBaseOnItemSockets(e) {
  const spreadSheet = e.source;
  const activeRange = e.source.getActiveRange();
  const sheetName = spreadSheet.getActiveSheet().getName();

  if (sheetName !== workingSheet) {
    return;
  }
  const activeSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const row = activeRange.getRow();
  const column = activeRange.getColumn();
  const rowInDataSheet = activeSheet.getRange(`A${row}`).getValue();
  const currentCellValue = activeRange.getValue();

  if (
    ![33, 12, 17, 22].includes(column) &&
    [11, 12, 13].includes(row) &&
    (column !== 5 || row > 23 || row < 7)
  ) {
    return;
  }

  if (
    !(column === 33 && [11, 12, 13].includes(row)) &&
    currentCellValue.includes("Phase")
  ) {
    resetGemCells(row, activeSheet);
    return;
  }

  activeSheet.getRange("H5").setValue(`${row} ${column}`);

  const additionalSocket = bsAndSocketBelt?.[`${row}${column}`];
  const dataSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheetName);
  const targetRange = dataSheet.getRange(
    `N${rowInDataSheet}:P${rowInDataSheet}`
  );
  const data = targetRange.getValues()[0].filter((gem) => gem !== "");
  activeSheet.getRange("I5").setValue(additionalSocket);

  if (![12, 17, 22].includes(column)) resetGemCells(row, activeSheet);
  //   resetGemCells(row, activeSheet);

  if (
    additionalSocket &&
    activeSheet.getRange(additionalSocket).getValue() === true
  ) {
    data.push("Meta");
  }

  const bonusForGems = generateGemsDropDownMenu(activeSheet, row, data);
  const isBonusSocket = bonusForGems === data.length && bonusForGems > 0;
  setBonusSocket(activeSheet, row, isBonusSocket);
  activeSheet
    .getRange("J5")
    .setValue(`${isBonusSocket} ${bonusForGems} ${data.length}`);
}

function resetGemCells(row, sheet) {
  Object.values(colors).forEach((cell) => {
    const cellToChange = sheet.getRange(`${cell}${row}`);
    cellToChange.setDataValidation(null);
    cellToChange.setBackground(row % 2 === 0 ? "#ffd1dc" : "");
    cellToChange.setValue("");
  });
}

function setBonusSocket(sheet, row, bonus) {
  const bonusCellToChange = sheet.getRange(`Z${row}`);
  bonusCellToChange.setValue(bonus);
  bonusCellToChange.setFontColor(bonus ? "#00FF00" : "Red");
}

function generateGemsDropDownMenu(activeSheet, row, data) {
  let bonusTotal = 0;
  data.forEach((color, i) => {
    const cellToChange = activeSheet.getRange(`${colors[i]}${row}`);
    // Browser.msgBox(color);
    bonusTotal += backgroundColorsBonus[backgroundColors[color]].includes(
      cellToChange.getValue() || ""
    )
      ? 1
      : 0;
    bonusTotal += color === "Meta" ? 1 : 0;
    cellToChange.setBackground(backgroundColors[color]);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(dropdownOptions)
      .build();
    cellToChange.setDataValidation(rule);
  });
  return bonusTotal;
}
