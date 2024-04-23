var workingSheet = "CharacterSettings";
var dataSheetName = "DataItems";
var ActiveColumns = [5, 12, 17, 22, 33];
var ActiveRows = Array.from({ length: 17 }, (_, index) => index + 7);
var enchantHeadOptionsDropDown = [
  "Arcanum of Burning Mysteries",
  "Arcanum of Blissful Mending",
];
var enchantShoulderOptionsDropDown = [
  "Master's Inscription of the Storm",
  "Master's Inscription of the Crag",
  "Greater Inscription of the Storm",
  "Greater Inscription of the Crag",
];
var enchantBackOptionsDropDown = [
  "Springy Arachnoweave",
  "Lightweave Embroidery ",
  "Darkglow Embroidery ",
  "Greater Speed",
];
var enchantChestOptionsDropDown = [
  "Powerful Stats",
  "Greater Mana Restoration",
  "Exceptional Mana",
];
var enchantHandsOptionsDropDown = [
  "Hyperspeed Accelerators",
  "Exceptional Spellpower",
  "Minor Haste",
];
var enchantWristOptionsDropDown = [
  "Fur Lining - Spell Power",
  "Superior Spellpower",
  "Exceptional Intellect",
];
var enchantsLegsOptionsDropDown = [
  "Brilliant Spellthread",
  "Sapphire Spellthread",
];
var enchantsFeetOptionsDropDown = [
  "Nitro Boosts",
  "Icewalker",
  "Greater Vitality",
  "Tuskarr's Vitality",
];
var enchantFingerOptionsDropDown = ["Greater Spellpower"];
var enchantMainHandOptionsDropDown = ["Mighty Spellpower", "Major Intellect"];
var enchantOffHandOptionsDropDown = ["Greater Intellect"];
var enchantOptions = {
  7: enchantHeadOptionsDropDown,
  8: enchantChestOptionsDropDown,
  9: enchantsLegsOptionsDropDown,
  10: enchantShoulderOptionsDropDown,
  11: enchantHandsOptionsDropDown,
  12: enchantWristOptionsDropDown,
  14: enchantsFeetOptionsDropDown,
  15: enchantBackOptionsDropDown,
  17: enchantFingerOptionsDropDown,
  18: enchantFingerOptionsDropDown,
  21: enchantMainHandOptionsDropDown,
  22: enchantOffHandOptionsDropDown,
};
var excludeGemNames = [
  "Prismatic",
  "Yellow",
  "Red",
  "Orange",
  "Purple",
  "Green",
];
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
    "Nightmare Tear",
    "Enchanted Tear",
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
    "Nightmare Tear",
    "Enchanted Tear",
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
    "Nightmare Tear",
    "Enchanted Tear",
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
var correctNumberGems = {
  Tear: 1,
  "Dragon's": 3,
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
var dropDownOptionsMeta = [
  "Meta",
  "Insightful Earthsiege Diamond",
  "Ember Skyflare Diamond",
  "Beaming Earthsiege Diamond",
  "Revitalizing Skyflare Diamond",
];

function applyColorBaseOnItemSockets(e) {
  const spreadSheet = e.source;
  const activeRange = e.source.getActiveRange();
  const sheetName = spreadSheet.getActiveSheet().getName();

  if (sheetName !== workingSheet) return;

  const activeSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const row = activeRange.getRow();
  const column = activeRange.getColumn();

  if (!ActiveColumns.includes(column) || !ActiveRows.includes(row)) return;
  const rowInDataSheet = activeSheet.getRange(`A${row}`).getValue();
  const currentCellValue = activeRange.getValue();

  if (
    !(column === 33 && [11, 12, 13].includes(row)) &&
    containsExactWord(currentCellValue, "Phase")
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
  isGemsCorrectPrismaticAndDragons(activeSheet, currentCellValue, activeRange);

  if (![12, 17, 22, 33].includes(column)) resetGemCells(row, activeSheet);

  if (
    column === 33 &&
    additionalSocket &&
    !activeSheet.getRange(additionalSocket).getValue()
  ) {
    resetOnlyExtraGemSlot(activeSheet, row, data.length);
  }

  if (additionalSocket && activeSheet.getRange(additionalSocket).getValue()) {
    data.push("Meta");
  }

  var bonusForGems = generateGemsDropDownMenu(activeSheet, row, data);
  const isBonusSocket = bonusForGems === data.length && bonusForGems > 0;
  setBonusSocket(activeSheet, row, isBonusSocket, data);
  enchantsDropDown(activeSheet, row);

  activeSheet
    .getRange("J5")
    .setValue(`${isBonusSocket} ${bonusForGems} ${data.length}`);
}

function isGemsCorrectPrismaticAndDragons(sheet, activeCellValue, activeRange) {
  let gemType = "";
  for (const gem in correctNumberGems) {
    if (containsExactWord(activeCellValue, gem)) {
      gemType = gem;
      break;
    }
  }
  if (gemType === "") return;

  const range = sheet.getRange("L7:V23");
  const values = range.getValues().map((x) => JSON.stringify(x).split(","));
  let found = 0;
  values.map((x) =>
    x.map((cellValue) => {
      if (containsExactWord(cellValue, gemType)) found++;
    })
  );
  if (found > correctNumberGems[gemType]) activeRange.setValue("");
}

function containsExactWord(text, word) {
  const regex = new RegExp(`\\b${word}\\b`, "i");
  return regex.test(text);
}

function resetOnlyExtraGemSlot(sheet, row, gems) {
  const cellToChange = sheet.getRange(
    `${gems === 0 ? "L" : gems === 1 ? "Q" : gems === 2 ? "V" : ""}${row}`
  );
  cellToChange.setValue("").setDataValidation(null);
  cellToChange.setBackground("");
}

function resetGemCells(row, sheet) {
  Object.values(colors).forEach((cell) => {
    const cellToChange = sheet.getRange(`${cell}${row}`);
    sheet.getRange(`Z${row}`).setValue("");
    sheet.getRange(`AB${row}`).setValue("").setDataValidation(null);
    cellToChange.setDataValidation(null);
    cellToChange.setBackground("");
    cellToChange.setValue("");
  });
}

function enchantsDropDown(activeSheet, row) {
  const cellToChange = activeSheet.getRange(`AB${row}`);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(enchantOptions[row])
    .build();
  cellToChange.setDataValidation(rule);
}

function setBonusSocket(sheet, row, bonus, data) {
  const bonusCellToChange = sheet.getRange(`Z${row}`);
  bonusCellToChange.setValue(
    data.length === 0 || (data.length === 1 && data.includes("Meta"))
      ? ""
      : bonus
      ? "Yes"
      : "No"
  );
  bonusCellToChange.setFontColor(bonus ? "#6aa84f" : "Red");
}

function generateGemsDropDownMenu(activeSheet, row, data) {
  let bonusTotal = 0;
  data.forEach((color, i) => {
    const cellToChange = activeSheet.getRange(`${colors[i]}${row}`);
    const cellValue = cellToChange.getValue();
    // Browser.msgBox(color);
    bonusTotal += backgroundColorsBonus[backgroundColors[color]].includes(
      cellValue || ""
    )
      ? 1
      : 0;
    const isMetaColpr = color === "Meta" && cellValue !== "Meta";
    bonusTotal += isMetaColpr ? 1 : 0;
    cellToChange.setBackground(backgroundColors[color]);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(
        color === "Meta" && row === 7 ? dropDownOptionsMeta : dropdownOptions
      )
      .build();
    cellToChange.setDataValidation(rule);
  });
  return bonusTotal;
}