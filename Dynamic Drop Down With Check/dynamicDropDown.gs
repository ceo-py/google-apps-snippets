var workingSheet = "CharacterSettings";
var dataSheetName = "DataItems";
var ActiveColumns = [5, 12, 17, 22, 35];
var ActiveRows = Array.from({ length: 17 }, (_, index) => index + 7);
var skipingUniqueItems = [
  "(A) Solace of the Defeated (245)",
  "(A) Solace of the Defeated (258)",
  "(A) Ring of the Darkmender (245)",
  "(A) Ring of the Darkmender (258)",
  "(H) Solace of the Fallen (245)",
  "(H) Solace of the Fallen (258)",
  "(H) Circle of the Darkmender (245)",
  "(H) Circle of the Darkmender (258)",
];
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
  1135: "AI11",
  1235: "AI12",
  1335: "AI13",
  115: "AI11",
  125: "AI12",
  135: "AI13",

  1112: "AI11",
  1117: "AI11",
  1122: "AI11",

  1212: "AI12",
  1217: "AI12",
  1222: "AI12",

  1312: "AI13",
  1317: "AI13",
  1322: "AI13",
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
var metaActivationReq = {
  "Insightful Earthsiege Diamond": { Red: 1, Yellow: 1, Blue: 1 },
  "Ember Skyflare Diamond": { Red: 3 },
  "Beaming Earthsiege Diamond": { Red: 2, Yellow: 1 },
  "Revitalizing Skyflare Diamond": { Red: 2 },
};
var setBonus = {
  range: { row: [7, 8, 9, 10, 11], col: 5 },
  bonusActiveRange: ["P26", "P27", "P28", "P29", "P30", "P31", "P32", "P33"],
  tiers: {
    tier7: {
      name: "Heroes' Redemption",
      "2 Pieces": "P26",
      "4 Pieces": "P27",
    },
    tier71: {
      name: "Valorous Redemption",
      "2 Pieces": "P26",
      "4 Pieces": "P27",
    },
    tier8: {
      name: "Valorous Aegis",
      "2 Pieces": "P28",
      "4 Pieces": "P29",
    },
    tier81: {
      name: "Conqueror's Aegis",
      "2 Pieces": "P28",
      "4 Pieces": "P29",
    },
    tier9: {
      name: "Turalyon's",
      "2 Pieces": "P30",
      "4 Pieces": "P31",
    },
    tier91: {
      name: "Liadrin's",
      "2 Pieces": "P30",
      "4 Pieces": "P31",
    },
    tier10: {
      name: "Lightsworn",
      "2 Pieces": "P32",
      "4 Pieces": "P33",
    },
  },
};

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
    !(column === 35 && [11, 12, 13].includes(row)) &&
    containsExactWord(currentCellValue, "Phase")
  ) {
    resetGemCells(row, activeSheet);
    isSetBonus(activeSheet);
    return;
  }
  // activeSheet.getRange("H5").setValue(`${row} ${column}`);

  const additionalSocket = bsAndSocketBelt?.[`${row}${column}`];
  const dataSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheetName);
  const targetRange = dataSheet.getRange(
    `N${rowInDataSheet}:P${rowInDataSheet}`
  );
  const data = targetRange.getValues()[0].filter((gem) => gem !== "");
  // activeSheet.getRange("I5").setValue(additionalSocket);

  if ([12, 17, 22, 35].includes(column)) {
    isGemsCorrectPrismaticAndDragons(
      activeSheet,
      currentCellValue,
      activeRange
    );
  }

  const gameVersion = activeSheet.getRange("B4").getValue();

  if (
    column === 5 &&
    [17, 18, 19, 20].includes(row) &&
    gameVersion === "Wrath Classic (v3.4.3)" &&
    isItemUnique(activeSheet, activeRange, row)
  )
    return;
  else if (
    column === 5 &&
    gameVersion === "Wrath OG (v3.3.5)" &&
    [17, 18, 19, 20].includes(row) &&
    !isRingTrinketUniqueOldVersion(
      activeSheet,
      activeRange,
      currentCellValue,
      row
    )
  ) {
    return;
  }

  if (![12, 17, 22, 35].includes(column)) resetGemCells(row, activeSheet);
  // Extra Socket Change only
  if (
    column === 35 &&
    additionalSocket &&
    !activeSheet.getRange(additionalSocket).getValue()
  ) {
    resetOnlyExtraGemSlot(activeSheet, row, data.length);
  }

  if (
    additionalSocket &&
    activeSheet.getRange(additionalSocket).getValue() &&
    !containsExactWord(activeSheet.getRange(`E${row}`).getValue(), "Phase")
  ) {
    data.push("Meta");
  }

  var bonusForGems = generateGemsDropDownMenu(activeSheet, row, data);
  const isBonusSocket = bonusForGems === data.length && bonusForGems > 0;

  if ([5, 12, 17, 22].includes(column)) {
    setBonusSocket(activeSheet, row, isBonusSocket, data);
  }

  if ([5].includes(column)) {
    enchantsDropDown(activeSheet, row);
  }

  isMetaActive(activeSheet, activeSheet.getRange("L7").getValue(), activeSheet);

  if (setBonus.range.row.includes(row) && setBonus.range.col === column) {
    isSetBonus(activeSheet);
  }

  // activeSheet
  //   .getRange("J5")
  //   .setValue(`${isBonusSocket} ${bonusForGems} ${data.length}`);
}
function getTotalGems(data) {
  const subColors = {
    Orange: { Red: 0, Yellow: 0 },
    Purple: { Red: 0, Blue: 0 },
    Green: { Yellow: 0, Blue: 0 },
  };
  const foundGems = {};
  for (const cValue of data) {
    if (cValue === "") continue;

    if (cValue === "Nightmare Tear" || cValue === "Enchanted Tear") {
      foundGems["Tear"] = 1;
      continue;
    }

    Object.keys(correctGemColors).forEach((c) => {
      const gemCorrect = correctGemColors[c].includes(cValue);
      if (subColors.hasOwnProperty(c) && gemCorrect) {
        for (const key in subColors[c]) {
          subColors[c][key] += 1;
        }
      } else if (gemCorrect) {
        if (!foundGems.hasOwnProperty(c)) foundGems[c] = 0;
        foundGems[c]++;
      }
    });
  }
  return { foundGems, subColors };
}

function tearGemAdd(metaType, foundGems) {
  if (foundGems.hasOwnProperty("Tear")) {
    for (const color in metaActivationReq[metaType]) {
      metaActivationReq[metaType][color]--;
      if (metaActivationReq[metaType][color] <= 0)
        delete metaActivationReq[metaType][color];
    }
    delete foundGems["Tear"];
  }
  return foundGems;
}

function decrementColors(initialDict, compareDict) {
  if (
    !isColorsCorrect(
      initialDict,
      Object.values(compareDict)
        .map((x) => Object.keys(x).map((y) => y))
        .flat()
    )
  )
    return false;

  while (true) {
    if (Object.keys(initialDict).length === 0) return true;
    const topColor = Object.keys(initialDict).reduce((a, b) =>
      initialDict[a] > initialDict[b] ? a : b
    );

    if (Object.keys(compareDict).length === 0) return false;
    const maxKey = getKeyWithMostValue(compareDict, topColor);
    if (!maxKey) return false;

    if (initialDict[topColor] >= compareDict[maxKey][topColor]) {
      initialDict[topColor] -= compareDict[maxKey][topColor];
      delete compareDict[maxKey];
      if (initialDict[topColor] === 0) delete initialDict[topColor];
    } else if (initialDict[topColor] < compareDict[maxKey][topColor]) {
      const deductedValue = initialDict[topColor];
      delete initialDict[topColor];
      Object.values(compareDict).forEach((c) => (c -= deductedValue));
    }
  }
}

function getKeyWithMostValue(dict, topColor) {
  let maxKey = null;
  let maxValue = -Infinity;

  for (const key in dict) {
    let totalValue = 0;
    const colorValues = dict[key];
    for (const color in colorValues) {
      totalValue += colorValues[color];
    }
    if (totalValue > maxValue && dict[key].hasOwnProperty(topColor)) {
      maxValue = totalValue;
      maxKey = key;
    }
  }
  return maxKey;
}

function isColorsCorrect(initialDict, compareDict) {
  const initialSet = new Set(Object.keys(initialDict).map((x) => x));
  const compareSet = new Set(compareDict);
  for (let item of initialSet) {
    if (!compareSet.has(item)) {
      return false;
    }
  }
  return true;
}

function isMetaActive(sheet, metaType, sheet) {
  if (!metaActivationReq.hasOwnProperty(metaType)) return;

  const range = sheet.getRange("L7:V23");
  const values = range
    .getValues()
    .map((x) => JSON.stringify(x))
    .map((x) => JSON.parse(x))
    .filter((x) => x.length !== 0)
    .flat();

  let { foundGems, subColors } = getTotalGems(values);
  foundGems = tearGemAdd(metaType, foundGems);

  for (const color in metaActivationReq[metaType]) {
    if (foundGems.hasOwnProperty(color)) {
      metaActivationReq[metaType][color] -= foundGems[color];
      if (metaActivationReq[metaType][color] <= 0)
        delete metaActivationReq[metaType][color];
      delete foundGems[color];
    }
  }

  // sheet
  //   .getRange("J2")
  //   .setValue(`${JSON.stringify(metaActivationReq[metaType])}`);

  // sheet.getRange("J3").setValue(`${JSON.stringify(foundGems)}`);

  const missingToActive = Object.values(metaActivationReq[metaType]).reduce(
    (a, b) => a + b,
    0
  );
  const metaActive = missingToActive === 0;

  // sheet.getRange("J1").setValue(missingToActive);
  // sheet.getRange("I2").setValue(subColors);
  // sheet.getRange("L7").setFontColor("Red");
  // sheet.getRange("AI14").setValue(metaActive ? "Yes" : "No");

  if (metaActive || decrementColors(metaActivationReq[metaType], subColors)) {
    sheet.getRange("L7").setFontColor("#6aa84f");
    sheet.getRange("AI14").setValue("Yes");
    return;
  }
  sheet.getRange("AI14").setValue("No");
  sheet.getRange("L7").setFontColor("Red");
}

function resetSetBonus(sheet) {
  setBonus.bonusActiveRange.forEach((cell) => {
    sheet.getRange(cell).setValue("No");
    sheet.getRange(cell).setFontColor("Red");
  });
}

function isSetBonus(sheet) {
  const setBonusItems = sheet.getRange("E7:E11").getValues();
  resetSetBonus(sheet);
  Object.keys(setBonus.tiers).forEach((key) => {
    const tier = setBonus.tiers[key];
    const countItems = setBonusItems.filter((x) =>
      containsExactWord(x, tier.name)
    ).length;
    const fourPeaceCell = sheet.getRange(tier["4 Pieces"]);
    const twoPeaceCell = sheet.getRange(tier["2 Pieces"]);
    if ([4, 5].includes(countItems)) {
      fourPeaceCell.setValue("Yes");
      fourPeaceCell.setFontColor("#6aa84f");
      twoPeaceCell.setValue("Yes");
      twoPeaceCell.setFontColor("#6aa84f");
    } else if ([2, 3].includes(countItems)) {
      twoPeaceCell.setValue("Yes");
      twoPeaceCell.setFontColor("#6aa84f");
    }
  });
}

function isRingTrinketUniqueOldVersion(
  sheet,
  activeRange,
  currentCellValue,
  row
) {
  if (
    sheet.getRange("E18").getValue() !== sheet.getRange("E17").getValue() &&
    sheet.getRange("E20").getValue() !== sheet.getRange("E19").getValue() &&
    skipingUniqueItems.includes(currentCellValue)
  )
    return true;
  return isItemUnique(sheet, activeRange, row);
}

function isItemUnique(sheet, activeRange, row) {
  const range = sheet.getRange("E17:E20");
  const values = range
    .getValues()
    .flat()
    .map((x, i) =>
      x === "" ? i : containsExactWord(x, "Phase") ? i : x.split(" (")[0]
    );
  if (values.length !== new Set(values).size) {
    resetGemCells(row, sheet);
    activeRange.setValue("");
    return false;
  }
  return true;
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
  const values = range
    .getValues()
    .map((x) => JSON.stringify(x).split(","))
    .flat();
  let found = 0;
  values.map((cellValue) => {
    if (containsExactWord(cellValue, gemType)) found++;
  });
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
    const isMetaColpr =
      (color === "Meta" && row !== 7 && cellValue !== "Meta") ||
      (row === 7 &&
        color === "Meta" &&
        cellValue !== "Meta" &&
        cellValue !== "");
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
