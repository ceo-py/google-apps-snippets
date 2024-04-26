const glyphsRange = ["U29", "U30", "U31", "U32", "U33", "U34"];

function ResetCell(cell) {
  var tier, ss, mainSheet, dropdown, rule;
  tier = Generate_Buttons_Options(0, 0);
  ss = SpreadsheetApp.getActiveSpreadsheet();
  mainSheet = ss.getSheetByName("Talents");
  dropdown = mainSheet.getRange(cell);
  rule = SpreadsheetApp.newDataValidation().requireValueInList(tier).build();
  dropdown.setDataValidation(rule);
  dropdown.setBackground("#ff0000");
  dropdown.setBorder(true, true, true, true, false, false);
  dropdown.setFontSize(11);
  dropdown.setFontWeight("bold");
  dropdown.setValue(0);
}

function CreateSpec(cell, spend_points, max_points) {
  var tier, ss, mainSheet, dropdown, rule;
  tier = Generate_Buttons_Options(0, max_points);
  ss = SpreadsheetApp.getActiveSpreadsheet();
  mainSheet = ss.getSheetByName("Talents");
  dropdown = mainSheet.getRange(cell);
  rule = SpreadsheetApp.newDataValidation().requireValueInList(tier).build();
  dropdown.setDataValidation(rule);
  dropdown.setBackground("#ffd1dc");
  dropdown.setValue(spend_points);
}

function CreateCell(cell, tier_point_max, poinst_left, cell_value) {
  var cell_value, tier, ss, mainSheet, dropdown, rule;
  tier = Generate_Buttons_Options(0, tier_point_max);
  ss = SpreadsheetApp.getActiveSpreadsheet();
  mainSheet = ss.getSheetByName("Talents");
  dropdown = mainSheet.getRange(cell);
  rule = SpreadsheetApp.newDataValidation().requireValueInList(tier).build();
  dropdown.setDataValidation(rule);
  dropdown.setBackground("#ffd1dc");
}

function ButtonReset(cell, sheet) {
  var checkbox = SpreadsheetApp.getActive().getRange(sheet + cell);
  checkbox.uncheck();
}

function ButtonCheck(cell, sheet) {
  var checkbox = SpreadsheetApp.getActive().getRange(sheet + cell);
  checkbox.check();
}

function AutoTickFourBonus(cell, sheet) {
  ButtonCheck(ResetTickButtons()[cell], sheet);
  for (const [key, _] of Object.entries(ResetTickButtons())) {
    if (key !== ResetTickButtons()[cell] && key !== cell) {
      ButtonReset(key, sheet);
    }
  }
}

function GetCellValue(cell) {
  return SpreadsheetApp.getActiveSheet().getRange(cell).getValue();
}

function Generate_Buttons_Options(start, end) {
  return Array(end - start + 1)
    .fill()
    .map((_, idx) => start + idx);
}

function takeRangeToReset(tierRow, start, tierTree) {
  Object.keys(GenerateSpec())
    .slice(start, tierTree)
    .forEach((cell) => {
      if (cell.includes(tierRow)) {
        ResetCell(cell);
      }
    });
}

function takeRangeToCreate(tierRow, start, tierTree) {
  for (const [key, value] of Object.entries(GenerateSpec()).slice(
    start,
    tierTree
  )) {
    Object.keys(GenerateSpec())
      .slice(start, tierTree)
      .forEach((cell) => {
        if (key.includes(tierRow)) {
          CreateCell(key, value[2], value[2]);
        }
      });
  }
}

function ResetTickButtons() {
  return {
    V4: "V4",
    V5: "V4",
    V6: "V6",
    V7: "V6",
    V8: "V8",
    V9: "V8",
    V10: "V10",
    V11: "V10",
  };
}

function GenerateSpec() {
  spec = {
    G7: [5, 5, 5],
    I7: [0, 0, 5],
    E11: [3, 3, 3],
    G11: [5, 5, 5],
    I11: [0, 0, 2],
    E15: [1, 1, 1],
    G15: [5, 5, 5],
    I15: [2, 2, 2],
    E19: [0, 0, 3],
    I19: [0, 0, 2],
    K19: [1, 1, 2],
    E23: [0, 0, 2],
    G23: [1, 1, 1],
    I23: [3, 3, 3],
    E27: [0, 0, 2],
    I27: [5, 5, 5],
    E31: [3, 3, 3],
    G31: [1, 1, 1],
    I31: [0, 0, 3],
    E35: [0, 0, 3],
    I35: [5, 5, 5],
    E39: [1, 1, 1],
    I39: [5, 5, 5],
    G43: [2, 2, 2],
    I43: [2, 2, 2],
    G47: [1, 1, 1],
    S7: [5, 5, 5],
    U7: [0, 0, 5],
    Q11: [3, 0, 3],
    S11: [2, 0, 2],
    U11: [0, 0, 5],
    Q15: [1, 0, 1],
    S15: [3, 0, 3],
    U15: [1, 0, 5],
    Q19: [2, 0, 2],
    S19: [0, 0, 2],
    U19: [3, 0, 3],
    AE7: [0, 0, 5],
    AG7: [0, 5, 5],
    AC11: [0, 0, 2],
    AE11: [0, 3, 3],
    AG11: [0, 2, 2],
    AC15: [0, 0, 2],
    AE15: [0, 5, 5],
    AG15: [0, 0, 1],
    AI15: [0, 0, 2],
    AC19: [0, 0, 2],
    AG19: [0, 0, 3],
    AI19: [0, 0, 3],
  };
  return spec;
}

function holyTier1() {
  if (GetCellValue("B14") != 0) {
    takeRangeToReset("15", 0, 25);
  } else {
    takeRangeToCreate("15", 0, 25);
  }
}

function holyTier2() {
  if (GetCellValue("B18") != 0) {
    takeRangeToReset("19", 0, 25);
  } else {
    takeRangeToCreate("19", 0, 25);
  }
}

function holyTier3() {
  if (GetCellValue("B22") != 0) {
    takeRangeToReset("23", 0, 25);
  } else {
    CreateCell("E23", 2, 0);
    CreateCell("I23", 3, 0);
    var checkCell = GetCellValue("G15");
    if (checkCell == 5) {
      CreateCell("G23", 1, 0);
    }
    if (checkCell != 5) {
      ResetCell("G23");
    }
  }
}

function holyTier4() {
  if (GetCellValue("B26") != 0) {
    takeRangeToReset("27", 0, 25);
  } else {
    takeRangeToCreate("27", 0, 25);
  }
}

function holyTier5() {
  if (GetCellValue("B30") != 0) {
    takeRangeToReset("31", 0, 25);
  } else {
    CreateCell("E31", 3, 0);
    CreateCell("I31", 3, 0);
    var checkCell = GetCellValue("G23");
    if (checkCell == 1) {
      CreateCell("G31", 1, 0);
    }
    if (checkCell != 1) {
      ResetCell("G31");
    }
  }
}

function holyTier6() {
  if (GetCellValue("B34") != 0) {
    takeRangeToReset("35", 0, 25);
  } else {
    takeRangeToCreate("35", 0, 25);
  }
}

function holyTier7() {
  if (GetCellValue("B38") != 0) {
    takeRangeToReset("39", 0, 25);
  } else {
    takeRangeToCreate("39", 0, 25);
  }
}

function holyTier8() {
  if (GetCellValue("B42") != 0) {
    takeRangeToReset("43", 0, 25);
  } else {
    CreateCell("I43", 2, 0);
    var checkCell = GetCellValue("G31");
    if (checkCell == 1) {
      CreateCell("G43", 2, 0);
    }
    if (checkCell != 1) {
      ResetCell("G43");
    }
  }
}

function holyTier9() {
  if (GetCellValue("B46") != 0) {
    ResetCell("G47");
  } else {
    CreateCell("G47", 1, 0);
  }
}

function protTier1() {
  if (GetCellValue("N14") != 0) {
    takeRangeToReset("15", 26, 37);
  } else {
    takeRangeToCreate("15", 26, 37);
  }
}

function protTier2() {
  if (GetCellValue("N18") != 0) {
    takeRangeToReset("19", 26, 37);
  } else {
    CreateCell("S19", 2, 0);
    CreateCell("U19", 3, 0);
    var checkCell = GetCellValue("Q15");
    if (checkCell == 1) {
      CreateCell("Q19", 2, 0);
    }
    if (checkCell != 1) {
      ResetCell("Q19");
    }
  }
}

function retTier1() {
  if (GetCellValue("Z14") != 0) {
    takeRangeToReset("15", 37, 50);
  } else {
    takeRangeToCreate("15", 37, 50);
  }
}

function retTier2() {
  if (GetCellValue("Z18") != 0) {
    takeRangeToReset("19", 37, 50);
  } else {
    takeRangeToCreate("19", 37, 50);
  }
}

function onEdit(e) {
  var spreadSheet,
    sheetName,
    cell,
    a1,
    poinst_left_to_spend,
    cell,
    check_box_value,
    buttons_check_number;
  spreadSheet = e.source;
  sheetName = spreadSheet.getActiveSheet().getName();
  cell = SpreadsheetApp.getActiveSheet().getActiveCell();
  a1 = cell.getA1Notation();
  // Browser.msgBox(a1)
  if (sheetName != "Talents") return;

  // if (cell) return
  if (glyphsRange.includes(a1)) {
    const rangeValues = spreadSheet.getRange("U29:U34").getValues();
    const check_box_value = rangeValues.filter((x) =>
      JSON.stringify(x).includes("true")
    ).length;
    if (check_box_value > 3) cell.setValue("");
    return;
  }

  if (
    sheetName === "Talents" &&
    Object.keys(GenerateSpec()).slice(0, 25).includes(a1)
  ) {
    ButtonReset("AL5", "Talents!");
    ButtonReset("AL6", "Talents!");
    if (GetCellValue("B10") != 0) {
      Object.keys(GenerateSpec())
        .slice(0, 26)
        .forEach((cell) => {
          if (cell !== "I7" && cell !== "G7") {
            ResetCell(cell);
          }
        });
    } else {
      takeRangeToCreate("11", 0, 25);
      if (Object.keys(GenerateSpec()).slice(0, 5).includes(a1)) {
        holyTier1();
        holyTier2();
        holyTier3();
        holyTier4();
        holyTier5();
        holyTier6();
        holyTier7();
      }
      if (Object.keys(GenerateSpec()).slice(0, 8).includes(a1)) {
        holyTier2();
        holyTier3();
        holyTier4();
        holyTier6();
        holyTier7();
        holyTier8();
        holyTier9();
      }
      if (Object.keys(GenerateSpec()).slice(0, 11).includes(a1)) {
        holyTier3();
        holyTier4();
        holyTier5();
        holyTier6();
        holyTier7();
        holyTier8();
        holyTier9();
      }
      if (Object.keys(GenerateSpec()).slice(0, 14).includes(a1)) {
        holyTier4();
        holyTier5();
        holyTier6();
        holyTier7();
        holyTier8();
        holyTier9();
      }
      if (Object.keys(GenerateSpec()).slice(0, 16).includes(a1)) {
        holyTier5();
        holyTier6();
        holyTier7();
        holyTier8();
        holyTier9();
      }
      if (Object.keys(GenerateSpec()).slice(0, 19).includes(a1)) {
        holyTier6();
        holyTier7();
        holyTier8();
        holyTier9();
      }
      if (Object.keys(GenerateSpec()).slice(0, 21).includes(a1)) {
        holyTier7();
        holyTier8();
        holyTier9();
      }
      if (Object.keys(GenerateSpec()).slice(0, 25).includes(a1)) {
        holyTier8();
        holyTier9();
      }
    }
  }
  if (
    sheetName === "Talents" &&
    Object.keys(GenerateSpec()).slice(26, 37).includes(a1)
  ) {
    ButtonReset("AL5", "Talents!");
    ButtonReset("AL6", "Talents!");
    if (GetCellValue("N10") != 0) {
      Object.keys(GenerateSpec())
        .slice(26, 37)
        .forEach(function (cell) {
          if (cell !== "S7" && cell !== "U7") {
            ResetCell(cell);
          }
        });
    } else {
      takeRangeToCreate("11", 26, 37);
      if (Object.keys(GenerateSpec()).slice(26, 31).includes(a1)) {
        protTier1();
        protTier2();
      }
      if (Object.keys(GenerateSpec()).slice(26, 37).includes(a1)) {
        protTier2();
      }
    }
  }
  if (
    sheetName === "Talents" &&
    Object.keys(GenerateSpec()).slice(37, 50).includes(a1)
  ) {
    ButtonReset("AL5", "Talents!");
    ButtonReset("AL6", "Talents!");
    if (GetCellValue("Z10") != 0) {
      Object.keys(GenerateSpec())
        .slice(37, 50)
        .forEach(function (cell) {
          if (cell !== "AE7" && cell !== "AG7") {
            ResetCell(cell);
          }
        });
    } else {
      takeRangeToCreate("11", 37, 50);
      if (Object.keys(GenerateSpec()).slice(37, 42).includes(a1)) {
        retTier1();
        retTier2();
      }
      if (Object.keys(GenerateSpec()).slice(37, 46).includes(a1)) {
        retTier2();
      }
    }
  }
  cell = SpreadsheetApp.getActiveSheet().getActiveCell();
  check_box_value = cell.getValue();
  if (sheetName === "Talents" && a1 === "AL5" && check_box_value === true) {
    ButtonReset("AL6", "Talents!");
    for (const [key, value] of Object.entries(GenerateSpec())) {
      CreateSpec(key, value[0], value[2]);
    }
    Object.keys(GenerateSpec())
      .slice(37, 50)
      .forEach(function (cell) {
        if (cell !== "AE7" && cell !== "AG7") {
          ResetCell(cell);
        }
      });
  }
  if (sheetName === "Talents" && a1 === "AL6" && check_box_value === true) {
    ButtonReset("AL5", "Talents!");
    for (const [key, value] of Object.entries(GenerateSpec())) {
      CreateSpec(key, value[1], value[2]);
    }
    Object.keys(GenerateSpec())
      .slice(26, 37)
      .forEach(function (cell) {
        if (
          cell !== "S7" &&
          cell !== "U7" &&
          cell !== "Q11" &&
          cell !== "S11" &&
          cell !== "U11"
        ) {
          ResetCell(cell);
        }
      });
  }
  if (
    sheetName === "Calculator" &&
    Object.keys(ResetTickButtons())
      .filter((_, i) => i % 2 == 1)
      .includes(a1) &&
    check_box_value == true
  ) {
    AutoTickFourBonus(a1, "Calculator!");
  }

  if (
    sheetName === "Calculator" &&
    Object.keys(ResetTickButtons())
      .filter((_, i) => i % 2 == 0)
      .includes(a1)
  ) {
    Object.keys(ResetTickButtons())
      .filter((_, i) => i % 2 == 1)
      .forEach(function (cell) {
        ButtonReset(cell, "Calculator!");
      });
    buttons_check_number = 0;
    Object.keys(ResetTickButtons())
      .filter((_, i) => i % 2 == 0)
      .forEach((cell) => {
        if (SpreadsheetApp.getActiveSheet().getRange(cell).getValue()) {
          buttons_check_number++;
        }
      });
    if (buttons_check_number > 2) {
      Object.keys(ResetTickButtons())
        .filter((_, i) => i % 2 == 0)
        .forEach((cell) => {
          if (cell != a1) {
            ButtonReset(cell, "Calculator!");
          }
        });
    }
  }
}
