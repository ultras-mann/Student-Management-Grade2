/**
 * ------------------------------------------------------------------
 * 1. UI MENU & DIALOG ROUTING
 * ------------------------------------------------------------------
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Student Management')
    .addItem('1. បញ្ជីឈ្មោះ​សិស្ស (Sheet 1)', 'showFormDialog')
    .addItem('2. តារាងពិន្ទុសិស្ស (Sheet 2)', 'showPointsDialog')
    .addItem('3. តារាងចំណាត់ថ្នាក់ (Sheet 3)', 'showAnalysisDialog')
    .addToUi();
}

function showFormDialog() {
  var html = HtmlService.createHtmlOutputFromFile('index').setWidth(650).setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, 'Student Registration Form');
}

function showPointsDialog() {
  var html = HtmlService.createHtmlOutputFromFile('points').setWidth(550).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Student Points Entry Form');
}

function showAnalysisDialog() {
  var html = HtmlService.createHtmlOutputFromFile('analysis').setWidth(850).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Student Rank & Analysis');
}

function doGet(e) {
  var page = e && e.parameter && e.parameter.page ? e.parameter.page : 'index';
  return HtmlService.createHtmlOutputFromFile(page)
    .setTitle('Student System')
    .setXframeOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * ------------------------------------------------------------------
 * HELPER: Retrieve or create sheet by period (Month 1-12 or Semester 1-2)
 * ------------------------------------------------------------------
 */
function getPointsSheetByMonth(ss, period) {
  var sheetName = "";
  
  if (period === "S1" || period === "Semester 1") {
    sheetName = "Point of Student - Semester 1";
  } else if (period === "S2" || period === "Semester 2") {
    sheetName = "Point of Student - Semester 2";
  } else {
    sheetName = "Point of Student - Month " + (period || 1);
  }

  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    // Add default headers on row 9 for 7 subjects + Total (Columns A to I: ល.រ, ឈ្មោះសិស្ស, ខ្មែរ, គណិត, ចែងសេចក្តី, សរសេរតាមអាន, អំណាន, អក្សរផ្ចង់, គំនូរ, សរុប)
    sheet.getRange("A9:J9").setValues([[
      "ល.រ", "ឈ្មោះសិស្ស", "ខ្មែរ", "គណិត", "តែងសេចក្តី", "សរសេរតាមអាន", "អំណាន", "អក្សរផ្ចង់", "គំនូរ", "សរុប"
    ]]);
    sheet.getRange("A9:J9").setFontWeight("bold");
  }
  return sheet;
}

/**
 * ------------------------------------------------------------------
 * 2. STUDENT REGISTRATION (index.html)
 * ------------------------------------------------------------------
 */
function getAllStudentList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Student Management") || ss.getSheets()[0];
    var lastRow = sheet.getLastRow();
    if (lastRow < 10) return [];
    
    var data = sheet.getRange(10, 2, lastRow - 9, 2).getValues();
    var names = [];

    for (var i = 0; i < data.length; i++) {
      var firstName = data[i][0];
      var lastName = data[i][1] || "";
      if (firstName !== "" && firstName !== null) {
        names.push((firstName + " " + lastName).trim());
      } else {
        break;
      }
    }
    return names;
  } catch (err) {
    return [];
  }
}

function getStudentByName(fullName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Student Management") || ss.getSheets()[0];
    var lastRow = sheet.getLastRow();
    if (lastRow < 10) return { found: false };

    var data = sheet.getRange(10, 2, lastRow - 9, 8).getValues();

    for (var i = 0; i < data.length; i++) {
      var currentName = (data[i][0] + " " + (data[i][1] || "")).trim();
      if (currentName === fullName) {
        return {
          found: true,
          rowIndex: 10 + i,
          firstName: data[i][0],
          lastName: data[i][1],
          phone: data[i][2] ? data[i][2].toString() : "",
          village: data[i][3],
          sangkat: data[i][4],
          district: data[i][5],
          province: data[i][6],
          checklist: data[i][7]
        };
      }
    }
    return { found: false };
  } catch (err) {
    return { found: false, error: err.toString() };
  }
}

function processForm(formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Student Management") || ss.getSheets()[0];

    var fullName = (formData.firstName + " " + formData.lastName).trim();
    var existing = getStudentByName(fullName);

    var cleanPhone = (formData.phone || "").toString().replace(/[^0-9]/g, '');
    var formattedPhone = cleanPhone ? "'" + cleanPhone : "";

    var village = formData.village || "";
    var sangkat = formData.sangkat || "";
    var district = formData.district || "";
    var province = formData.province || "";
    var checklist = formData.checklist || "";

    if (existing.found) {
      sheet.getRange(existing.rowIndex, 2, 1, 8).setValues([[
        formData.firstName,
        formData.lastName,
        formattedPhone,
        village,
        sangkat,
        district,
        province,
        checklist
      ]]);
      return { result: "success", message: "កែប្រែទិន្នន័យសិស្សបានជោគជ័យ!" };
    } else {
      var lastRow = sheet.getLastRow();
      var filledRows = 0;
      if (lastRow >= 10) {
        var colBValues = sheet.getRange(10, 2, lastRow - 9, 1).getValues();
        for (var i = 0; i < colBValues.length; i++) {
          if (colBValues[i][0] !== "") filledRows++; else break;
        }
      }
      
      var nextRow = 10 + filledRows;
      var nNumber = filledRows + 1;

      sheet.getRange(nextRow, 1, 1, 9).setValues([[
        nNumber, 
        formData.firstName, 
        formData.lastName, 
        formattedPhone, 
        village, 
        sangkat, 
        district, 
        province, 
        checklist
      ]]);

      return { result: "success", message: "បញ្ចូលទិន្នន័យសិស្សថ្មីបានជោគជ័យ!" };
    }
  } catch (err) {
    return { result: "error", error: err.toString() };
  }
}

function deleteStudentData(fullName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Student Management") || ss.getSheets()[0];
    var existing = getStudentByName(fullName);

    if (existing.found) {
      sheet.deleteRow(existing.rowIndex);
      
      var lastRow = sheet.getLastRow();
      if (lastRow >= 10) {
        var colBValues = sheet.getRange(10, 2, lastRow - 9, 1).getValues();
        var count = 0;
        for (var i = 0; i < colBValues.length; i++) {
          if (colBValues[i][0] !== "") count++; else break;
        }
        
        if (count > 0) {
          var newIndexes = [];
          for (var j = 1; j <= count; j++) {
            newIndexes.push([j]);
          }
          sheet.getRange(10, 1, count, 1).setValues(newIndexes);
        }
      }

      return { result: "success", message: "លុបទិន្នន័យសិស្សជោគជ័យ!" };
    } else {
      return { result: "error", error: "រកមិនឃើញឈ្មោះសិស្សនេះទេ!" };
    }
  } catch (err) {
    return { result: "error", error: err.toString() };
  }
}

/**
 * ------------------------------------------------------------------
 * 3. POINTS ENTRY (points.html)
 * ------------------------------------------------------------------
 */
function getStudentNames() {
  return getAllStudentList();
}

function getStudentPointsByNameAndMonth(name, month) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getPointsSheetByMonth(ss, month);
    var lastRow = sheet.getLastRow();
    if (lastRow < 10) return { found: false };

    // Column B is Name (col 2), Columns C to I are the 7 subjects (width 7)
    var data = sheet.getRange(10, 2, lastRow - 9, 8).getValues();

    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === name) {
        return {
          found: true,
          rowIndex: 10 + i,
          khmer: data[i][1],
          math: data[i][2],
          chaengSectrey: data[i][3],
          sorsevTamAn: data[i][4],
          amnanh: data[i][5],
          aksorPchong: data[i][6],
          kounu: data[i][7]
        };
      }
    }
    return { found: false };
  } catch (err) {
    return { found: false, error: err.toString() };
  }
}

function processPointsForm(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var period = data.month || 1;
    var sheet = getPointsSheetByMonth(ss, period);
    
    var khmer = parseFloat(data.khmer) || 0;
    var math = parseFloat(data.math) || 0;
    var chaengSectrey = parseFloat(data.chaengSectrey) || 0;
    var sorsevTamAn = parseFloat(data.sorsevTamAn) || 0;
    var amnanh = parseFloat(data.amnanh) || 0;
    var aksorPchong = parseFloat(data.aksorPchong) || 0;
    var kounu = parseFloat(data.kounu) || 0;
    
    var total = khmer + math + chaengSectrey + sorsevTamAn + amnanh + aksorPchong + kounu;

    var existing = getStudentPointsByNameAndMonth(data.studentName, period);
    var periodLabel = (period === "S1" ? "ឆមាសទី១" : (period === "S2" ? "ឆមាសទី២" : "ខែ " + period));

    if (existing.found) {
      // Columns C to J (cols 3 to 10) represent the 7 subjects + Total
      sheet.getRange(existing.rowIndex, 3, 1, 8).setValues([[
        khmer, math, chaengSectrey, sorsevTamAn, amnanh, aksorPchong, kounu, total
      ]]);
      return { result: "success", message: "កែប្រែពិន្ទុ " + periodLabel + " រួចរាល់!" };
    } else {
      var lastRow = sheet.getLastRow();
      var filledRows = 0;
      
      if (lastRow >= 10) {
        var colBValues = sheet.getRange(10, 2, lastRow - 9, 1).getValues();
        for (var i = 0; i < colBValues.length; i++) {
          if (colBValues[i][0] !== "") filledRows++; else break;
        }
      }
      
      var nextRow = 10 + filledRows;
      var nNumber = filledRows + 1;
      // Columns A to J (cols 1 to 10)
      sheet.getRange(nextRow, 1, 1, 10).setValues([[
        nNumber, data.studentName, khmer, math, chaengSectrey, sorsevTamAn, amnanh, aksorPchong, kounu, total
      ]]);
      return { result: "success", message: "បញ្ចូលពិន្ទុ " + periodLabel + " ថ្មីបានជោគជ័យ!" };
    }
  } catch (err) {
    return { result: "error", error: err.toString() };
  }
}

function deleteStudentPoints(studentName, month) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getPointsSheetByMonth(ss, month);
    var existing = getStudentPointsByNameAndMonth(studentName, month);
    var periodLabel = (month === "S1" ? "ឆមាសទី១" : (month === "S2" ? "ឆមាសទី២" : "ខែ " + month));

    if (existing.found) {
      sheet.deleteRow(existing.rowIndex);
      
      var lastRow = sheet.getLastRow();
      if (lastRow >= 10) {
        var colBValues = sheet.getRange(10, 2, lastRow - 9, 1).getValues();
        var count = 0;
        for (var i = 0; i < colBValues.length; i++) {
          if (colBValues[i][0] !== "") count++; else break;
        }
        
        if (count > 0) {
          var newIndexes = [];
          for (var j = 1; j <= count; j++) {
            newIndexes.push([j]);
          }
          sheet.getRange(10, 1, count, 1).setValues(newIndexes);
        }
      }

      return { result: "success", message: "លុបទិន្នន័យ " + periodLabel + " រួចរាល់!" };
    } else {
      return { result: "error", error: "រកមិនឃើញឈ្មោះសិស្សនេះក្នុង " + periodLabel + " ទេ!" };
    }
  } catch (err) {
    return { result: "error", error: err.toString() };
  }
}

/**
 * ------------------------------------------------------------------
 * 4. RANKING & ANALYSIS (analysis.html)
 * ------------------------------------------------------------------
 */
function getAnalysisData(month) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getPointsSheetByMonth(ss, month || 1);
    var lastRow = sheet.getLastRow();
    
    if (lastRow < 10) return [];

    var data = sheet.getRange(10, 2, lastRow - 9, 8).getValues();
    var result = [];

    for (var i = 0; i < data.length; i++) {
      var name = data[i][0];
      if (name && name.toString().trim() !== "") {
        result.push({
          name: name.toString().trim(),
          khmer: Number(data[i][1]) || 0,
          math: Number(data[i][2]) || 0,
          chaengSectrey: Number(data[i][3]) || 0,
          sorsevTamAn: Number(data[i][4]) || 0,
          amnanh: Number(data[i][5]) || 0,
          aksorPchong: Number(data[i][6]) || 0,
          kounu: Number(data[i][7]) || 0
        });
      }
    }
    return result;
  } catch (err) {
    return [];
  }
}
