function doGet() {
  var template = HtmlService.createTemplateFromFile('Index');
  
  var sheetData = getSheetData(); 
  template.allDataJson = JSON.stringify(sheetData.rows);
  template.headersJson = JSON.stringify(sheetData.headers);
  
  // Load dimension tables
  var metadata = getMetadata();
  template.metadataJson = JSON.stringify(metadata);
  
  var htmlOutput = template.evaluate()
    .setTitle('Cheki Database')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  
  return htmlOutput;
}

function getSheetData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { headers: [], rows: [] };
    
    var tz = ss.getSpreadsheetTimeZone();
    var headers = data[0];
    var rows = data.slice(1).map(function(row, idx) {
      var obj = { _rowIndex: idx + 2 }; // Header is row 1, first data row is 2
      headers.forEach(function(header, index) {
        var val = row[index];
        if (val instanceof Date) {
          obj[header] = Utilities.formatDate(val, tz, "yyyy-MM-dd");
        } else {
          obj[header] = val;
        }
      });
      return obj;
    });
    return { headers: headers, rows: rows };
  } catch (e) {
    throw new Error("Failed to read sheet data: " + e.message);
  }
}

function getMetadata() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    members: getSheetRows(ss, "dim_member"),
    companies: getSheetRows(ss, "dim_company"),
    groups: getSheetRows(ss, "dim_group"),
    colors: getSheetRows(ss, "dim_color"),
    types: getSheetRows(ss, "dim_type"),
    countries: getSheetRows(ss, "dim_country"),
    locations: getSheetRows(ss, "dim_location") || []
  };
}

function getSheetRows(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  var tz = ss.getSpreadsheetTimeZone();
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(header, index) {
      var val = row[index];
      if (val instanceof Date) {
        obj[header] = Utilities.formatDate(val, tz, "yyyy-MM-dd");
      } else {
        obj[header] = val;
      }
    });
    return obj;
  });
}

function saveTransaction(rowData, rowIndex) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    var rowValues = [];
    var dateVal = rowData['Date'];
    var monthVal = "";
    var yearVal = "";
    if (dateVal) {
      monthVal = dateVal.substring(0, 7);
      yearVal = dateVal.substring(0, 4);
    }

    if (rowIndex) {
      // Update existing row
      var existingFormulas = sheet.getRange(rowIndex, 1, 1, headers.length).getFormulas()[0];
      
      headers.forEach(function(header, index) {
        var formula = existingFormulas[index];
        if (formula) {
          rowValues.push(formula);
        } else if (header === 'Month') {
          rowValues.push(monthVal);
        } else if (header === 'Year') {
          rowValues.push(yearVal);
        } else {
          rowValues.push(rowData[header] !== undefined ? rowData[header] : "");
        }
      });
      
      sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowValues]);
    } else {
      // Add new row
      var targetRowIndex = sheet.getLastRow() + 1;
      var lastRowIndex = sheet.getLastRow();
      var lastRowFormulas = [];
      if (lastRowIndex >= 2) {
        lastRowFormulas = sheet.getRange(lastRowIndex, 1, 1, headers.length).getFormulas()[0];
      }
      
      headers.forEach(function(header, index) {
        var formula = lastRowFormulas[index];
        if (formula) {
          var newFormula = copyFormulaForNewRow(formula, lastRowIndex, targetRowIndex);
          rowValues.push(newFormula);
        } else if (header === 'Month') {
          rowValues.push(monthVal);
        } else if (header === 'Year') {
          rowValues.push(yearVal);
        } else {
          rowValues.push(rowData[header] !== undefined ? rowData[header] : "");
        }
      });
      
      sheet.getRange(targetRowIndex, 1, 1, headers.length).setValues([rowValues]);
    }
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function copyFormulaForNewRow(formula, lastRowIndex, newRowIndex) {
  if (!formula) return "";
  var regex = new RegExp('([A-Za-z]+)' + lastRowIndex + '\\b', 'g');
  return formula.replace(regex, '$1' + newRowIndex);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}