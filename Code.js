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
  return data.slice(1).map(function(row, idx) {
    var obj = { _rowIndex: idx + 2 };
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
      // Add new row at the top (Row 2)
      // Insert a blank row before row 2 (header is row 1)
      sheet.insertRowBefore(2);
      
      // Get formulas from the pushed-down data row (which is now row 3)
      var formulas = [];
      if (sheet.getLastRow() >= 3) {
        formulas = sheet.getRange(3, 1, 1, headers.length).getFormulas()[0];
      }
      
      headers.forEach(function(header, index) {
        var formula = formulas[index];
        if (formula) {
          // Translate formula from row 3 references to row 2 references
          var newFormula = copyFormulaForNewRow(formula, 3, 2);
          rowValues.push(newFormula);
        } else if (header === 'Month') {
          rowValues.push(monthVal);
        } else if (header === 'Year') {
          rowValues.push(yearVal);
        } else {
          rowValues.push(rowData[header] !== undefined ? rowData[header] : "");
        }
      });
      
      // Write the new values to row 2
      sheet.getRange(2, 1, 1, headers.length).setValues([rowValues]);
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

function deleteTransactionRow(rowIndex) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    
    if (rowIndex) {
      sheet.deleteRow(rowIndex);
      return { success: true };
    }
    return { success: false, error: "Missing row index." };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function ensureCompanyAndGroup(group, country, company) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Ensure Company exists in dim_company
  if (company && company.trim()) {
    var companySheet = ss.getSheetByName("dim_company");
    if (companySheet) {
      var companyData = companySheet.getDataRange().getValues();
      var existingCompanies = companyData.slice(1).map(function(row) {
        return String(row[0]).trim().toLowerCase();
      });
      if (existingCompanies.indexOf(company.trim().toLowerCase()) === -1) {
        companySheet.appendRow([company.trim()]);
      }
    }
  }
  
  // 2. Ensure Group exists in dim_group
  if (group && group.trim()) {
    var groupSheet = ss.getSheetByName("dim_group");
    if (groupSheet) {
      var groupData = groupSheet.getDataRange().getValues();
      var existingGroups = groupData.slice(1).map(function(row) {
        return String(row[0]).trim().toLowerCase();
      });
      if (existingGroups.indexOf(group.trim().toLowerCase()) === -1) {
        groupSheet.appendRow([group.trim(), country ? country.trim() : "", company ? company.trim() : ""]);
      }
    }
  }
}

function addMemberMetadata(memberData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // If inline group details are provided, ensure they exist first
    if (memberData.new_group_details) {
      var gd = memberData.new_group_details;
      ensureCompanyAndGroup(gd.group, gd.country, gd.company);
    }
    
    var sheet = ss.getSheetByName("dim_member");
    if (!sheet) {
      return { success: false, error: "dim_member sheet not found" };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var lastRowIdx = sheet.getLastRow();
    
    // Read formulas from the preceding row if it exists
    var formulas = [];
    if (lastRowIdx >= 2) {
      formulas = sheet.getRange(lastRowIdx, 1, 1, headers.length).getFormulas()[0];
    }
    
    var newRowIdx = lastRowIdx + 1;
    var rowValues = [];
    
    headers.forEach(function(header, index) {
      var formula = formulas[index];
      if (formula) {
        // Translate formula
        var newFormula = copyFormulaForNewRow(formula, lastRowIdx, newRowIdx);
        rowValues.push(newFormula);
      } else if (header === 'country') {
        rowValues.push("=VLOOKUP(D" + newRowIdx + ", dim_group!A:C, 2, FALSE)");
      } else if (header === 'company') {
        rowValues.push("=VLOOKUP(D" + newRowIdx + ", dim_group!A:C, 3, FALSE)");
      } else {
        var val = memberData[header];
        if (val === undefined || val === null) {
          rowValues.push("");
        } else {
          rowValues.push(val);
        }
      }
    });
    
    sheet.appendRow(rowValues);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function addGroupMetadata(groupData) {
  try {
    ensureCompanyAndGroup(groupData.group, groupData.country, groupData.company);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function updateMemberMetadata(rowIndex, memberData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // If inline group details are provided, ensure they exist first
    if (memberData.new_group_details) {
      var gd = memberData.new_group_details;
      ensureCompanyAndGroup(gd.group, gd.country, gd.company);
    }
    
    var sheet = ss.getSheetByName("dim_member");
    if (!sheet) {
      return { success: false, error: "dim_member sheet not found" };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    // Read formulas from row rowIndex
    var existingFormulas = sheet.getRange(rowIndex, 1, 1, headers.length).getFormulas()[0];
    
    var rowValues = [];
    headers.forEach(function(header, index) {
      var formula = existingFormulas[index];
      if (formula) {
        rowValues.push(formula);
      } else if (header === 'country') {
        rowValues.push("=VLOOKUP(D" + rowIndex + ", dim_group!A:C, 2, FALSE)");
      } else if (header === 'company') {
        rowValues.push("=VLOOKUP(D" + rowIndex + ", dim_group!A:C, 3, FALSE)");
      } else {
        var val = memberData[header];
        rowValues.push(val !== undefined && val !== null ? val : "");
      }
    });
    
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowValues]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function deleteMemberMetadata(rowIndex) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("dim_member");
    if (!sheet) return { success: false, error: "dim_member sheet not found" };
    
    if (rowIndex) {
      sheet.deleteRow(rowIndex);
      return { success: true };
    }
    return { success: false, error: "Missing row index" };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function updateGroupMetadata(rowIndex, groupData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var groupSheet = ss.getSheetByName("dim_group");
    if (!groupSheet) return { success: false, error: "dim_group sheet not found" };
    
    // First, save company to dim_company if new
    var company = groupData.company;
    if (company && company.trim()) {
      var companySheet = ss.getSheetByName("dim_company");
      if (companySheet) {
        var companyData = companySheet.getDataRange().getValues();
        var existingCompanies = companyData.slice(1).map(function(row) {
          return String(row[0]).trim().toLowerCase();
        });
        if (existingCompanies.indexOf(company.trim().toLowerCase()) === -1) {
          companySheet.appendRow([company.trim()]);
        }
      }
    }
    
    var oldGroup = groupSheet.getRange(rowIndex, 1).getValue();
    var newGroup = groupData.group.trim();
    
    // Write new values
    groupSheet.getRange(rowIndex, 1, 1, 3).setValues([[newGroup, groupData.country, company]]);
    
    // Update members referencing old group
    if (oldGroup && newGroup && oldGroup.toLowerCase() !== newGroup.toLowerCase()) {
      var memberSheet = ss.getSheetByName("dim_member");
      if (memberSheet) {
        var memberData = memberSheet.getDataRange().getValues();
        var memberHeaders = memberData[0];
        var groupColIdx = memberHeaders.indexOf("group") + 1;
        if (groupColIdx > 0) {
          for (var i = 1; i < memberData.length; i++) {
            if (String(memberData[i][groupColIdx - 1]).toLowerCase() === oldGroup.toLowerCase()) {
              memberSheet.getRange(i + 1, groupColIdx).setValue(newGroup);
            }
          }
        }
      }
    }
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function deleteGroupMetadata(rowIndex) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var groupSheet = ss.getSheetByName("dim_group");
    if (!groupSheet) return { success: false, error: "dim_group sheet not found" };
    
    var groupName = groupSheet.getRange(rowIndex, 1).getValue();
    
    // Delete group row
    groupSheet.deleteRow(rowIndex);
    
    // Update members to group "N/A"
    if (groupName) {
      var memberSheet = ss.getSheetByName("dim_member");
      if (memberSheet) {
        var memberData = memberSheet.getDataRange().getValues();
        var memberHeaders = memberData[0];
        var groupColIdx = memberHeaders.indexOf("group") + 1;
        if (groupColIdx > 0) {
          for (var i = 1; i < memberData.length; i++) {
            if (String(memberData[i][groupColIdx - 1]).toLowerCase() === groupName.toLowerCase()) {
              memberSheet.getRange(i + 1, groupColIdx).setValue("N/A");
            }
          }
        }
      }
    }
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}