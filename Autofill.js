/**
 * Google Apps Script Triggers for Autofill & Formula Propagation
 * Sheet target: 'fact_cheki_transaction'
 */

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== 'fact_cheki_transaction') return;

  const startRow = e.range.getRow();
  const numRows = e.range.getNumRows();
  autofillRows(sheet, startRow, numRows);
}

function onChange(e) {
  if (!e) return;
  const ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  const sheet = ss.getActiveSheet();
  if (!sheet || sheet.getName() !== 'fact_cheki_transaction') return;

  // When a row is inserted or structural change happens, get active range or fallback
  const activeRange = sheet.getActiveRange();
  if (activeRange) {
    const startRow = activeRange.getRow();
    const numRows = activeRange.getNumRows();
    autofillRows(sheet, startRow, numRows);
  } else {
    // Fallback: process row 2 (top inserted row) down to last row
    autofillRows(sheet, 2, 1);
  }
}

/**
 * Main autofill function that populates default values and copies missing formulas
 * for specified row(s) in 'fact_cheki_transaction'.
 */
function autofillRows(sheet, startRow, numRows) {
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1 || lastCol === 0) return;

  // Scan top 10 rows to locate header row and column indexes
  const maxScanRows = Math.min(10, lastRow);
  const headerScan = sheet.getRange(1, 1, maxScanRows, lastCol).getValues();

  let headerRow = -1;
  let quantityCol = -1;
  let typeCol = -1;
  let locationCol = -1;
  let dateCol = -1;
  let monthCol = -1;
  let yearCol = -1;

  for (let r = 0; r < headerScan.length; r++) {
    const rowStr = headerScan[r].map(c => String(c || '').trim().toLowerCase());
    if (rowStr.includes('quantity')) {
      headerRow = r + 1;
      quantityCol = rowStr.indexOf('quantity') + 1;
      typeCol = rowStr.indexOf('type') + 1;
      locationCol = rowStr.indexOf('location') + 1;
      dateCol = rowStr.indexOf('date') + 1;
      monthCol = rowStr.indexOf('month') + 1;
      yearCol = rowStr.indexOf('year') + 1;
      break;
    }
  }

  if (headerRow === -1) return;

  // Identify a reference row containing formula definitions (prefer row below target, else row above)
  let refRowIndex = -1;
  for (let r = startRow + numRows; r <= lastRow; r++) {
    if (r > headerRow) { refRowIndex = r; break; }
  }
  if (refRowIndex === -1) {
    for (let r = startRow - 1; r > headerRow; r--) {
      if (r > headerRow) { refRowIndex = r; break; }
    }
  }

  for (let i = 0; i < numRows; i++) {
    const currentRow = startRow + i;
    if (currentRow <= headerRow) continue;

    // 1. Copy missing formulas from reference row to new row
    if (refRowIndex > 0 && refRowIndex !== currentRow) {
      const refFormulas = sheet.getRange(refRowIndex, 1, 1, lastCol).getFormulas()[0];
      for (let c = 1; c <= lastCol; c++) {
        const formula = refFormulas[c - 1];
        if (formula) {
          const targetCell = sheet.getRange(currentRow, c);
          if (targetCell.isBlank() && targetCell.getFormula() === "") {
            const refCell = sheet.getRange(refRowIndex, c);
            refCell.copyTo(targetCell, SpreadsheetApp.CopyPasteType.PASTE_FORMULAS, false);
          }
        }
      }
    }

    // 2. Quantity → Default to 1
    if (quantityCol > 0) {
      const cell = sheet.getRange(currentRow, quantityCol);
      if (cell.isBlank() && cell.getFormula() === "") cell.setValue(1);
    }

    // 3. Type → Default to "Cheki"
    if (typeCol > 0) {
      const cell = sheet.getRange(currentRow, typeCol);
      if (cell.isBlank() && cell.getFormula() === "") cell.setValue('Cheki');
    }

    // 4. Location → Default to "Bangkok"
    if (locationCol > 0) {
      const cell = sheet.getRange(currentRow, locationCol);
      if (cell.isBlank() && cell.getFormula() === "") cell.setValue('Bangkok');
    }

    // 5. Date → Auto-derive Month (YYYY-MM) and Year (YYYY) if Date present & Month/Year blank
    if (dateCol > 0) {
      const dateCellVal = sheet.getRange(currentRow, dateCol).getValue();
      if (dateCellVal) {
        let dateStr = "";
        if (dateCellVal instanceof Date) {
          dateStr = Utilities.formatDate(dateCellVal, sheet.getParent().getSpreadsheetTimeZone(), "yyyy-MM-dd");
        } else if (typeof dateCellVal === 'string' && dateCellVal.match(/^\d{4}-\d{2}-\d{2}/)) {
          dateStr = dateCellVal.substring(0, 10);
        }

        if (dateStr) {
          if (monthCol > 0) {
            const mCell = sheet.getRange(currentRow, monthCol);
            if (mCell.isBlank() && mCell.getFormula() === "") {
              mCell.setValue(dateStr.substring(0, 7));
            }
          }
          if (yearCol > 0) {
            const yCell = sheet.getRange(currentRow, yearCol);
            if (yCell.isBlank() && yCell.getFormula() === "") {
              yCell.setValue(dateStr.substring(0, 4));
            }
          }
        }
      }
    }
  }
}

/**
 * Manual test function to verify row insertion autofill
 */
function testAutofillNewRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('fact_cheki_transaction');
  if (!sheet) {
    Logger.log("Sheet fact_cheki_transaction not found!");
    return;
  }
  sheet.insertRowBefore(2);
  autofillRows(sheet, 2, 1);
  Logger.log("Autofill test completed successfully for row 2.");
}


