function onEdit(e) {
  if (!e || !e.range) return; // safety guard

  const sheet = e.range.getSheet();
  if (sheet.getName() !== 'fact_cheki_transaction') return;

  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  if (lastRow === 0 || lastCol === 0) return;

  const headerScan = sheet
    .getRange(1, 1, Math.min(10, lastRow), lastCol)
    .getValues();

  let headerRow = -1;
  let quantityCol = -1;
  let typeCol = -1;
  let locationCol = -1;

  for (let r = 0; r < headerScan.length; r++) {
    const row = headerScan[r];

    if (row.includes('Quantity')) {
      headerRow = r + 1;
      quantityCol = row.indexOf('Quantity') + 1;
      typeCol = row.indexOf('Type') + 1;
      locationCol = row.indexOf('Location') + 1;
      break;
    }
  }

  if (headerRow === -1) return;

  const startRow = e.range.getRow();
  const numRows = e.range.getNumRows();

  for (let r = 0; r < numRows; r++) {
    const currentRow = startRow + r;
    if (currentRow <= headerRow) continue;

    // Quantity → 1
    if (quantityCol > 0) {
      const cell = sheet.getRange(currentRow, quantityCol);
      if (cell.isBlank()) cell.setValue(1);
    }

    // Type → "Cheki"
    if (typeCol > 0) {
      const cell = sheet.getRange(currentRow, typeCol);
      if (cell.isBlank()) cell.setValue('Cheki');
    }

    // Location → "Bangkok"
    if (locationCol > 0) {
      const cell = sheet.getRange(currentRow, locationCol);
      if (cell.isBlank()) cell.setValue('Bangkok');
    }
  }
}

