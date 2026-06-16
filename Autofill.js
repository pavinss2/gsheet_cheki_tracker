function onEdit(e) {
  if (!e || !e.range) return; // safety guard

  const sheet = e.range.getSheet();
  if (sheet.getName() !== '(fact_cheki_transaction') return;

  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();

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

  const editedRow = e.range.getRow();
  if (editedRow <= headerRow) return;

  // Quantity → 1
  if (quantityCol > 0) {
    const cell = sheet.getRange(editedRow, quantityCol);
    if (cell.isBlank()) cell.setValue(1);
  }

  // Type → "Cheki"
  if (typeCol > 0) {
    const cell = sheet.getRange(editedRow, typeCol);
    if (cell.isBlank()) cell.setValue('Cheki');
  }

  // Location → "Bangkok"
  if (locationCol > 0) {
    const cell = sheet.getRange(editedRow, locationCol);
    if (cell.isBlank()) cell.setValue('Bangkok');
  }
}
