// ═══════════════════════════════════════════════════════════════════
// MEDIADORES LISBOA — Google Apps Script Backend
// Cole este código em: sheets.new → Extensões → Apps Script
// Depois: Implementar → Nova implementação → App Web → Qualquer pessoa
// ═══════════════════════════════════════════════════════════════════

const SHEET_NAME = 'Visitas';

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['idx', 'name', 'visited', 'notes', 'updated_at', 'updated_by']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  return sheet;
}

function doGet(e) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
  return ContentService
    .createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();

    // Find existing row for this idx
    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == body.idx) { rowIdx = i + 1; break; }
    }

    const now = new Date().toLocaleString('pt-PT', {timeZone: 'Europe/Lisbon'});
    const rowData = [body.idx, body.name || '', body.visited, body.notes || '', now, body.user || 'app'];

    if (rowIdx > 0) {
      sheet.getRange(rowIdx, 1, 1, 6).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    // Color row green if visited
    const lastRow = rowIdx > 0 ? rowIdx : sheet.getLastRow();
    const color = body.visited ? '#d4f7dc' : '#ffffff';
    sheet.getRange(lastRow, 1, 1, 6).setBackground(color);

    return ContentService
      .createTextOutput(JSON.stringify({ok: true}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ok: false, error: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
