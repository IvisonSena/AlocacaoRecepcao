const NOTES_SHEET = 'Anotações';

function doGet() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(NOTES_SHEET);
  const notes = {};
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues().forEach(([date, note]) => {
      const normalizedDate = normalizeDate_(date);
      if (normalizedDate && note) notes[normalizedDate] = note;
    });
  }
  return json_({ ok: true, notes });
}

function doPost(event) {
  const payload = JSON.parse((event && event.postData && event.postData.contents) || '{}');
  const date = String(payload.date || '').trim();
  const note = String(payload.note || '').trimEnd();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return json_({ ok: false, error: 'Data inválida' });

  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const sheet = getNotesSheet_();
    const lastRow = sheet.getLastRow();
    const dates = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat().map(normalizeDate_) : [];
    const index = dates.indexOf(date);
    const row = index >= 0 ? index + 2 : lastRow + 1;
    if (!note && index >= 0) {
      sheet.deleteRow(row);
    } else if (note) {
      sheet.getRange(row, 1).setNumberFormat('@').setValue(date);
      sheet.getRange(row, 2).setValue(note).setWrap(true);
    }
    return json_({ ok: true, date, note });
  } finally {
    lock.releaseLock();
  }
}

function getNotesSheet_() {
  const spreadsheet = SpreadsheetApp.getActive();
  let sheet = spreadsheet.getSheetByName(NOTES_SHEET);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(NOTES_SHEET);
    sheet.getRange('A1:B1').setValues([['Data', 'Anotação']]);
  }
  return sheet;
}

function normalizeDate_(value) {
  const text = String(value || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  const match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  return match ? `${match[3]}-${match[2].padStart(2, '0')}-${match[1].padStart(2, '0')}` : '';
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
