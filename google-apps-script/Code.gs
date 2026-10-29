var NOTES_SHEET = 'Anotações';

function doGet() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(NOTES_SHEET);
  var notes = {};
  if (sheet && sheet.getLastRow() > 1) {
    var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3);
    var values = range.getDisplayValues();
    var backgrounds = range.getBackgrounds();
    var fontColors = range.getFontColors();
    values.forEach(function(row, index) {
      var normalizedDate = normalizeDate_(row[0]);
      if (normalizedDate && (row[1] || row[2])) {
        notes[normalizedDate] = {
          morning: { text: row[1], background: backgrounds[index][1], color: fontColors[index][1] },
          afternoon: { text: row[2], background: backgrounds[index][2], color: fontColors[index][2] }
        };
      }
    });
  }
  return json_({ ok: true, notes: notes });
}

function doPost(event) {
  var payload = JSON.parse((event && event.postData && event.postData.contents) || '{}');
  var date = String(payload.date || '').trim();
  var morning = normalizeTurn_(payload.morning || (payload.note ? { text: payload.note } : {}));
  var afternoon = normalizeTurn_(payload.afternoon || {});
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return json_({ ok: false, error: 'Data inválida' });

  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    var sheet = getNotesSheet_();
    var lastRow = sheet.getLastRow();
    var dates = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().map(function(row) { return normalizeDate_(row[0]); }) : [];
    var index = dates.indexOf(date);
    var row = index >= 0 ? index + 2 : lastRow + 1;
    if (!morning.text && !afternoon.text && index >= 0) {
      sheet.deleteRow(row);
    } else if (morning.text || afternoon.text) {
      var parts = date.split('-').map(Number);
      sheet.getRange(row, 1).setValue(new Date(parts[0], parts[1] - 1, parts[2])).setNumberFormat('dd/MM/yyyy');
      var turns = sheet.getRange(row, 2, 1, 2);
      turns.setValues([[morning.text, afternoon.text]]).setWrap(true).setVerticalAlignment('top');
      turns.setBackgrounds([[morning.background, afternoon.background]]);
      turns.setFontColors([[morning.color, afternoon.color]]);
      sheet.autoResizeRows(row, 1);
    }
    return json_({ ok: true, date: date, morning: morning, afternoon: afternoon });
  } finally {
    lock.releaseLock();
  }
}

function getNotesSheet_() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName(NOTES_SHEET);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(NOTES_SHEET);
  }
  sheet.getRange('A1:C1').setValues([['Data', 'Manhã', 'Tarde']]);
  return sheet;
}

function normalizeTurn_(value) {
  return {
    text: String((value && value.text) || '').replace(/\s+$/, ''),
    background: validColor_(value && value.background, '#ffffff'),
    color: validColor_(value && value.color, '#012e73')
  };
}

function validColor_(value, fallback) {
  var color = String(value || '').trim().toLowerCase();
  return /^#[0-9a-f]{6}$/.test(color) ? color : fallback;
}

function normalizeDate_(value) {
  var text = String(value || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  var match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  return match ? match[3] + '-' + ('0' + match[2]).slice(-2) + '-' + ('0' + match[1]).slice(-2) : '';
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
