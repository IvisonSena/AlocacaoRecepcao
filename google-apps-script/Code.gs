var NOTES_SHEET = 'Anotações';

function doGet() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(NOTES_SHEET);
  var notes = {};
  if (sheet && sheet.getLastRow() > 1) {
    var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5);
    var values = range.getDisplayValues();
    var backgrounds = range.getBackgrounds();
    var fontColors = range.getFontColors();
    values.forEach(function(row, index) {
      var normalizedDate = normalizeDate_(row[0]);
      if (normalizedDate && (row[1] || row[2])) {
        notes[normalizedDate] = {
          morning: turnFromSheet_(row[1], row[3], backgrounds[index][1], fontColors[index][1]),
          afternoon: turnFromSheet_(row[2], row[4], backgrounds[index][2], fontColors[index][2])
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
      sheet.getRange(row, 4, 1, 2).setValues([[JSON.stringify(morning.lines), JSON.stringify(afternoon.lines)]]);
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
  sheet.getRange('A1:E1').setValues([['Data', 'Manhã', 'Tarde', 'Estilos Manhã', 'Estilos Tarde']]);
  if (!sheet.isColumnHiddenByUser(4)) sheet.hideColumns(4, 2);
  return sheet;
}

function normalizeTurn_(value) {
  value = value || {};
  var fallback = {
    background: validColor_(value.background, '#ffffff'),
    color: validColor_(value.color, '#012e73'),
    fontSize: validFontSize_(value.fontSize, 11)
  };
  var source = Array.isArray(value.lines) && value.lines.length ? value.lines : String(value.text || '').split('\n').filter(function(line) { return line.trim(); });
  var lines = source.map(function(line) {
    if (typeof line === 'string') line = { text: line };
    return {
      text: String(line.text || '').trim(),
      background: validColor_(line.background, fallback.background),
      color: validColor_(line.color, fallback.color),
      fontSize: validFontSize_(line.fontSize, fallback.fontSize)
    };
  }).filter(function(line) { return line.text; });
  return {
    text: lines.map(function(line) { return line.text; }).join('\n'),
    background: lines.length ? lines[0].background : fallback.background,
    color: lines.length ? lines[0].color : fallback.color,
    fontSize: fallback.fontSize,
    lines: lines
  };
}

function turnFromSheet_(text, serializedLines, background, color) {
  var lines = [];
  try {
    var parsed = JSON.parse(serializedLines || '[]');
    if (Array.isArray(parsed)) lines = parsed;
  } catch (error) {}
  var currentText = String(text || '');
  if (!lines.length || lines.map(function(line) { return String(line.text || ''); }).join('\n') !== currentText) {
    lines = currentText.split('\n').filter(function(line) { return line.trim(); }).map(function(line, index) {
      var previous = lines[index] || {};
      return { text: line, background: previous.background || background, color: previous.color || color, fontSize: previous.fontSize || 11 };
    });
  }
  return normalizeTurn_({ text: currentText, background: background, color: color, lines: lines });
}

function validColor_(value, fallback) {
  var color = String(value || '').trim().toLowerCase();
  return /^#[0-9a-f]{6}$/.test(color) ? color : fallback;
}

function validFontSize_(value, fallback) {
  var size = Number(value || fallback || 11);
  return Math.min(22, Math.max(7, isFinite(size) ? size : 11));
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
