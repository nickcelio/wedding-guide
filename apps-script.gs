// Wedding guide reservations backend — Google Apps Script
// Stores shuttle & Catarina reservations in the attached Google Sheet.
// Admin code is required (server-side) for edit/delete.

var ADMIN_CODE = '1916';
var SHEET_NAME = 'Reservations';
var HEADERS = ['id', 'type', 'names', 'airline', 'flight', 'date', 'time', 'party', 'created'];

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function listAll() {
  var rows = getSheet().getDataRange().getValues();
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    var e = {};
    for (var j = 0; j < HEADERS.length; j++) e[HEADERS[j]] = String(rows[i][j]);
    out.push(e);
  }
  return out;
}

function doGet() {
  return json({ ok: true, entries: listAll() });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var req = JSON.parse(e.postData.contents);
    var sheet = getSheet();

    if (req.action === 'add') {
      var id = Utilities.getUuid();
      sheet.appendRow([
        id, String(req.type || ''), String(req.names || ''),
        String(req.airline || ''), String(req.flight || ''),
        String(req.date || ''), String(req.time || ''),
        String(req.party || ''), new Date().toISOString()
      ]);
      return json({ ok: true, id: id, entries: listAll() });
    }

    if (req.action === 'update' || req.action === 'delete') {
      if (String(req.admin) !== ADMIN_CODE) return json({ ok: false, error: 'Bad admin code' });
      var rows = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(req.id)) {
          if (req.action === 'delete') {
            sheet.deleteRow(i + 1);
          } else {
            sheet.getRange(i + 1, 3, 1, 6).setValues([[
              String(req.names || ''), String(req.airline || ''),
              String(req.flight || ''), String(req.date || ''),
              String(req.time || ''), String(req.party || '')
            ]]);
          }
          return json({ ok: true, entries: listAll() });
        }
      }
      return json({ ok: false, error: 'Entry not found' });
    }

    return json({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
