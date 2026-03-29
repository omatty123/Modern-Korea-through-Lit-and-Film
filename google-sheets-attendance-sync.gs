function doPost(e) {
  var rawPayload = '';
  if (e && e.parameter && e.parameter.payload) {
    rawPayload = e.parameter.payload;
  } else if (e && e.postData && e.postData.contents) {
    rawPayload = e.postData.contents;
  }

  var payload = JSON.parse(rawPayload || '{}');
  var sheetName = payload.sheetName || 'Attendance Matrix';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  if (payload.test) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, test: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  ensureHeaderRow(sheet, payload.allSessions || []);
  var students = payload.students || [];
  var sessionColumn = findSessionColumn(sheet, payload.sessionShortLabel || payload.sessionDate || payload.sessionId || '');
  var rowMap = ensureStudentRows(sheet, students);

  for (var i = 0; i < students.length; i++) {
    var student = students[i];
    var row = rowMap[student.name || student.id];
    var cell = sheet.getRange(row, sessionColumn);
    var absent = student.status === 'absent';
    cell.setValue(absent ? 'x' : '');
    cell.setHorizontalAlignment('center');
    cell.setFontWeight(absent ? 'bold' : 'normal');
    cell.setFontColor(absent ? '#b33636' : '#161616');
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, rowsWritten: students.length, sheet: sheetName, session: payload.sessionShortLabel || '' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function ensureHeaderRow(sheet, sessions) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1).setValue('Student');
  }
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  sheet.getRange(1, 1).setFontWeight('bold');
  sheet.setColumnWidth(1, 180);

  for (var i = 0; i < sessions.length; i++) {
    var col = i + 2;
    sheet.getRange(1, col).setValue(sessions[i]);
    sheet.getRange(1, col).setFontWeight('bold');
    sheet.getRange(1, col).setHorizontalAlignment('center');
    sheet.setColumnWidth(col, 68);
  }
}

function findSessionColumn(sheet, sessionLabel) {
  var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  var idx = headers.indexOf(sessionLabel);
  if (idx === -1) {
    throw new Error('Session date not found in header row: ' + sessionLabel);
  }
  return idx + 1;
}

function ensureStudentRows(sheet, students) {
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var names = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(function(row) { return row[0]; }) : [];
  var rowMap = {};

  for (var i = 0; i < names.length; i++) {
    rowMap[names[i]] = i + 2;
  }

  for (var j = 0; j < students.length; j++) {
    var student = students[j];
    var key = student.name || student.id;
    if (!rowMap[key]) {
      var newRow = sheet.getLastRow() + 1;
      sheet.getRange(newRow, 1).setValue(key);
      rowMap[key] = newRow;
    }
  }

  for (var k = 0; k < students.length; k++) {
    var item = students[k];
    var row = rowMap[item.name || item.id];
    sheet.getRange(row, 1).setValue(item.name || item.id || '');
    sheet.getRange(row, 1).setFontWeight('bold');
  }

  return rowMap;
}
