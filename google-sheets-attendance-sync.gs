function doPost(e) {
  var rawPayload = '';
  if (e && e.parameter && e.parameter.payload) {
    rawPayload = e.parameter.payload;
  } else if (e && e.postData && e.postData.contents) {
    rawPayload = e.postData.contents;
  }

  var payload = JSON.parse(rawPayload || '{}');
  var sheetName = payload.sheetName || 'Attendance Matrix';
  var notesSheetName = payload.notesSheetName || (sheetName + ' Notes');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var matrixSheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  var notesSheet = ss.getSheetByName(notesSheetName) || ss.insertSheet(notesSheetName);

  if (payload.test) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, test: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  ensureMatrixSheet(matrixSheet);
  ensureNotesSheet(notesSheet);

  var sessionDate = payload.sessionDate || payload.sessionId || '';
  var sessionHeader = payload.sessionShortLabel || payload.sessionDate || payload.sessionLabel || payload.sessionId || 'Session';
  var topic = payload.topic || '';
  var students = payload.students || [];
  var sessionColumn = ensureSessionColumn(matrixSheet, sessionDate, sessionHeader, topic);
  var rowMap = ensureStudentRows(matrixSheet, students);

  for (var i = 0; i < students.length; i++) {
    var student = students[i];
    var row = rowMap[student.name || student.id];
    var cell = matrixSheet.getRange(row, sessionColumn);
    cell.setValue(statusCode(student.status));
    applyStatusFormat(cell, student.status);
  }

  upsertNotesRow(notesSheet, payload);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, rowsWritten: students.length, sheet: sheetName }))
    .setMimeType(ContentService.MimeType.JSON);
}

function ensureMatrixSheet(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1).setValue('Student');
  }
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  sheet.getRange(1, 1).setFontWeight('bold');
  sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), Math.max(sheet.getLastColumn(), 1)).setVerticalAlignment('middle');
  sheet.setColumnWidth(1, 180);
}

function ensureNotesSheet(sheet) {
  var headers = ['sessionDate', 'topic', 'present', 'absent', 'notes', 'syncedAt'];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  sheet.setFrozenRows(1);
}

function ensureSessionColumn(sheet, sessionDate, sessionHeader, topic) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headerValues = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var match = headerValues.indexOf(sessionDate);
  var column = match >= 0 ? match + 1 : lastCol + 1;

  if (column === 1) {
    column = 2;
  }

  sheet.getRange(1, column).setValue(sessionDate || sessionHeader);
  sheet.getRange(1, column).setNote(topic);
  sheet.getRange(1, column).setFontWeight('bold');
  sheet.getRange(1, column).setHorizontalAlignment('center');
  sheet.setColumnWidth(column, 84);
  return column;
}

function ensureStudentRows(sheet, students) {
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var ids = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(function(row) { return row[0]; }) : [];
  var rowMap = {};

  for (var i = 0; i < ids.length; i++) {
    rowMap[ids[i]] = i + 2;
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

function statusCode(status) {
  return status === 'absent' ? '●' : '';
}

function applyStatusFormat(cell, status) {
  var backgrounds = {
    absent: '#fbeaea'
  };
  var fonts = {
    absent: '#b33636'
  };

  cell.setHorizontalAlignment('center');
  cell.setFontWeight('bold');
  cell.setBackground(backgrounds[status] || '#ffffff');
  cell.setFontColor(fonts[status] || '#161616');
}

function upsertNotesRow(sheet, payload) {
  var sessionDate = payload.sessionDate || payload.sessionId || '';
  var counts = payload.counts || {};
  var rowData = [
    sessionDate,
    payload.topic || '',
    counts.present || 0,
    counts.absent || 0,
    payload.notes || '',
    new Date()
  ];

  var lastRow = sheet.getLastRow();
  var existingDates = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(function(row) { return row[0]; }) : [];
  var existingIndex = existingDates.indexOf(sessionDate);
  var targetRow = existingIndex >= 0 ? existingIndex + 2 : lastRow + 1;

  sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
}
