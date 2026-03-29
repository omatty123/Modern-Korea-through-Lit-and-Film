function doPost(e) {
  var payload = JSON.parse(e.postData.contents || '{}');
  var sheetName = payload.sheetName || 'Attendance';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  var headers = ['course', 'sessionId', 'sessionDate', 'sessionLabel', 'topic', 'studentId', 'studentName', 'status', 'notes', 'syncedAt'];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }

  var existing = sheet.getDataRange().getValues();
  for (var i = existing.length; i >= 2; i--) {
    if (existing[i - 1][0] === payload.course && existing[i - 1][1] === payload.sessionId) {
      sheet.deleteRow(i);
    }
  }

  var syncedAt = new Date();
  var rows = (payload.students || []).map(function(student) {
    return [
      payload.course || '',
      payload.sessionId || '',
      payload.sessionDate || '',
      payload.sessionLabel || '',
      payload.topic || '',
      student.id || '',
      student.name || '',
      student.status || '',
      payload.notes || '',
      syncedAt
    ];
  });

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, rowsWritten: rows.length }))
    .setMimeType(ContentService.MimeType.JSON);
}
