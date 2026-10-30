function ensureSharedNoteSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.SHARED_NOTES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.SHARED_NOTES);
  }

  const headerRange = sheet.getRange(1, 1, 1, SHARED_NOTE_SHEET_HEADER.length);
  const currentHeader = headerRange.getValues()[0];
  for (let i = 0; i < SHARED_NOTE_SHEET_HEADER.length; i++) {
    if (currentHeader[i] !== SHARED_NOTE_SHEET_HEADER[i]) {
      headerRange.setValues([SHARED_NOTE_SHEET_HEADER]);
      break;
    }
  }
  return sheet;
}

function newSharedNoteId_() {
  return 'NOTE-' + Utilities.getUuid().replace(/-/g, '').substring(0, 12);
}

function getSharedNoteRecipientCandidates(empId, authToken) {
  requireLoggedIn_(authToken);

  const employees = getEmployees();
  const candidates = [];
  for (let i = 0; i < employees.length; i++) {
    if (String(employees[i].id) === String(empId)) {
      continue;
    }
    candidates.push(employees[i]);
  }
  return candidates;
}

function saveSharedNoteAndClockOut(empId, recipientIds, noteText, authToken, ticket) {
  requireLoggedIn_(authToken);

  return recordClockWithValidation_(empId, '退勤', ticket, authToken, function (nameByEmpId) {
    const senderName = nameByEmpId[String(empId)];
    noteText = (noteText === null || noteText === undefined) ? '' : String(noteText).trim();

    const uniqueRecipientIds = [];
    const seen = {};
    for (let i = 0; i < (recipientIds || []).length; i++) {
      const recipientId = String(recipientIds[i]);
      if (recipientId === String(empId) || seen[recipientId]) {
        continue;
      }
      if (!nameByEmpId[recipientId]) {
        continue;
      }
      seen[recipientId] = true;
      uniqueRecipientIds.push(recipientId);
    }

    if (uniqueRecipientIds.length === 0 || noteText === '') {
      throw new Error('共有先を1人以上選び、共有メモを入力してください');
    }
    saveSharedNoteRows_(empId, senderName, uniqueRecipientIds, nameByEmpId, noteText);
  });
}

function saveSharedNoteRows_(senderEmpId, senderName, recipientIds, nameByEmpId, noteText) {
  const sheet = ensureSharedNoteSheet();
  const now = new Date();

  const rows = [];
  for (let i = 0; i < recipientIds.length; i++) {
    const recipientId = recipientIds[i];
    rows.push([
      newSharedNoteId_(),
      senderEmpId,
      senderName,
      recipientId,
      nameByEmpId[recipientId],
      noteText,
      now,
      ''
    ]);
  }

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, SHARED_NOTE_SHEET_HEADER.length).setValues(rows);
}

function getSharedNotesForClockIn(empId) {
  return collectSharedNotesForDisplay_(empId, true);
}

function getSharedNotesForSelection(empId) {
  return collectSharedNotesForDisplay_(empId, false);
}

function collectSharedNotesForDisplay_(empId, revealPending) {
  const sheet = ensureSharedNoteSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  const todayKey = getTodayDateKey();
  const values = sheet.getRange(2, 1, lastRow - 1, SHARED_NOTE_SHEET_HEADER.length).getValues();

  const notes = [];
  const rowsToReveal = [];
  const rowsToDelete = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const recipientId = row[3];
    if (recipientId === '' || String(recipientId) !== String(empId)) {
      continue;
    }

    const sheetRowNumber = i + 2;
    const firstShownAt = row[7];
    const firstShownKey = (firstShownAt instanceof Date)
      ? Utilities.formatDate(firstShownAt, TIME_ZONE, 'yyyy/MM/dd')
      : '';

    if (firstShownKey === '') {

      if (revealPending) {
        rowsToReveal.push(sheetRowNumber);
        notes.push(buildSharedNoteView_(row));
      }
      continue;
    }

    if (firstShownKey === todayKey) {
      notes.push(buildSharedNoteView_(row));
      continue;
    }

    rowsToDelete.push(sheetRowNumber);
  }

  for (let i = 0; i < rowsToReveal.length; i++) {
    sheet.getRange(rowsToReveal[i], 8).setValue(new Date(todayKey));
  }

  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }

  notes.sort(function (a, b) { return a.createdAtMs - b.createdAtMs; });
  for (let i = 0; i < notes.length; i++) {
    delete notes[i].createdAtMs;
  }
  return notes;
}

function buildSharedNoteView_(row) {
  const createdAt = row[6];
  return {
    'senderName': row[2] || '',
    'body': row[5] || '',
    'createdAtLabel': (createdAt instanceof Date) ? Utilities.formatDate(createdAt, TIME_ZONE, 'M/d HH:mm') : '',
    'createdAtMs': (createdAt instanceof Date) ? createdAt.getTime() : 0
  };
}
