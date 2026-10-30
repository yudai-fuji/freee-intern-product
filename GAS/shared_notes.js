/**
 * 共有メモシートを用意し、送信者・受信者・本文・表示日等の見出しをそろえる。
 * 戻り値 - GoogleAppsScript.Spreadsheet.Sheet
 */
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

/**
 * 共有メモの各受信者行を識別する新しいIDを作る。
 * 戻り値 - string
 */
function newSharedNoteId_() {
  return 'NOTE-' + Utilities.getUuid().replace(/-/g, '').substring(0, 12);
}

/**
 * 本人を除く有効な従業員を共有先候補として取得する。
 * empId - unknown
 * authToken - string
 * 戻り値 - ({id: unknown, name: string, number: string})[]
 */
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

/**
 * 共有メモを保存してから退勤を記録する（確認画面の「退勤」を押したときに呼ぶ） 保存に失敗した場合は退勤処理を行わない。
 * empId - unknown
 * recipientIds - unknown
 * noteText - unknown
 * authToken - unknown
 * ticket - unknown
 * 戻り値 - {message: string, clockState: {employeeId: string, date: string, state: string, allowedTypes: string[], error?: undefined または string}}
 */
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

/**
 * 共有メモを受信者1人につき1行保存する。
 * senderEmpId - unknown
 * senderName - unknown
 * recipientIds - string[]
 * nameByEmpId - unknown
 * noteText - unknown
 * 戻り値 - なし
 */
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

/**
 * 出勤に成功した直後に呼ぶ。
 * empId - unknown
 * 戻り値 - ({senderName: unknown, body: unknown, createdAtLabel: string, createdAtMs: number})[]
 */
function getSharedNotesForClockIn(empId) {
  return collectSharedNotesForDisplay_(empId, true);
}

/**
 * 従業員一覧から名前を選んだときに呼ぶ。
 * empId - string
 * 戻り値 - ({senderName: unknown, body: unknown, createdAtLabel: string, createdAtMs: number})[]
 */
function getSharedNotesForSelection(empId) {
  return collectSharedNotesForDisplay_(empId, false);
}

/**
 * 受信者1人分の共有メモを集める。
 * empId - unknown
 * revealPending - boolean
 * 戻り値 - ({senderName: unknown, body: unknown, createdAtLabel: string, createdAtMs: number})[]
 */
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

/**
 * シート1行から送信者名・本文・作成日時の表示と並び替え用の時刻を取り出す。
 * row - unknown[]
 * 戻り値 - {senderName: unknown, body: unknown, createdAtLabel: string, createdAtMs: number}
 */
function buildSharedNoteView_(row) {
  const createdAt = row[6];
  return {
    'senderName': row[2] || '',
    'body': row[5] || '',
    'createdAtLabel': (createdAt instanceof Date) ? Utilities.formatDate(createdAt, TIME_ZONE, 'M/d HH:mm') : '',
    'createdAtMs': (createdAt instanceof Date) ? createdAt.getTime() : 0
  };
}
