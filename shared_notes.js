// 共有メモ（退勤時に他の従業員へ残す業務連絡・助言）を管理する
// 「1つの共有メモ × 受信者1人」を1行として保存し、受信者ごとに初回表示日を個別管理する

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

/**
 * 退勤しようとしている本人を除いた、共有先に選べる有効な従業員の一覧を返す
 * 店長は従業員名簿に紐づかない別ログインのため、getEmployees() の時点で対象に含まれない
 * @param {*} empId 共有メモを送ろうとしている（退勤しようとしている）従業員番号
 * @param {string} authToken ログイントークン
 * @return {Array} [{ id, name, number }, ...]
 */
function getSharedNoteRecipientCandidates(empId, authToken) {
  requireLoggedIn_(authToken);

  const employees = getEmployees();
  const candidates = [];
  for (let i = 0; i < employees.length; i++) {
    if (String(employees[i].id) === String(empId)) {
      continue; // 本人は共有先に選べない
    }
    candidates.push(employees[i]);
  }
  return candidates;
}

/**
 * 共有メモを保存してから退勤を記録する（確認画面の「退勤」を押したときに呼ぶ）
 * 保存に失敗した場合は退勤処理を行わない。保存が成功したあとに退勤時刻を記録することで、
 * 「共有メモが保存できていないのに退勤済みに見える」状態を防ぐ
 * @param {*} empId 退勤する従業員番号
 * @param {Array} recipientIds 共有先の従業員番号の配列
 * @param {string} noteText 共有メモ本文
 * @param {string} authToken ログイントークン
 * @return {string} recordClock() が返す退勤メッセージ
 */
function saveSharedNoteAndClockOut(empId, recipientIds, noteText, authToken) {
  requireLoggedIn_(authToken);

  // 削除済み・名簿にない従業員の打刻を受け付けないのと同じ考え方で、送信者も確認する
  const nameByEmpId = getActiveEmployeeNameMap();
  const senderName = nameByEmpId[String(empId)];
  if (!senderName) {
    throw new Error('この従業員は選べません。画面を再読み込みしてください');
  }

  noteText = (noteText === null || noteText === undefined) ? '' : String(noteText).trim();

  const uniqueRecipientIds = [];
  const seen = {};
  for (let i = 0; i < (recipientIds || []).length; i++) {
    const recipientId = String(recipientIds[i]);
    if (recipientId === String(empId) || seen[recipientId]) {
      continue; // 本人・重複指定は無視する
    }
    if (!nameByEmpId[recipientId]) {
      continue; // 削除済み・無効な従業員は無視する
    }
    seen[recipientId] = true;
    uniqueRecipientIds.push(recipientId);
  }

  if (uniqueRecipientIds.length === 0 || noteText === '') {
    throw new Error('共有先を1人以上選び、共有メモを入力してください');
  }

  saveSharedNoteRows_(empId, senderName, uniqueRecipientIds, nameByEmpId, noteText);

  // 共有メモの保存が終わってから、確認画面の「退勤」を押した時刻で退勤を記録する
  return recordClock(empId, '退勤');
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
      '' // 初回表示日は未表示の間は空欄のまま
    ]);
  }

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, SHARED_NOTE_SHEET_HEADER.length).setValues(rows);
}

/**
 * 出勤に成功した直後に呼ぶ。まだ一度も表示していない共有メモを「今日はじめて表示した」ことにしてから返す
 * @param {*} empId 出勤した従業員番号
 * @return {Array} その日表示してよい共有メモ [{ senderName, body, createdAtLabel }, ...]
 */
function getSharedNotesForClockIn(empId) {
  return collectSharedNotesForDisplay_(empId, true);
}

/**
 * 従業員一覧から名前を選んだときに呼ぶ。まだ一度も表示していない共有メモは表示対象にしない
 * （表示は出勤時が起点。ここでは、その日すでに表示済みの共有メモだけを返す）
 * @param {*} empId 選択された従業員番号
 * @return {Array} その日表示してよい共有メモ [{ senderName, body, createdAtLabel }, ...]
 */
function getSharedNotesForSelection(empId) {
  return collectSharedNotesForDisplay_(empId, false);
}

/**
 * 受信者1人分の共有メモを集める。呼ぶたびに、初回表示日の翌日以降になった行は削除する
 * @param {*} empId 受信者の従業員番号
 * @param {boolean} revealPending true のときだけ、未表示の共有メモを「今日はじめて表示した」ことにして含める
 * @return {Array} 画面に出す共有メモ（送信日時が古い順）
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
  const rowsToReveal = []; // 今回はじめて表示する行（シート上の行番号）
  const rowsToDelete = []; // 表示期限が過ぎたので削除する行（シート上の行番号）

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const recipientId = row[3];
    if (recipientId === '' || String(recipientId) !== String(empId)) {
      continue;
    }

    const sheetRowNumber = i + 2; // ヘッダー行の分だけ配列の添字より2大きい
    const firstShownAt = row[7];
    const firstShownKey = (firstShownAt instanceof Date)
      ? Utilities.formatDate(firstShownAt, TIME_ZONE, 'yyyy/MM/dd')
      : '';

    if (firstShownKey === '') {
      // まだ一度も出勤で表示していない共有メモ。出勤時だけ、ここで初表示にする
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

    // 初回表示日の翌日以降になったので、もう表示しない（不要になったデータとして削除する）
    rowsToDelete.push(sheetRowNumber);
  }

  // 行番号がずれる削除より先に、初表示の書き込みを済ませる
  for (let i = 0; i < rowsToReveal.length; i++) {
    sheet.getRange(rowsToReveal[i], 8).setValue(new Date(todayKey));
  }
  // 行番号が大きい行から順に削除し、これから消す行の番号が途中でずれないようにする
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
