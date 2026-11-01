/**
 * 指定名のシートを取得する。
 * sheetName - string
 * 戻り値 - GoogleAppsScript.Spreadsheet.Sheet
 */
function getSheet(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('シートが見つかりません: ' + sheetName);
  }
  return sheet;
}

/**
 * 共有ロック内で指定処理を実行し、書き込みを反映してロックを解放する。
 * callback - () => unknown
 * 戻り値 - unknown
 */
function withAppLock_(callback) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('他の処理が進行中です。少し待ってからもう一度お試しください');
  }
  try {
    return callback();
  } finally {
    try {
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
  }
}
