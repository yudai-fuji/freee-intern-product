//　複数の機能から共通して使う処理をまとめる

function getSheet(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('シートが見つかりません: ' + sheetName);
  }
  return sheet;
}

// 設定更新・データ移行も打刻と同じScriptLockを使う。呼び出しを入れ子にしない。
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
