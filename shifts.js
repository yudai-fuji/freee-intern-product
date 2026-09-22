// シフト作成画面に残す月間カレンダーと、旧シフトデータの一度きりの削除を扱う。

/**
 * 表示年月と今日の日付だけを返す。割当・必要人数・作成状態は扱わない。
 * 年月を両方省略した場合は、店舗のタイムゾーンで今日を含む月を選ぶ。
 * @return {Object} { year, month, today }
 */
function getShiftMonth(year, month, authToken) {
  requireManagerRole_(authToken);

  const today = getTodayDateKey();
  const useCurrentMonth = (year === null || year === undefined) &&
    (month === null || month === undefined);
  const targetYear = useCurrentMonth ? Number(today.slice(0, 4)) : Number(year);
  const targetMonth = useCurrentMonth ? Number(today.slice(5, 7)) : Number(month);
  if (!Number.isInteger(targetYear) || !Number.isInteger(targetMonth) ||
      targetYear < 1900 || targetYear > 2999 || targetMonth < 1 || targetMonth > 12) {
    throw new Error('表示する年月が不正です: ' + year + '年' + month + '月');
  }

  removeLegacyShiftSheets_();
  return { 'year': targetYear, 'month': targetMonth, 'today': today };
}

/**
 * 店長がカレンダーを初めて開いたときだけ、旧シフトの2シートと全データを削除する。
 * 名前が完全一致するシートのみ対象にし、ブックごとの完了フラグで再実行を防ぐ。
 * 途中で失敗した場合は完了扱いにせず、次回の表示で残りの削除を再試行する。
 */
function removeLegacyShiftSheets_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error('対象のスプレッドシートが見つかりません');
    }
    const properties = PropertiesService.getScriptProperties();
    const completedKey = 'legacy_shift_sheets_removed_v1:' + spreadsheet.getId();
    if (properties.getProperty(completedKey) === 'true') {
      return;
    }

    const legacySheetNames = ['シフト', 'シフト必要人数'];
    for (let i = 0; i < legacySheetNames.length; i++) {
      const sheet = spreadsheet.getSheetByName(legacySheetNames[i]);
      if (!sheet) {
        continue;
      }
      if (sheet.getName() !== legacySheetNames[i]) {
        throw new Error('旧シフトシートの名前が一致しないため削除を中止しました');
      }
      spreadsheet.deleteSheet(sheet);
    }
    SpreadsheetApp.flush();
    for (let i = 0; i < legacySheetNames.length; i++) {
      if (spreadsheet.getSheetByName(legacySheetNames[i])) {
        throw new Error('旧シフトシートの削除を確認できませんでした');
      }
    }
    properties.setProperty(completedKey, 'true');
  } finally {
    lock.releaseLock();
  }
}
