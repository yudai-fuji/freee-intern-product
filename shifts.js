// シフト確認・週間編集の取得と、登録ボタンで送られた下書きの一括保存を扱う。
// 「シフト」シートを再利用し、旧データやシフト必要人数シートは削除しない。
const SHIFT_SHEET_NAME = 'シフト';
const SHIFT_FIELDS = ['employeeId', 'date', 'workType', 'startTime', 'endTime'];
const SHIFT_HEADERS = ['従業員ID', '日付', '勤務区分', '開始時刻', '終了時刻'];
const SHIFT_HEADER_ALIASES = {
  employeeId: ['従業員ID', '従業員番号'],
  date: ['日付'],
  workType: ['勤務区分'],
  startTime: ['開始時刻', '出勤時刻'],
  endTime: ['終了時刻', '退勤時刻']
};
const SHIFT_TIME_PATTERN = /^([01][0-9]|2[0-3]):[0-5][0-9]$/;
const SHIFT_CUSTOM_TIME_PATTERN = /^(1[1-9]|2[0-2]):[0-5][0-9]$/;

/** 月の保存済みシフトを返す。氏名は削除済み従業員についても表示できる。 */
function getShiftMonth(year, month, authToken) {
  requireManagerRole_(authToken);
  const today = getTodayDateKey().replace(/\//g, '-');
  const useCurrentMonth = (year === null || year === undefined) &&
    (month === null || month === undefined);
  const targetYear = useCurrentMonth ? Number(today.slice(0, 4)) : Number(year);
  const targetMonth = useCurrentMonth ? Number(today.slice(5, 7)) : Number(month);
  if (!Number.isInteger(targetYear) || !Number.isInteger(targetMonth) ||
      targetYear < 1900 || targetYear > 2999 || targetMonth < 1 || targetMonth > 12) {
    throw new Error('表示する年月が不正です: ' + year + '年' + month + '月');
  }
  const prefix = String(targetYear) + '-' + ('0' + targetMonth).slice(-2) + '-';
  const skillMap = getEmployeeSkillMap();
  const employees = readEmployeeRows_().filter(function (row) { return row.name !== ''; }).map(function (row) {
    return { id: String(row.id), name: row.name, number: formatEmployeeNumber(row.id),
      isDeleted: row.isDeleted, skills: skillMap[String(row.id)] || getDefaultSkills() };
  });
  return {
    year: targetYear,
    month: targetMonth,
    today: today,
    employees: employees,
    holidays: Object.keys(getHolidayDateSet()).map(function (date) { return date.replace(/\//g, '-'); })
      .filter(function (date) { return date.indexOf(prefix) === 0; }),
    shifts: readShiftStore_().records.filter(function (record) {
      return record.date.indexOf(prefix) === 0;
    }).map(shiftPublicRecord_)
  };
}

/** 省略時は店舗の今日を含む日曜からの7日間。有効な全従業員を返す。 */
function getShiftWeek(weekStart, authToken) {
  requireManagerRole_(authToken);
  const today = getTodayDateKey().replace(/\//g, '-');
  let start = weekStart;
  if (start === null || start === undefined) {
    const todayDate = parseShiftDate_(today);
    todayDate.setUTCDate(todayDate.getUTCDate() - todayDate.getUTCDay());
    start = shiftDateText_(todayDate);
  }
  const startDate = parseShiftDate_(start);
  if (!startDate) throw new Error('表示する週の日付が不正です');
  const endDate = new Date(startDate.getTime());
  endDate.setUTCDate(endDate.getUTCDate() + 6);
  const end = shiftDateText_(endDate);
  return {
    weekStart: start,
    today: today,
    employees: getEmployees().map(function (employee) {
      return { id: String(employee.id), name: employee.name, number: employee.number };
    }),
    shifts: readShiftStore_().records.filter(function (record) {
      return record.date >= start && record.date <= end;
    }).map(shiftPublicRecord_)
  };
}

/**
 * 全変更を検証してから、従業員IDと日付が一致する行を更新する。
 * ロック中に既存データを読み直すので、再送でも同じセルを重複追加しない。
 * changes: [{ employeeId, date, operation: 'upsert'|'delete', workType, startTime, endTime }]
 */
function saveShiftDrafts(changes, authToken) {
  requireManagerRole_(authToken);
  const normalized = validateShiftChanges_(changes);
  if (normalized.length === 0) return { savedCount: 0, shifts: [] };
  return withAppLock_(function () {
    const activeIds = Object.create(null);
    getEmployees().forEach(function (employee) { activeIds[String(employee.id)] = true; });
    normalized.forEach(function (change) {
      if (!activeIds[change.employeeId]) throw new Error('無効な従業員が含まれています。従業員一覧を更新してください');
    });

    // 未知のヘッダーや既存の重複は、書き換える前にエラーとして知らせる。
    const store = readShiftStore_();
    const existing = Object.create(null);
    store.records.forEach(function (record) { existing[shiftRecordKey_(record)] = record; });
    const writes = [];
    let nextRow = store.sheet ? Math.max(store.sheet.getLastRow(), 1) + 1 : 2;
    normalized.forEach(function (change) {
      const record = existing[shiftRecordKey_(change)];
      if (change.operation === 'delete') {
        if (record) writes.push({ rowNumber: record.rowNumber, values: ['', '', '', '', ''] });
      } else {
        writes.push({ rowNumber: record ? record.rowNumber : nextRow++, values: SHIFT_FIELDS.map(function (field) {
          return change[field];
        }) });
      }
    });

    if (writes.length > 0) {
      prepareShiftStoreForWrite_(store);
      writeShiftChanges_(store, writes);
    }
    return {
      savedCount: normalized.length,
      shifts: normalized.filter(function (change) { return change.operation === 'upsert'; }).map(shiftPublicRecord_)
    };
  });
}

function validateShiftChanges_(changes) {
  if (!Array.isArray(changes)) throw new Error('登録するシフトの形式が不正です');
  const seen = Object.create(null);
  return changes.map(function (change) {
    if (!change || typeof change !== 'object' ||
        (typeof change.employeeId !== 'string' && typeof change.employeeId !== 'number')) {
      throw new Error('登録する従業員IDが不正です');
    }
    const employeeId = String(change.employeeId).trim();
    const date = parseShiftDate_(change.date);
    if (!employeeId || !date) throw new Error('登録する従業員IDまたは日付が不正です');
    if (date.getUTCDay() === 1) throw new Error('月曜日は定休日のためシフトを編集できません');
    if (change.operation !== 'upsert' && change.operation !== 'delete') throw new Error('シフトの編集操作が不正です');
    const result = { employeeId: employeeId, date: change.date, operation: change.operation };
    const key = shiftRecordKey_(result);
    if (seen[key]) throw new Error('同じ従業員・日付の変更が重複しています');
    seen[key] = true;
    if (change.operation === 'delete') return result;

    result.workType = change.workType;
    if (change.workType === 'lunch') {
      result.startTime = '11:30';
      result.endTime = '15:00';
    } else if (change.workType === 'dinner') {
      result.startTime = '17:00';
      result.endTime = '22:00';
    } else if (change.workType === 'all') {
      result.startTime = '11:30';
      result.endTime = '22:00';
    } else if (change.workType === 'unavailable') {
      result.startTime = '';
      result.endTime = '';
    } else if (change.workType === 'custom') {
      if (typeof change.startTime !== 'string' || typeof change.endTime !== 'string' ||
          !SHIFT_CUSTOM_TIME_PATTERN.test(change.startTime) || !SHIFT_CUSTOM_TIME_PATTERN.test(change.endTime) ||
          change.startTime >= change.endTime) {
        throw new Error('時間を指定する場合は、開始・終了をHH:mmで指定し、終了を開始より後にしてください');
      }
      result.startTime = change.startTime;
      result.endTime = change.endTime;
    } else {
      throw new Error('勤務区分が不正です');
    }
    return result;
  });
}

/** 表示は読み取りのみ。旧列名を再利用し、区分のない旧勤務は時刻を保持したcustomとする。 */
function readShiftStore_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('対象のスプレッドシートが見つかりません');
  const sheet = spreadsheet.getSheetByName(SHIFT_SHEET_NAME);
  const store = { spreadsheet: spreadsheet, sheet: sheet, columns: {}, records: [] };
  if (!sheet || sheet.getLastRow() === 0) return store;

  const values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  SHIFT_FIELDS.forEach(function (field) {
    values[0].forEach(function (header, index) {
      if (SHIFT_HEADER_ALIASES[field].indexOf(String(header).trim()) === -1) return;
      if (store.columns[field]) throw new Error('シフトシートの「' + header + '」に対応する列が複数あります。既存データを確認してください');
      store.columns[field] = index + 1;
    });
  });
  ['employeeId', 'date', 'startTime', 'endTime'].forEach(function (field) {
    if (!store.columns[field]) throw new Error('既存のシフトシートの列構成を読み取れません。従業員ID・日付・開始時刻・終了時刻の見出しを確認してください（既存データは変更していません）');
  });

  const seen = Object.create(null);
  for (let index = 1; index < values.length; index++) {
    const raw = values[index];
    const employeeId = String(raw[store.columns.employeeId - 1]).trim();
    const rawDate = raw[store.columns.date - 1];
    if (employeeId === '' && (rawDate === '' || rawDate === null)) continue;
    const date = rawDate instanceof Date
      ? Utilities.formatDate(rawDate, TIME_ZONE, 'yyyy-MM-dd')
      : String(rawDate).trim().replace(/\//g, '-');
    if (employeeId === '' || !parseShiftDate_(date)) {
      throw new Error('シフトシートの' + (index + 1) + '行目の従業員IDまたは日付を確認してください（既存データは変更していません）');
    }
    let workType = store.columns.workType ? String(raw[store.columns.workType - 1]).trim() : '';
    const legacyTypes = { 'ランチ': 'lunch', 'ディナー': 'dinner', 'オール': 'all', '時間編集': 'custom', '出勤不可': 'unavailable' };
    workType = legacyTypes[workType] || workType || 'custom';
    if (['lunch', 'dinner', 'all', 'custom', 'unavailable'].indexOf(workType) === -1) {
      throw new Error('シフトシートの' + (index + 1) + '行目の勤務区分を確認してください（既存データは変更していません）');
    }
    const record = {
      employeeId: employeeId,
      date: date,
      workType: workType,
      startTime: readShiftTime_(raw[store.columns.startTime - 1], index + 1),
      endTime: readShiftTime_(raw[store.columns.endTime - 1], index + 1),
      rowNumber: index + 1
    };
    const key = shiftRecordKey_(record);
    if (seen[key]) throw new Error('シフトシートに同じ従業員・日付の行が複数あります（従業員ID: ' + employeeId + '、' + date + '）。既存データを確認してください');
    seen[key] = true;
    store.records.push(record);
  }
  return store;
}

function prepareShiftStoreForWrite_(store) {
  if (!store.sheet) store.sheet = store.spreadsheet.insertSheet(SHIFT_SHEET_NAME);
  const sheet = store.sheet;
  if (sheet.getLastRow() === 0) {
    if (sheet.getMaxColumns() < SHIFT_HEADERS.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), SHIFT_HEADERS.length - sheet.getMaxColumns());
    sheet.getRange(1, 1, 1, SHIFT_HEADERS.length).setValues([SHIFT_HEADERS]);
    SHIFT_FIELDS.forEach(function (field, index) { store.columns[field] = index + 1; });
  } else if (!store.columns.workType) {
    // 旧シートの出退勤時刻・ポジション・備考等をそのまま残し、不足する区分列だけ末尾に追加。
    const column = sheet.getLastColumn() + 1;
    if (column > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), column - sheet.getMaxColumns());
    sheet.getRange(1, column).setValue('勤務区分');
    store.columns.workType = column;
  }
}

/** 変更対象の行とシフトの5列だけを書き込み、他の列の値や数式を保持する。 */
function writeShiftChanges_(store, writes) {
  writes.sort(function (a, b) { return a.rowNumber - b.rowNumber; });
  const lastRow = writes[writes.length - 1].rowNumber;
  if (lastRow > store.sheet.getMaxRows()) store.sheet.insertRowsAfter(store.sheet.getMaxRows(), lastRow - store.sheet.getMaxRows());
  const fields = SHIFT_FIELDS.map(function (field, index) {
    return { column: store.columns[field], valueIndex: index };
  }).sort(function (a, b) { return a.column - b.column; });
  const batches = [];

  // 連続する行・列はまとめて保存する。間の無関係な行や列は書き戻さない。
  for (let rowStart = 0; rowStart < writes.length;) {
    let rowEnd = rowStart + 1;
    while (rowEnd < writes.length && writes[rowEnd].rowNumber === writes[rowEnd - 1].rowNumber + 1) rowEnd++;
    for (let colStart = 0; colStart < fields.length;) {
      let colEnd = colStart + 1;
      while (colEnd < fields.length && fields[colEnd].column === fields[colEnd - 1].column + 1) colEnd++;
      const values = writes.slice(rowStart, rowEnd).map(function (write) {
        return fields.slice(colStart, colEnd).map(function (field) { return write.values[field.valueIndex]; });
      });
      const range = store.sheet.getRange(writes[rowStart].rowNumber, fields[colStart].column, rowEnd - rowStart, colEnd - colStart);
      const original = range.getValues();
      const formulas = range.getFormulas();
      original.forEach(function (row, rowIndex) {
        row.forEach(function (value, columnIndex) {
          if (formulas[rowIndex][columnIndex]) original[rowIndex][columnIndex] = formulas[rowIndex][columnIndex];
        });
      });
      batches.push({ range: range, values: values, original: original, formats: range.getNumberFormats() });
      colStart = colEnd;
    }
    rowStart = rowEnd;
  }

  // 列が離れた旧シートでも、途中の書込失敗でIDだけの不完全な行を残さない。
  let attempted = 0;
  try {
    batches.forEach(function (batch) {
      attempted++;
      batch.range.setNumberFormat('@').setValues(batch.values);
    });
    SpreadsheetApp.flush();
  } catch (error) {
    let restored = true;
    for (let index = attempted - 1; index >= 0; index--) {
      try {
        batches[index].range.setNumberFormats(batches[index].formats).setValues(batches[index].original);
      } catch (restoreError) {
        restored = false;
      }
    }
    try { SpreadsheetApp.flush(); } catch (flushError) { restored = false; }
    if (!restored) throw new Error('シフト保存と元データの復元に失敗しました。シフトシートの内容を確認してから再登録してください');
    throw error;
  }
}

function shiftPublicRecord_(record) {
  const result = {};
  SHIFT_FIELDS.forEach(function (field) { result[field] = record[field]; });
  return result;
}

function shiftRecordKey_(record) {
  return JSON.stringify([record.employeeId, record.date]);
}

function parseShiftDate_(value) {
  if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value)) return null;
  const date = new Date(value + 'T00:00:00Z');
  if (!isFinite(date.getTime()) || date.getUTCFullYear() < 1900 || date.getUTCFullYear() > 2999 ||
      shiftDateText_(date) !== value) return null;
  return date;
}

function shiftDateText_(date) {
  return date.toISOString().slice(0, 10);
}

function readShiftTime_(value, rowNumber) {
  if (value === '' || value === null || value === undefined) return '';
  if (value instanceof Date) return Utilities.formatDate(value, TIME_ZONE, 'HH:mm');
  let text = String(value).trim();
  if (/^\d:\d{2}$/.test(text)) text = '0' + text;
  if (!SHIFT_TIME_PATTERN.test(text)) throw new Error('シフトシートの' + rowNumber + '行目の時刻を確認してください（既存データは変更していません）');
  return text;
}
