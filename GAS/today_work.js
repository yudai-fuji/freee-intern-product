// その日の担当ポジション・確認事項を管理する

function ensureTodayWorkSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.TODAY_WORK);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.TODAY_WORK);
  }

  const headerRange = sheet.getRange(1, 1, 1, TODAY_WORK_SHEET_HEADER.length);
  const currentHeader = headerRange.getValues()[0];
  for (let i = 0; i < TODAY_WORK_SHEET_HEADER.length; i++) {
    if (currentHeader[i] !== TODAY_WORK_SHEET_HEADER[i]) {
      headerRange.setValues([TODAY_WORK_SHEET_HEADER]);
      break;
    }
  }
  return sheet;
}

function getTodayDateKey() {
  return Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy/MM/dd');
}

function findTodayWorkRow(sheet, dateKey, empId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return -1;
  }
  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // A:日付 B:従業員ID
  for (let i = 0; i < values.length; i++) {
    const rowDate = values[i][0];
    const rowEmpId = values[i][1];
    if (rowEmpId === '') {
      continue;
    }
    const rowDateKey = (rowDate instanceof Date) ? Utilities.formatDate(rowDate, TIME_ZONE, 'yyyy/MM/dd') : String(rowDate);
    if (rowDateKey === dateKey && String(rowEmpId) === String(empId)) {
      return i + 2; // ヘッダー行の分だけシート上の行番号は配列の添字より2大きい
    }
  }
  return -1;
}

function getTodayWorkMapForDate(dateKey) {
  const sheet = ensureTodayWorkSheet();
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow < 2) {
    return map;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, TODAY_WORK_SHEET_HEADER.length).getValues();
  for (let i = 0; i < values.length; i++) {
    const rowDate = values[i][0];
    const rowEmpId = values[i][1];
    if (rowEmpId === '') {
      continue;
    }
    const rowDateKey = (rowDate instanceof Date) ? Utilities.formatDate(rowDate, TIME_ZONE, 'yyyy/MM/dd') : String(rowDate);
    if (rowDateKey !== dateKey) {
      continue; // 対象日以外の行は無視する
    }
    map[String(rowEmpId)] = {
      'position': values[i][2] || '',
      'noticesText': values[i][3] || ''
    };
  }
  return map;
}

function splitNoticesText(noticesText) {
  if (!noticesText) {
    return [];
  }
  const lines = String(noticesText).split(/\r\n|\r|\n/);
  const notices = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (line !== '') {
      notices.push(line);
    }
  }
  return notices;
}

function getTodayWorkSettings(authToken) {
  requireManagerRole_(authToken);

  const now = new Date();
  const weekdayLabels = ['日', '月', '火', '水', '木', '金', '土'];
  const dateLabel = Utilities.formatDate(now, TIME_ZONE, 'yyyy年M月d日') +
    '（' + weekdayLabels[now.getDay()] + '）';

  const employees = getEmployees();
  const todayMap = getTodayWorkMapForDate(getTodayDateKey());

  const settings = [];
  for (let i = 0; i < employees.length; i++) {
    const empId = employees[i]['id'];
    const saved = todayMap[String(empId)];
    settings.push({
      'id': empId,
      'name': employees[i]['name'],
      'position': saved ? saved.position : '',
      'noticesText': saved ? saved.noticesText : ''
    });
  }

  return {
    'date': dateLabel,
    'settings': settings
  };
}

function saveTodayWork(empId, position, noticesText, authToken) {
  requireManagerRole_(authToken);

  // 有効な従業員だけを保存先にする（削除済みの人は本日の業務の設定対象から外す）
  const nameByEmpId = getActiveEmployeeNameMap();
  const empName = nameByEmpId[String(empId)];
  console.log(empName)
  if (!empName) {
    throw new Error('この従業員は選べません。画面を開き直してください（従業員番号 ' + empId + '）');
  }

  position = position ? String(position) : '';
  if (position !== '' && POSITION_ITEMS.indexOf(position) === -1) {
    throw new Error('不正なポジションです: ' + position);
  }
  noticesText = noticesText ? String(noticesText) : '';

  const sheet = ensureTodayWorkSheet();
  const dateKey = getTodayDateKey();
  let targetRow = findTodayWorkRow(sheet, dateKey, empId);
  if (targetRow === -1) {
    targetRow = sheet.getLastRow() + 1;
    sheet.getRange(targetRow, 1).setValue(new Date(dateKey));
    sheet.getRange(targetRow, 2).setValue(empId);
  }
  sheet.getRange(targetRow, 3).setValue(position);
  sheet.getRange(targetRow, 4).setValue(noticesText);

  return empName + 'さんの本日の業務を保存しました';
}

function getTodayWorkInfo(empId) {
  const todayMap = getTodayWorkMapForDate(getTodayDateKey());
  const saved = todayMap[String(empId)];

  const position = saved ? saved.position : '';
  const notices = splitNoticesText(saved ? saved.noticesText : '');
  // 引継ぎ連携（未実装）は、実装時にここで notices へ追加する

  return {
    'position': position,
    'notices': notices
  };
}