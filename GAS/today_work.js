/**
 * 本日の業務シートを用意し、日付・従業員ID・ポジション・確認事項の見出しをそろえる。
 * 戻り値 - GoogleAppsScript.Spreadsheet.Sheet
 */
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

/**
 * 今日を日本時間のyyyy/MM/ddとして取得する。
 * 戻り値 - string
 */
function getTodayDateKey() {
  return Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy/MM/dd');
}

/**
 * 日付と従業員IDの両方が一致する業務行を探す。
 * sheet - GoogleAppsScript.Spreadsheet.Sheet
 * dateKey - string
 * empId - unknown
 * 戻り値 - number
 */
function findTodayWorkRow(sheet, dateKey, empId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return -1;
  }
  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let i = 0; i < values.length; i++) {
    const rowDate = values[i][0];
    const rowEmpId = values[i][1];
    if (rowEmpId === '') {
      continue;
    }
    const rowDateKey = (rowDate instanceof Date) ? Utilities.formatDate(rowDate, TIME_ZONE, 'yyyy/MM/dd') : String(rowDate);
    if (rowDateKey === dateKey && String(rowEmpId) === String(empId)) {
      return i + 2;
    }
  }
  return -1;
}

/**
 * 指定日の業務だけを取り出し、従業員IDで引ける対応表にする。
 * dateKey - string
 * 戻り値 - {[key: string]: {position: unknown, noticesText: unknown}}
 */
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
      continue;
    }
    map[String(rowEmpId)] = {
      'position': values[i][2] || '',
      'noticesText': values[i][3] || ''
    };
  }
  return map;
}

/**
 * 確認事項を改行で分割し、前後の空白と空行を除いて配列にする。
 * noticesText - unknown
 * 戻り値 - string[]
 */
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

/**
 * 店長向けに、有効な全従業員と今日の担当・確認事項をまとめて返す。
 * authToken - string
 * 戻り値 - {date: string, settings: ({id: unknown, name: string, position: unknown, noticesText: unknown})[]}
 */
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

/**
 * 従業員とポジションを検証し、今日の担当・確認事項を保存する。
 * empId - unknown
 * position - string
 * noticesText - string
 * authToken - string
 * 戻り値 - string
 */
function saveTodayWork(empId, position, noticesText, authToken) {
  requireManagerRole_(authToken);

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

/**
 * 従業員1人の今日のポジションと確認事項配列を返す。
 * empId - unknown
 * 戻り値 - {position: unknown, notices: string[]}
 */
function getTodayWorkInfo(empId) {
  const todayMap = getTodayWorkMapForDate(getTodayDateKey());
  const saved = todayMap[String(empId)];

  const position = saved ? saved.position : '';
  const notices = splitNoticesText(saved ? saved.noticesText : '');

  return {
    'position': position,
    'notices': notices
  };
}