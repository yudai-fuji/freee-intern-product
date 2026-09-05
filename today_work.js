// その日の担当ポジション・確認事項を管理する

function ensureTodayWorkSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.TODAY_WORK);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.TODAY_WORK);
  }

  var headerRange = sheet.getRange(1, 1, 1, TODAY_WORK_SHEET_HEADER.length);
  var currentHeader = headerRange.getValues()[0];
  for (var i = 0; i < TODAY_WORK_SHEET_HEADER.length; i++) {
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
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return -1;
  }
  var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // A:日付 B:従業員ID
  for (var i = 0; i < values.length; i++) {
    var rowDate = values[i][0];
    var rowEmpId = values[i][1];
    if (rowEmpId === '') {
      continue;
    }
    var rowDateKey = (rowDate instanceof Date) ? Utilities.formatDate(rowDate, TIME_ZONE, 'yyyy/MM/dd') : String(rowDate);
    if (rowDateKey === dateKey && String(rowEmpId) === String(empId)) {
      return i + 2; // ヘッダー行の分だけシート上の行番号は配列の添字より2大きい
    }
  }
  return -1;
}

function getTodayWorkMapForDate(dateKey) {
  var sheet = ensureTodayWorkSheet();
  var lastRow = sheet.getLastRow();
  var map = {};
  if (lastRow < 2) {
    return map;
  }

  var values = sheet.getRange(2, 1, lastRow - 1, TODAY_WORK_SHEET_HEADER.length).getValues();
  for (var i = 0; i < values.length; i++) {
    var rowDate = values[i][0];
    var rowEmpId = values[i][1];
    if (rowEmpId === '') {
      continue;
    }
    var rowDateKey = (rowDate instanceof Date) ? Utilities.formatDate(rowDate, TIME_ZONE, 'yyyy/MM/dd') : String(rowDate);
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
  var lines = String(noticesText).split(/\r\n|\r|\n/);
  var notices = [];
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (line !== '') {
      notices.push(line);
    }
  }
  return notices;
}

function getTodayWorkSettings() {
  var now = new Date();
  var weekdayLabels = ['日', '月', '火', '水', '木', '金', '土'];
  var dateLabel = Utilities.formatDate(now, TIME_ZONE, 'yyyy年M月d日') +
    '（' + weekdayLabels[now.getDay()] + '）';

  var employees = getEmployees();
  var todayMap = getTodayWorkMapForDate(getTodayDateKey());

  var settings = [];
  for (var i = 0; i < employees.length; i++) {
    var empId = employees[i]['id'];
    var saved = todayMap[String(empId)];
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

function saveTodayWork(empId, position, noticesText) {
  var nameByEmpId = getEmployeeNameMap();
  var empName = nameByEmpId[String(empId)];
  if (!empName) {
    throw new Error('従業員名簿に存在しない従業員番号です: ' + empId);
  }

  position = position ? String(position) : '';
  if (position !== '' && POSITION_ITEMS.indexOf(position) === -1) {
    throw new Error('不正なポジションです: ' + position);
  }
  noticesText = noticesText ? String(noticesText) : '';

  var sheet = ensureTodayWorkSheet();
  var dateKey = getTodayDateKey();
  var targetRow = findTodayWorkRow(sheet, dateKey, empId);
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
  var todayMap = getTodayWorkMapForDate(getTodayDateKey());
  var saved = todayMap[String(empId)];

  var position = saved ? saved.position : '';
  var notices = splitNoticesText(saved ? saved.noticesText : '');
  // 食材管理・引継ぎ連携（未実装）は、実装時にここで notices へ追加する

  return {
    'position': position,
    'notices': notices
  };
}