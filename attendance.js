// 出退勤・休憩・労働時間・給与計算を行う

function recordClock(empId, type) {
  if (CLOCK_TYPES.indexOf(type) === -1) {
    throw new Error('不正な打刻種別です')
  }
  var timeClocksSheet = getSheet(SHEET_NAMES.TIME_CLOCKS)
  var targetRow = timeClocksSheet.getLastRow() + 1
  var now = new Date()
  timeClocksSheet.getRange(targetRow, 1).setValue(empId)
  timeClocksSheet.getRange(targetRow, 2).setValue(type)
  timeClocksSheet.getRange(targetRow, 3).setValue(now)

  var formattedTime = Utilities.formatDate(now, TIME_ZONE, "HH:mm:ss")
  return type + 'を記録しました（' + formattedTime + '）'
}

function resetClockRecords() {
  var sheet = getSheet(SHEET_NAMES.TIME_CLOCKS);
  var lastRow = sheet.getLastRow();

  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
  }
}

function getTodayAttendances() {
  var now = new Date();
  var today = Utilities.formatDate(now, TIME_ZONE, 'yyyy/MM/dd');
  var weekdayLabels = ['日', '月', '火', '水', '木', '金', '土'];
  var dateLabel = Utilities.formatDate(now, TIME_ZONE, 'yyyy年M月d日') +
    '（' + weekdayLabels[now.getDay()] + '）';

  return {
    'date': dateLabel,
    'attendances': buildAttendances(today)
  };
}

function buildAttendances(targetDate) {
  var timeClocksSheet = getSheet(SHEET_NAMES.TIME_CLOCKS);
  var lastRow = timeClocksSheet.getLastRow();
  if (lastRow < 2) { // ヘッダー行しか無い＝打刻が1件も無い
    return [];
  }

  // [従業員番号, 種別, 日時] の配列をまとめて取得する
  var clockValues = timeClocksSheet.getRange(2, 1, lastRow - 1, 3).getValues();

  var eventsByEmpId = {}; // 従業員番号 -> その日の打刻イベント { type, time } の配列
  var empIds = []; // 集計結果を配列に戻すための従業員番号の一覧

  for (var i = 0; i < clockValues.length; i++) {
    var empId = clockValues[i][0];
    var type = clockValues[i][1];
    var clockedAt = clockValues[i][2];

    // 従業員番号か日時が欠けている行は、打刻として成立していないので集計しない
    if (empId === '' || !(clockedAt instanceof Date)) {
      continue;
    }
    if (Utilities.formatDate(clockedAt, TIME_ZONE, 'yyyy/MM/dd') !== targetDate) {
      continue; // 対象日以外の打刻は無視する
    }

    if (!eventsByEmpId[empId]) {
      eventsByEmpId[empId] = [];
      empIds.push(empId);
    }
    eventsByEmpId[empId].push({ 'type': type, 'time': clockedAt });
  }

  var summaryByEmpId = {}; // 従業員番号 -> summarizeDayEvents() の結果
  for (var e = 0; e < empIds.length; e++) {
    summaryByEmpId[empIds[e]] = summarizeDayEvents(eventsByEmpId[empIds[e]]);
  }

  // 出勤が早い順に並べる。出勤の打刻が無い従業員は並べようが無いので末尾へ回す
  empIds.sort(function (a, b) {
    var aClockInAt = summaryByEmpId[a].clockInAt;
    var bClockInAt = summaryByEmpId[b].clockInAt;
    if (aClockInAt === null && bClockInAt === null) {
      return 0;
    }
    if (aClockInAt === null) {
      return 1;
    }
    if (bClockInAt === null) {
      return -1;
    }
    return aClockInAt.getTime() - bClockInAt.getTime();
  });

  var nameByEmpId = getEmployeeNameMap();
  var rateByEmpId = getEmployeeRateMap();
  var isHolidayRate = isHolidayDate(targetDate); // 土日、または「祝日」シートに登録された日なら休日加給を適用する

  var attendances = [];
  for (var j = 0; j < empIds.length; j++) {
    var id = empIds[j];
    var summary = summaryByEmpId[id];
    var actualMinutes = calcActualWorkingMinutes(summary.clockInAt, summary.clockOutAt, summary.breakMinutes);
    var rate = rateByEmpId[String(id)] || { 'base': 0, 'holidayBonus': 0 };

    attendances.push({
      // 名簿から消された従業員の打刻が残っていても行が消えないよう、番号を出して気付けるようにする
      'name': nameByEmpId[id] || ('従業員番号 ' + id),
      'clockIn': formatClockTime(summary.clockInAt),
      'clockOut': formatClockTime(summary.clockOutAt),
      'breakTime': formatBreakTime(summary.breakMinutes, summary.breakOngoing),
      'workingTime': formatMinutesLabel(actualMinutes),
      'estimatedWage': formatEstimatedWage(actualMinutes, rate, isHolidayRate)
    });
  }
  return attendances;
}

function summarizeDayEvents(events) {
  events.sort(function (a, b) { return a.time.getTime() - b.time.getTime(); });

  var clockInAt = null;
  var clockOutAt = null;
  var breakMinutes = 0;
  var breakStartAt = null;

  for (var i = 0; i < events.length; i++) {
    var ev = events[i];
    if (ev.type === '出勤') {
      if (clockInAt === null) {
        clockInAt = ev.time;
      }
    } else if (ev.type === '退勤') {
      clockOutAt = ev.time; // 時刻順に処理しているので、上書きしていけば最も遅い退勤が残る
    } else if (ev.type === '休憩開始') {
      if (breakStartAt === null) { // 休憩中に重ねて押された場合は無視する
        breakStartAt = ev.time;
      }
    } else if (ev.type === '休憩終了') {
      if (breakStartAt !== null) {
        breakMinutes += Math.floor((ev.time.getTime() - breakStartAt.getTime()) / (60 * 1000));
        breakStartAt = null;
      }
    }
  }

  return {
    'clockInAt': clockInAt,
    'clockOutAt': clockOutAt,
    'breakMinutes': breakMinutes,
    'breakOngoing': breakStartAt !== null
  };
}

function calcActualWorkingMinutes(clockInAt, clockOutAt, breakMinutes) {
  if (clockInAt === null || clockOutAt === null) {
    return null;
  }
  var grossMinutes = Math.floor((clockOutAt.getTime() - clockInAt.getTime()) / (60 * 1000));
  var actualMinutes = grossMinutes - breakMinutes;
  if (actualMinutes < 0) {
    return null;
  }
  return actualMinutes;
}

function getEmployeeRateMap() {
  var empSheet = getSheet(SHEET_NAMES.EMPLOYEES);
  var lastRow = empSheet.getLastRow();
  var rateByEmpId = {};
  if (lastRow < 2) {
    return rateByEmpId;
  }
  var empValues = empSheet.getRange(2, 1, lastRow - 1, 5).getValues(); // A:従業員番号 ... D:基本時給 E:休日加給
  for (var i = 0; i < empValues.length; i++) {
    var empId = empValues[i][0];
    if (empId === '') {
      continue;
    }
    rateByEmpId[String(empId)] = {
      'base': Number(empValues[i][3]) || 0,
      'holidayBonus': Number(empValues[i][4]) || 0
    };
  }
  return rateByEmpId;
}

function isHolidayDate(targetDate) {
  var date = new Date(targetDate);
  var isWeekend = (date.getDay() === 0 || date.getDay() === 6);
  if (isWeekend) {
    return true;
  }
  return !!getHolidayDateSet()[targetDate];
}

function getHolidayDateSet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.HOLIDAYS);
  if (!sheet) {
    return {};
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    return {};
  }
  var values = sheet.getRange(1, 1, lastRow, 1).getValues();
  var holidaySet = {};
  for (var i = 0; i < values.length; i++) {
    var value = values[i][0];
    if (value instanceof Date) {
      holidaySet[Utilities.formatDate(value, TIME_ZONE, 'yyyy/MM/dd')] = true;
    }
  }
  return holidaySet;
}

function formatClockTime(clockedAt) {
  if (clockedAt === null) {
    return UNRECORDED_LABEL;
  }
  return Utilities.formatDate(clockedAt, TIME_ZONE, 'HH:mm');
}

function formatMinutesLabel(totalMinutes) {
  if (totalMinutes === null) {
    return NO_VALUE_LABEL;
  }
  var hours = Math.floor(totalMinutes / 60);
  var minutes = totalMinutes % 60;
  if (hours === 0) {
    return minutes + '分';
  }
  if (minutes === 0) {
    return hours + '時間';
  }
  return hours + '時間' + minutes + '分';
}

function formatBreakTime(breakMinutes, breakOngoing) {
  if (breakOngoing) {
    return '休憩中';
  }
  if (breakMinutes === 0) {
    return NO_VALUE_LABEL;
  }
  return formatMinutesLabel(breakMinutes);
}

function formatEstimatedWage(actualMinutes, rate, isHolidayRate) {
  if (actualMinutes === null) {
    return NO_VALUE_LABEL;
  }
  var hourlyRate = rate.base + (isHolidayRate ? rate.holidayBonus : 0);
  var wage = Math.floor(actualMinutes / 60 * hourlyRate);
  return wage.toLocaleString('ja-JP') + '円';
}