// 出退勤・休憩・労働時間・給与計算を行う

//打刻履歴をおこなう
function recordClock(empId, type) {

  //打刻履歴に引数の種別があれば0,なければ-1
  if (CLOCK_TYPES.indexOf(type) === -1) {
    throw new Error('不正な打刻種別です');
  }

  //"打刻履歴"シートを取得
  const timeClocksSheet = getSheet(SHEET_NAMES.TIME_CLOCKS);

  //書き込み先の空白の行を指定
  const targetRow = timeClocksSheet.getLastRow() + 1;

  //時刻インスタンスを作成
  const now = new Date();

  //先ほどの行の 1列目:Id、2列目:種別、3列目:現在時刻 を書き込む
  timeClocksSheet.getRange(targetRow, 1).setValue(empId);
  timeClocksSheet.getRange(targetRow, 2).setValue(type);
  timeClocksSheet.getRange(targetRow, 3).setValue(now);

  //now の日時を、Asia/Tokyo の時刻として、時:分:秒 形式の文字列に変換
  const formattedTime = Utilities.formatDate(now, TIME_ZONE, "HH:mm:ss");
  
  return type + `を記録しました（${formattedTime}）`;
}

//
function resetClockRecords() {
  const sheet = getSheet(SHEET_NAMES.TIME_CLOCKS);
  const lastRow = sheet.getLastRow();

  //対象の範囲をすべて削除
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
  }
}



function getTodayAttendances() {
  const now = new Date();
  
  //プログラム内部用の日付
  const today = Utilities.formatDate(now, TIME_ZONE, 'yyyy/MM/dd');
  
  const weekdayLabels = ['日', '月', '火', '水', '木', '金', '土'];

  //now.getDay()で数字を取得。日曜日なら0。
  const dateLabel = Utilities.formatDate(now, TIME_ZONE, 'yyyy年M月d日') +
    `（${weekdayLabels[now.getDay()]}）`;

  return {
    'date': dateLabel,
    'attendances': buildAttendances(today)
  };
}

function buildAttendances(targetDate) {
  
  const timeClocksSheet = getSheet(SHEET_NAMES.TIME_CLOCKS);
  const lastRow = timeClocksSheet.getLastRow();

  //ヘッダーしかないので空の配列を返す
  if (lastRow < 2) {
    return [];
  }

  //[Id,種別,日付]の二次元配列を返す
  const clockValues = timeClocksSheet.getRange(2, 1, lastRow - 1, 3).getValues();

  const eventsByEmpId = {}; 
  const empIds = []; 

  

  for (let i = 0; i < clockValues.length; i++) {
    const empId = clockValues[i][0];
    const type = clockValues[i][1];
    const clockedAt = clockValues[i][2];

    //instanceofはこの変数の値はDateオブジェクトか？
    if (empId === '' || !(clockedAt instanceof Date)) {
      continue;
    }

    //まあ丁寧すぎるチェック
    if (Utilities.formatDate(clockedAt, TIME_ZONE, 'yyyy/MM/dd') !== targetDate) {
      continue; 
    }

    if (!eventsByEmpId[String(empId)]) {

      //{1: [],...}
      eventsByEmpId[String(empId)] = [];

      //重複なしで、その日に打刻が存在した従業員番号の配列を作成
      empIds.push(empId);
    }
    eventsByEmpId[String(empId)].push({ 'type': type, 'time': clockedAt });
  }

  const summaryByEmpId = {}; 

  for (let e = 0; e < empIds.length; e++) {
    summaryByEmpId[String(empIds[e])] = summarizeDayEvents(eventsByEmpId[String(empIds[e])]);
  }

  
  empIds.sort(function (a, b) {
    var aClockInAt = summaryByEmpId[String(a)].clockInAt;
    var bClockInAt = summaryByEmpId[String(b)].clockInAt;
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
  var isHolidayRate = isHolidayDate(targetDate);

  var attendances = [];
  for (var j = 0; j < empIds.length; j++) {
    var id = empIds[j];
    var summary = summaryByEmpId[String(id)];
    var actualMinutes = calcActualWorkingMinutes(summary.clockInAt, summary.clockOutAt, summary.breakMinutes);
    var rate = rateByEmpId[String(id)] || { 'base': 0, 'holidayBonus': 0 };

    attendances.push({
      'name': nameByEmpId[String(id)] || ('従業員番号 ' + id),
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