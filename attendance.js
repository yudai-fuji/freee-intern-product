// 出退勤・休憩・労働時間・給与計算を行う

// 画面を開き直しても、日本時間の当日履歴から打刻状態を復元する。
function getTodayClockState(empId, authToken) {
  requireLoggedIn_(authToken);
  requireActiveClockEmployee_(empId);
  return readTodayClockState_(empId, getSheet(SHEET_NAMES.TIME_CLOCKS), new Date());
}

function requireActiveClockEmployee_(empId) {
  const activeNameByEmpId = getActiveEmployeeNameMap();
  if (!activeNameByEmpId[String(empId)]) {
    throw new Error('この従業員は選べません。画面を再読み込みしてください');
  }
  return activeNameByEmpId;
}

function allowedClockTypes_(state) {
  if (state === 'not_clocked_in') return ['出勤'];
  if (state === 'working') return ['退勤', '休憩開始'];
  if (state === 'on_break') return ['休憩終了'];
  return [];
}

// 不正な順序は null。同日の再出勤は認めず、休憩終了後は勤務中に戻す。
function nextClockState_(state, type) {
  if (allowedClockTypes_(state).indexOf(type) === -1) return null;
  if (type === '退勤') return 'clocked_out';
  if (type === '休憩開始') return 'on_break';
  return 'working';
}

function readTodayClockState_(empId, sheet, now) {
  const date = Utilities.formatDate(now, TIME_ZONE, 'yyyy/MM/dd');
  const result = { employeeId: String(empId), date: date, state: 'not_clocked_in', allowedTypes: [] };
  const lastRow = sheet.getLastRow();
  const rows = lastRow < 2 ? [] : sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const events = [];
  for (let i = 0; i < rows.length; i++) {
    const time = rows[i][2];
    if (String(rows[i][0]) !== String(empId) || !(time instanceof Date) || isNaN(time.getTime())) continue;
    if (Utilities.formatDate(time, TIME_ZONE, 'yyyy/MM/dd') !== date) continue;
    events.push({ type: rows[i][1], time: time.getTime(), row: i + 2 });
  }
  events.sort(function (a, b) { return a.time - b.time || a.row - b.row; });
  for (let i = 0; i < events.length; i++) {
    const next = nextClockState_(result.state, events[i].type);
    if (next === null) {
      result.state = 'invalid';
      result.error = '本日の打刻履歴の順序が不正です（' + events[i].row + '行目）。店長に確認してください';
      return result;
    }
    result.state = next;
  }
  result.allowedTypes = allowedClockTypes_(result.state);
  return result;
}

function recordClock(empId, type, ticket, authToken) {
  return recordClockWithValidation_(empId, type, ticket, authToken);
}

// 通常打刻と共有メモ付き退勤で、状態の確認から保存まで同じロックを使う。
function recordClockWithValidation_(empId, type, ticket, authToken, beforeRecord) {
  requireLoggedIn_(authToken);
  if (CLOCK_TYPES.indexOf(type) === -1) throw new Error('不正な打刻種別です');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('他の打刻処理が進行中です。少し待ってからもう一度お試しください');
  }
  try {
    const activeNameByEmpId = requireActiveClockEmployee_(empId);
    const sheet = getSheet(SHEET_NAMES.TIME_CLOCKS);
    const now = new Date();
    const clockState = readTodayClockState_(empId, sheet, now);
    if (clockState.state === 'invalid') throw new Error(clockState.error);
    const next = nextClockState_(clockState.state, type);
    if (next === null) throw new Error('現在の打刻状態では「' + type + '」を記録できません。本日の打刻状態を確認してください');
    consumeClockPinTicket_(empId, type, ticket, authToken, now);

    // 状態検証後にメモを保存する。保存に失敗した場合は打刻を行わない。
    if (beforeRecord) {
      beforeRecord(activeNameByEmpId);
      SpreadsheetApp.flush();
    }
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, 3).setValues([[empId, type, now]]);
    SpreadsheetApp.flush();
    clockState.state = next;
    clockState.allowedTypes = allowedClockTypes_(next);
    return {
      message: type + 'を記録しました（' + Utilities.formatDate(now, TIME_ZONE, 'HH:mm:ss') + '）',
      clockState: clockState
    };
  } finally {
    lock.releaseLock();
  }
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



function getTodayAttendances(authToken) {
  requireManagerRole_(authToken);

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

  //functionは比較関数。empIdsの中身から2つとって大きさを比較。returnが正ならb a、負ならa bの順にする。
  empIds.sort(function (a, b) {
    const aClockInAt = summaryByEmpId[String(a)].clockInAt;
    const bClockInAt = summaryByEmpId[String(b)].clockInAt;
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

  const nameByEmpId = getEmployeeNameMap();
  const rateByEmpId = getEmployeeRateMap();
  const isHolidayRate = isHolidayDate(targetDate);

  const attendances = [];
  for (let j = 0; j < empIds.length; j++) {
    const id = empIds[j];
    const summary = summaryByEmpId[String(id)];
    const actualMinutes = calcActualWorkingMinutes(summary.clockInAt, summary.clockOutAt, summary.breakMinutes);
    const rate = rateByEmpId[String(id)] || { 'base': 0, 'holidayBonus': 0, 'skillAllowance': 0 };

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

  let clockInAt = null;
  let clockOutAt = null;
  let breakMinutes = 0;
  let breakStartAt = null;

  for (let i = 0; i < events.length; i++) {
    const ev = events[i];
    if (ev.type === '出勤') {
      if (clockInAt === null) {
        clockInAt = ev.time;
      }
    } else if (ev.type === '退勤') {
      clockOutAt = ev.time; 
    } else if (ev.type === '休憩開始') {
      if (breakStartAt === null) { 
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
  const grossMinutes = Math.floor((clockOutAt.getTime() - clockInAt.getTime()) / (60 * 1000));
  const actualMinutes = grossMinutes - breakMinutes;
  if (actualMinutes < 0) {
    return null;
  }
  return actualMinutes;
}

function getEmployeeRateMap() {
  const empSheet = getSheet(SHEET_NAMES.EMPLOYEES);
  const lastRow = empSheet.getLastRow();
  const rateByEmpId = {};
  if (lastRow < 2) {
    return rateByEmpId;
  }
  // スキル手当は名簿の右側に足した列にあるので、基本時給・休日加給とは別に読み取る
  const skillAllowanceByEmpId = getEmployeeSkillAllowanceMap();
  const empValues = empSheet.getRange(2, 1, lastRow - 1, 5).getValues(); // A:従業員番号 ... D:基本時給 E:休日加給
  for (let i = 0; i < empValues.length; i++) {
    const empId = empValues[i][0];
    if (empId === '') {
      continue;
    }
    rateByEmpId[String(empId)] = {
      'base': Number(empValues[i][3]) || 0,
      'holidayBonus': Number(empValues[i][4]) || 0,
      'skillAllowance': skillAllowanceByEmpId[String(empId)] || SKILL_ALLOWANCE_DEFAULT
    };
  }
  return rateByEmpId;
}

function isHolidayDate(targetDate) {
  const date = new Date(targetDate);
  const isWeekend = (date.getDay() === 0 || date.getDay() === 6);
  if (isWeekend) {
    return true;
  }
  return !!getHolidayDateSet()[targetDate];
}

function getHolidayDateSet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.HOLIDAYS);
  if (!sheet) {
    return {};
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    return {};
  }
  const values = sheet.getRange(1, 1, lastRow, 1).getValues();
  const holidaySet = {};
  for (let i = 0; i < values.length; i++) {
    const value = values[i][0];
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
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
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
  // 曜日ごとの時給の決め方はこれまでどおりで、スキル手当だけを1時間あたりに上乗せする
  // 平日は 基本時給＋スキル手当、土日祝は 基本時給＋休日加給＋スキル手当 になる
  const hourlyRate = rate.base + (isHolidayRate ? rate.holidayBonus : 0) + (rate.skillAllowance || 0);
  const wage = Math.floor(actualMinutes / 60 * hourlyRate);
  return wage.toLocaleString('ja-JP') + '円';
}
