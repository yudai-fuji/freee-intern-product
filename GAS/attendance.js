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
  const events = readAttendanceEvents_(sheet, date).filter(function (event) { return String(event.employeeId) === String(empId); });
  return Object.assign({ employeeId: String(empId), date: date }, getClockStateFromEvents_(events));
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
    const adjustmentColumn = getAttendanceAdjustmentColumn_(sheet, false);
    if (adjustmentColumn) { sheet.getRange(2, adjustmentColumn, lastRow - 1, 1).clearContent(); }
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
    'dateKey': today,
    'attendances': buildAttendances(today)
  };
}

function buildAttendances(targetDate) {
  const eventsByEmpId = new Map();
  const shiftDate = targetDate.replace(/\//g, '-');
  // 通常シフトの従業員だけを起点にし、未打刻でも勤怠行を用意する。
  readShiftStore_().records.forEach(function (record) {
    if (record.date === shiftDate && record.workType !== 'unavailable') {
      eventsByEmpId.set(String(record.employeeId), []);
    }
  });
  if (eventsByEmpId.size === 0) { return []; }
  readAttendanceEvents_(getSheet(SHEET_NAMES.TIME_CLOCKS), targetDate).forEach(function (event) {
    const id = String(event.employeeId);
    if (eventsByEmpId.has(id)) { eventsByEmpId.get(id).push(event); }
  });
  const names = getEmployeeNameMap();
  const rates = getEmployeeRateMap();
  const holiday = isHolidayDate(targetDate);
  return Array.from(eventsByEmpId, function (entry) {
    const id = entry[0];
    const rate = rates[id] || { base: BASE_HOURLY_WAGE, holidayBonus: HOLIDAY_HOURLY_BONUS, skillAllowance: SKILL_ALLOWANCE_DEFAULT };
    return buildAttendanceView_(id, targetDate, entry[1], names[id], rate, holiday);
  }).sort(function (a, b) {
    return (a.clockInInput || '99:99').localeCompare(b.clockInInput || '99:99') ||
      a.id.localeCompare(b.id, 'ja', { numeric: true });
  });
}

function summarizeDayEvents(events) {
  events.sort(compareClockEvents_);

  let clockInAt = null;
  let clockOutAt = null;
  let breakMinutes = 0;
  let breakStartAt = null;

  for (let i = 0; i < events.length; i++) {
    const ev = events[i];
    if (ev.type === '出勤') {
      breakMinutes += ev.breakAdjustment || 0;
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
  const rateByEmpId = {};
  const skillAllowanceByEmpId = getEmployeeSkillAllowanceMap();
  const employees = readEmployeeRows_();
  for (let i = 0; i < employees.length; i++) {
    const empId = employees[i].id;
    rateByEmpId[String(empId)] = {
      'base': BASE_HOURLY_WAGE,
      'holidayBonus': HOLIDAY_HOURLY_BONUS,
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


/** 時刻が同じ場合も出勤→休憩→退勤の順で扱う（分単位の修正に対応）。 */
function compareClockEvents_(a, b) {
  const priority = function (type) { return type === '出勤' ? -1 : type === '退勤' ? 1 : 0; };
  return a.time.getTime() - b.time.getTime() || priority(a.type) - priority(b.type) ||
    (a.rowNumber || 0) - (b.rowNumber || 0);
}

function getAttendanceAdjustmentColumn_(sheet, create) {
  const width = Math.max(3, sheet.getLastColumn());
  const header = sheet.getRange(1, 1, 1, width).getValues()[0];
  const index = header.indexOf(ATTENDANCE_BREAK_ADJUSTMENT_HEADER);
  if (index !== -1) { return index + 1; }
  if (!create) { return 0; }
  const column = width + 1;
  if (column > sheet.getMaxColumns()) { sheet.insertColumnsAfter(sheet.getMaxColumns(), column - sheet.getMaxColumns()); }
  sheet.getRange(1, column).setValue(ATTENDANCE_BREAK_ADJUSTMENT_HEADER);
  return column;
}

/** 既存の3列を維持し、追加列がない旧データの休憩補正は0分として読む。 */
function readAttendanceEvents_(sheet, dateKey) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { return []; }
  const adjustmentColumn = getAttendanceAdjustmentColumn_(sheet, false);
  return sheet.getRange(2, 1, lastRow - 1, Math.max(3, adjustmentColumn)).getValues()
    .map(function (row, index) {
      return { employeeId: row[0], type: row[1], time: row[2], rowNumber: index + 2,
        breakAdjustment: adjustmentColumn ? Number(row[adjustmentColumn - 1]) || 0 : 0 };
    }).filter(function (event) {
      return event.employeeId !== '' && event.time instanceof Date && !isNaN(event.time.getTime()) &&
        Utilities.formatDate(event.time, TIME_ZONE, 'yyyy/MM/dd') === dateKey;
    });
}

function getAttendanceRevision_(events) {
  return JSON.stringify(events.map(function (event) {
    return [event.rowNumber, event.type, event.time.getTime(), event.breakAdjustment];
  }));
}

function getClockStateFromEvents_(events) {
  let state = 'not_clocked_in';
  const ordered = events.slice().sort(compareClockEvents_);
  for (let i = 0; i < ordered.length; i++) {
    const next = nextClockState_(state, ordered[i].type);
    if (next === null) {
      return { state: 'invalid', allowedTypes: [],
        error: '打刻履歴の順序が不正です。出勤・退勤と休憩の時刻を確認してください（' + ordered[i].rowNumber + '行目）' };
    }
    state = next;
  }
  return { state: state, allowedTypes: allowedClockTypes_(state) };
}

function buildAttendanceView_(id, dateKey, events, name, rate, holiday) {
  const summary = summarizeDayEvents(events.slice());
  const actualMinutes = calcActualWorkingMinutes(summary.clockInAt, summary.clockOutAt, summary.breakMinutes);
  return {
    id: String(id), dateKey: dateKey, name: name || ('従業員番号 ' + id),
    clockIn: formatClockTime(summary.clockInAt), clockOut: formatClockTime(summary.clockOutAt),
    clockInInput: summary.clockInAt ? formatClockTime(summary.clockInAt) : '',
    clockOutInput: summary.clockOutAt ? formatClockTime(summary.clockOutAt) : '',
    breakMinutes: summary.breakMinutes, breakOngoing: summary.breakOngoing,
    breakTime: formatBreakTime(summary.breakMinutes, summary.breakOngoing),
    workingTime: formatMinutesLabel(actualMinutes),
    estimatedWage: formatEstimatedWage(actualMinutes, rate, holiday),
    revision: getAttendanceRevision_(events)
  };
}

function parseAttendanceTime_(value, dateKey, original, label) {
  if (value === '') { return null; }
  if (typeof value !== 'string' || !/^([01]\d|2[0-3]):[0-5]\d$/.test(value)) {
    throw new Error(label + 'は時:分の形式で入力してください');
  }
  // 未変更の時刻は秒も保持し、保存だけで労働時間が変わらないようにする。
  if (original && formatClockTime(original) === value) { return original; }
  const date = Utilities.parseDate(dateKey + ' ' + value, TIME_ZONE, 'yyyy/MM/dd HH:mm');
  if (Utilities.formatDate(date, TIME_ZONE, 'yyyy/MM/dd HH:mm') !== dateKey + ' ' + value) {
    throw new Error(label + 'が不正です');
  }
  return date;
}

/** 店長の行編集。通常打刻と同じロックで競合を確認し、対象従業員・対象日だけを書き換える。 */
function saveAttendance(input, authToken) {
  requireManagerRole_(authToken);
  const payload = input || {};
  return withAppLock_(function () {
    const dateKey = payload.dateKey;
    if (typeof dateKey !== 'string' || !/^\d{4}\/\d{2}\/\d{2}$/.test(dateKey)) { throw new Error('対象日が不正です'); }
    const sheet = getSheet(SHEET_NAMES.TIME_CLOCKS);
    const events = readAttendanceEvents_(sheet, dateKey).filter(function (event) {
      return String(event.employeeId) === String(payload.id);
    });
    const names = getEmployeeNameMap();
    const activeNames = getActiveEmployeeNameMap();
    if (!Object.prototype.hasOwnProperty.call(names, String(payload.id)) ||
        (!activeNames[String(payload.id)] && events.length === 0)) { throw new Error('この従業員の勤怠は編集できません'); }
    if (typeof payload.revision !== 'string' || payload.revision !== getAttendanceRevision_(events)) {
      throw new Error('編集中に打刻が更新されました。キャンセルして勤怠管理を開き直してください');
    }
    const original = summarizeDayEvents(events.slice());
    const clockIn = parseAttendanceTime_(payload.clockIn, dateKey, original.clockInAt, '出勤時刻');
    const clockOut = parseAttendanceTime_(payload.clockOut, dateKey, original.clockOutAt, '退勤時刻');
    const breakText = String(payload.breakMinutes === undefined || payload.breakMinutes === null ? '' : payload.breakMinutes).trim();
    if (!/^\d+$/.test(breakText) || !Number.isSafeInteger(Number(breakText)) || Number(breakText) > 1439) {
      throw new Error('休憩時間は0〜1439の整数（分）で入力してください');
    }
    const breakMinutes = Number(breakText);
    if (!clockIn) { throw new Error('出勤時刻を入力してください'); }
    if (clockOut && clockOut.getTime() < clockIn.getTime()) { throw new Error('退勤時刻は出勤時刻以降にしてください'); }
    const endOfDay = Utilities.parseDate(dateKey + ' 23:59', TIME_ZONE, 'yyyy/MM/dd HH:mm');
    endOfDay.setSeconds(59, 999);
    const maxBreak = Math.floor(((clockOut || endOfDay).getTime() - clockIn.getTime()) / 60000);
    if (breakMinutes > maxBreak) { throw new Error('休憩時間が勤務時間を超えています'); }

    let nextRow = sheet.getLastRow() + 1;
    const clockIns = events.filter(function (event) { return event.type === '出勤'; }).sort(compareClockEvents_);
    const clockOuts = events.filter(function (event) { return event.type === '退勤'; }).sort(compareClockEvents_);
    const edited = events.filter(function (event) { return event.type !== '出勤' && event.type !== '退勤'; })
      .map(function (event) { return Object.assign({}, event); });
    const inEvent = { employeeId: payload.id, type: '出勤', time: clockIn,
      rowNumber: clockIns.length ? clockIns[0].rowNumber : nextRow++, breakAdjustment: 0 };
    edited.push(inEvent);
    if (clockOut) {
      // 休憩中の退勤漏れを修正した場合は、その退勤時刻で休憩も終了させる。
      if (original.breakOngoing) {
        edited.push({ employeeId: payload.id, type: '休憩終了', time: clockOut, rowNumber: nextRow++, breakAdjustment: 0 });
      }
      edited.push({ employeeId: payload.id, type: '退勤', time: clockOut,
        rowNumber: clockOuts.length ? clockOuts[clockOuts.length - 1].rowNumber : nextRow++, breakAdjustment: 0 });
    }
    const state = getClockStateFromEvents_(edited);
    if (state.state === 'invalid') { throw new Error(state.error); }
    const rawSummary = summarizeDayEvents(edited.slice());
    inEvent.breakAdjustment = breakMinutes - rawSummary.breakMinutes;
    const rate = getEmployeeRateMap()[String(payload.id)] ||
      { base: BASE_HOURLY_WAGE, holidayBonus: HOLIDAY_HOURLY_BONUS, skillAllowance: SKILL_ALLOWANCE_DEFAULT };
    const result = buildAttendanceView_(payload.id, dateKey, edited.slice().sort(function (a, b) { return a.rowNumber - b.rowNumber; }),
      names[String(payload.id)], rate, isHolidayDate(dateKey));
    writeAttendanceEvents_(sheet, events, edited);
    return { attendance: result };
  });
}

function writeAttendanceEvents_(sheet, original, edited) {
  const adjustmentColumn = getAttendanceAdjustmentColumn_(sheet, true);
  const byRow = new Map();
  original.forEach(function (event) { byRow.set(event.rowNumber, null); });
  edited.forEach(function (event) { byRow.set(event.rowNumber, event); });
  const lastRow = Math.max.apply(null, Array.from(byRow.keys()));
  if (lastRow > sheet.getMaxRows()) { sheet.insertRowsAfter(sheet.getMaxRows(), lastRow - sheet.getMaxRows()); }
  const writes = [];
  byRow.forEach(function (event, rowNumber) {
    const values = event ? [event.employeeId, event.type, event.time] : ['', '', ''];
    const adjustment = event ? event.breakAdjustment : '';
    [[sheet.getRange(rowNumber, 1, 1, 3), [values]],
      [sheet.getRange(rowNumber, adjustmentColumn), [[adjustment]]]].forEach(function (entry) {
      const before = entry[0].getValues();
      if (JSON.stringify(before) === JSON.stringify(entry[1])) { return; }
      const formulas = entry[0].getFormulas();
      const restore = before.map(function (row, r) { return row.map(function (value, c) { return formulas[r][c] || value; }); });
      writes.push({ range: entry[0], values: entry[1], before: restore });
    });
  });
  try {
    writes.forEach(function (write) { write.range.setValues(write.values); });
    SpreadsheetApp.flush();
  } catch (error) {
    let restoreFailed = false;
    writes.forEach(function (write) {
      try { write.range.setValues(write.before); } catch (restoreError) { restoreFailed = true; }
    });
    try { SpreadsheetApp.flush(); } catch (restoreError) { restoreFailed = true; }
    throw new Error(error.message + (restoreFailed ? '（復元できなかった可能性があります。打刻履歴を確認してください）' : ''));
  }
}
