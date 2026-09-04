/**
 * シート名の定数
 * シートの並び順が変わっても壊れないよう、番号ではなく名前でシートを参照する
 */
var SHEET_NAMES = {
  CHECK_RESULT: 'チェック結果', // 現状どこからも参照していない
  EMPLOYEES: '従業員名簿',
  TIME_CLOCKS: '打刻履歴',
  HOLIDAYS: '祝日', // A列に祝日の日付を1行1件登録する。シート自体が無い場合は祝日なし（土日のみ休日加給）として扱う
  SKILLS: 'スキル管理', // A:従業員ID B:氏名 C列以降:評価と更新日のペア（SKILL_ITEMS と同じ並び順、詳細は SKILL_SHEET_HEADER を参照）
  TODAY_WORK: '本日の業務' // A:日付 B:従業員ID C:ポジション D:確認事項（複数件を改行区切りで1セルに保存）。1人・1日につき1行
};

/**
 * 管理するスキル項目
 * 「スキル管理」シートの評価列（C,E,G,I,K,M列）の並び順と一致させる
 * 各スキルの評価列のすぐ右（D,F,H,J,L,N列）にはそのスキルの更新日が入る
 */
var SKILL_ITEMS = ['接客', 'レジ', 'バー', '厨房', '肉の仕込み', '締め作業'];

/**
 * 「スキル管理」シートの1行目（ヘッダー）
 * 評価と更新日を同じセルに混在させず、必ず別列にする
 * （教育ペア提示・シフト作成支援・代替要員判断などから ◎/○/△/× をそのまま判定したいため）
 */
var SKILL_SHEET_HEADER = [
  '従業員ID', '氏名',
  '接客評価', '更新日',
  'レジ評価', '更新日',
  'バー評価', '更新日',
  '厨房評価', '厨房更新日',
  '肉の仕込み評価', '更新日',
  '締め作業評価', '更新日'
];

/**
 * スキルの評価値として受け付ける値。この4種類以外は保存できない
 * 配列の最後の値（×）を未評価の初期値として扱う
 */
var SKILL_LEVELS = ['◎', '○', '△', '×'];

/**
 * 「スキル管理」シートの1行目が SKILL_SHEET_HEADER と一致していることを保証する
 * 1行目が空、または内容が SKILL_SHEET_HEADER と1つでも異なる場合にのみ1行目を書き直す
 * 既に正しい場合は何もしない。2行目以降の既存データ（評価・更新日）には一切触れない
 */
function ensureSkillSheetHeader(sheet) {
  var headerRange = sheet.getRange(1, 1, 1, SKILL_SHEET_HEADER.length);
  var currentHeader = headerRange.getValues()[0];
  for (var i = 0; i < SKILL_SHEET_HEADER.length; i++) {
    if (currentHeader[i] !== SKILL_SHEET_HEADER[i]) {
      headerRange.setValues([SKILL_SHEET_HEADER]);
      return;
    }
  }
}

/**
 * 「スキル管理」シートを取得し、1行目のヘッダーが正しいことを保証してから返す
 * スキル管理機能（一覧の読み込み・保存）はシートの取得に必ずこの関数を使うことで、
 * migrateSkillSheetToDatedColumns() を手動実行しなくても1行目が SKILL_SHEET_HEADER と一致する
 */
function getSkillSheet() {
  var sheet = getSheet(SHEET_NAMES.SKILLS);
  ensureSkillSheetHeader(sheet);
  return sheet;
}

/**
 * 打刻時刻や日付の表示に使うタイムゾーン
 * スプレッドシートに保存された日時をどの地域の時刻として読むかを揃えるため、
 * 日時を文字列にする箇所ではこの定数を必ず指定する
 */
var TIME_ZONE = 'Asia/Tokyo';

/**
 * 打刻がまだ無いことを表す表示文字列
 */
var UNRECORDED_LABEL = '未打刻';

/**
 * 労働時間を計算できないことを表す表示文字列
 */
var NO_VALUE_LABEL = '-';

/**
 * 打刻種別として受け付ける値
 * 休憩の開始・終了も出勤・退勤と同じ打刻履歴シートに記録する
 */
var CLOCK_TYPES = ['出勤', '退勤', '休憩開始', '休憩終了'];

/**
 * シート名からシートを取得する
 * 見つからない場合はシート名を含むエラーを投げ、設定ミスをその場で気付けるようにする
 */
function getSheet(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('シートが見つかりません: ' + sheetName);
  }
  return sheet;
}

/**
 * ページを開いた時に最初に呼ばれるルートメソッド
 * 「店舗運営支援アプリ」の単一ページを表示する
 * 従業員画面と店長画面はこの1ページの中でJavaScriptにより表示を切り替える
 */
function doGet() {
  return HtmlService.createTemplateFromFile("view_home")
    .evaluate().setTitle("店舗運営支援アプリ")
}

/**
 * 従業員一覧
 * 従業員番号（A列）と名前（B列）が両方とも入力されている行だけを返す
 */
function getEmployees() {
  var empSheet = getSheet(SHEET_NAMES.EMPLOYEES)
  var last_row = empSheet.getLastRow()
  var empRange = empSheet.getRange(2, 1, last_row, 2);// シートの中のヘッダーを除く範囲を取得
  var empValues = empRange.getValues();// [従業員番号, 名前] の配列をまとめて取得
  var employees = [];
  for (var i = 0; i < empValues.length; i++) {
    var empId = empValues[i][0];
    var empName = empValues[i][1];
    if (empId === "" || empName === "") { // 従業員番号か名前が空の行は一覧に含めない
      continue;
    }
    employees.push({
      'id': empId,
      'name': empName
    })
  }
  return employees
}

/**
 * 新しい従業員を名簿に追加する
 * 従業員番号（A列）は既存の入力済み番号のうち、名前（B列）が空欄になっている
 * 最小の番号を使用する。新規追加した従業員は必ずアルバイト（C列に○）として登録する
 */
function addEmployee(lastName, firstName) {
  lastName = String(lastName).trim();
  firstName = String(firstName).trim();
  if (lastName === "" || firstName === "") {
    throw new Error('姓と名を入力してください')
  }

  var empSheet = getSheet(SHEET_NAMES.EMPLOYEES)
  var last_row = empSheet.getLastRow()
  var empRange = empSheet.getRange(2, 1, last_row, 2);
  var empValues = empRange.getValues();

  var targetRow = -1;// 空き番号が見つかったシート上の行番号
  for (var i = 0; i < empValues.length; i++) {
    var empId = empValues[i][0];
    var empName = empValues[i][1];
    if (empId !== "" && empName === "") {
      targetRow = i + 2;// ヘッダー行の分だけシート上の行番号は配列の添字より2大きい
      break;
    }
  }

  if (targetRow === -1) {
    throw new Error('空いている従業員番号がありません')
  }

  var partTimeMark = '\u25CB'; // U+25CB WHITE CIRCLE。既存データの○と完全に一致させるため、コードポイントで指定して誤字を防ぐ

  empSheet.getRange(targetRow, 2).setValue(lastName + ' ' + firstName)
  empSheet.getRange(targetRow, 3).setValue(partTimeMark)

  return '従業員を追加しました'
}

/**
 * 出勤・退勤の打刻をスプレッドシートに記録する
 */
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

/**
 * 打刻履歴シートのヘッダーを除く全データを消去する保守用の関数
 * 動作確認のたびに履歴をまっさらにするためのもので、Webアプリ（画面）からは呼び出していない。
 * 使う場合はスクリプトエディタから手動で実行する
 */
function resetClockRecords() {
  var sheet = getSheet(SHEET_NAMES.TIME_CLOCKS);
  var lastRow = sheet.getLastRow();

  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
  }
}

/**
 * 本日の勤怠一覧を返す（店長画面の「勤怠管理」から呼び出される）
 * recordClock が打刻履歴シートに残した記録のうち当日分だけを従業員ごとにまとめ、
 * 出勤・退勤・休憩・実労働時間・概算給与を表示用の文字列にして返す
 *
 * 同じ従業員が1日に何度も打刻していた場合（押し間違いや押し直し）は、
 * 最も早い出勤と最も遅い退勤を採用してその日の勤務時間帯とみなす
 * 休憩は開始・終了を時刻順にペアにして合計し、勤務時間から差し引く
 */
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

/**
 * 指定した日（'yyyy/MM/dd' 形式）の勤怠一覧を組み立てる
 * 打刻が1件も無い従業員は一覧に含めない
 */
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

/**
 * 1人の従業員の1日分の打刻イベントをまとめる
 * 出勤は最も早い時刻、退勤は最も遅い時刻を採用する
 * 休憩は「休憩開始」→「休憩終了」を時刻順にペアにして合計する
 * 「休憩開始」のまま「休憩終了」が無い場合は休憩中とみなし、その区間は合計に含めない
 */
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

/**
 * 実労働時間（分）を計算する
 * 出勤・退勤が揃っていない場合や、休憩を差し引くと負になる場合（打刻が壊れている）は null を返す
 */
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

/**
 * 従業員番号ごとの時給を対応表にする
 * 従業員名簿のD列（基本時給）・E列（休日加給）を読む。数値以外や空欄は0円として扱う
 */
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

/**
 * 対象日（'yyyy/MM/dd' 形式）が休日加給の対象かどうかを判定する
 * 土日は常に対象。加えて「祝日」シートのA列に登録された日付も対象とする
 * 「祝日」シートが無い場合は土日のみで判定する（祝日機能が未設定でも既存機能は壊さない）
 */
function isHolidayDate(targetDate) {
  var date = new Date(targetDate);
  var isWeekend = (date.getDay() === 0 || date.getDay() === 6);
  if (isWeekend) {
    return true;
  }
  return !!getHolidayDateSet()[targetDate];
}

/**
 * 「祝日」シートのA列から、祝日一覧を { 'yyyy/MM/dd': true } の形にして返す
 * 日付として入力されていない行（見出しなど）は無視する
 */
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

/**
 * 従業員番号から名前を引くための対応表を作る
 * 従業員番号は数値でも文字列でもキーとして同じ扱いになるよう、文字列に揃えて格納する
 */
function getEmployeeNameMap() {
  var employees = getEmployees();
  var nameByEmpId = {};
  for (var i = 0; i < employees.length; i++) {
    nameByEmpId[String(employees[i]['id'])] = employees[i]['name'];
  }
  return nameByEmpId;
}

/**
 * 打刻時刻を「17:00」の形式にする
 * まだ打刻されていない場合は「未打刻」を返す
 */
function formatClockTime(clockedAt) {
  if (clockedAt === null) {
    return UNRECORDED_LABEL;
  }
  return Utilities.formatDate(clockedAt, TIME_ZONE, 'HH:mm');
}

/**
 * 分数を「5時間30分」の形式にする
 * null の場合は計算できないので「-」を返す
 */
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

/**
 * 休憩時間の表示文字列を作る
 * 休憩開始のまま終了していない場合は「休憩中」、休憩が無ければ「-」を返す
 */
function formatBreakTime(breakMinutes, breakOngoing) {
  if (breakOngoing) {
    return '休憩中';
  }
  if (breakMinutes === 0) {
    return NO_VALUE_LABEL;
  }
  return formatMinutesLabel(breakMinutes);
}

/**
 * 概算給与の表示文字列を作る（実労働時間 × 適用時給、円未満切り捨て）
 * 休日加給は、対象日が休日加給の対象日（土日祝）の場合にのみ基本時給へ上乗せする
 * 実労働時間が計算できない場合は「-」を返す
 */
function formatEstimatedWage(actualMinutes, rate, isHolidayRate) {
  if (actualMinutes === null) {
    return NO_VALUE_LABEL;
  }
  var hourlyRate = rate.base + (isHolidayRate ? rate.holidayBonus : 0);
  var wage = Math.floor(actualMinutes / 60 * hourlyRate);
  return wage.toLocaleString('ja-JP') + '円';
}

/**
 * 全スキル項目を未評価（×）で埋めたスキルオブジェクトを作る
 * 「スキル管理」シートにまだ行が無い従業員の初期値として使う
 */
function getDefaultSkills() {
  var skills = {};
  var defaultLevel = SKILL_LEVELS[SKILL_LEVELS.length - 1]; // ×
  for (var i = 0; i < SKILL_ITEMS.length; i++) {
    skills[SKILL_ITEMS[i]] = defaultLevel;
  }
  return skills;
}

/**
 * スキル項目（SKILL_ITEMS の添字）から、「スキル管理」シート上の評価列（1始まり）を求める
 * 評価列のすぐ右が、その項目の更新日列になる
 */
function getSkillEvalColumn(skillIndex) {
  return 3 + skillIndex * 2; // 接客=C(3), レジ=E(5), バー=G(7), 厨房=I(9), 肉の仕込み=K(11), 締め作業=M(13)
}

/**
 * 従業員番号ごとのスキル評価を対応表にする
 * 「スキル管理」シートの評価列（C,E,G,I,K,M列。SKILL_ITEMS の並び順）を読み、
 * 項目名をキーにしたオブジェクトにして従業員番号ごとにまとめる
 * ◎/○/△/× 以外の値が入っていた場合は不正な入力とみなし、未評価（×）として扱う
 * 更新日（各評価列の右隣）はここでは扱わない。更新日が必要な場合は
 * getEmployeeSkillUpdatedAtMap() を使う
 *
 * 教育ペア提示・シフト作成支援・代替要員判断など、今後の機能から
 * 「スキル評価そのもの」を再利用する際は、この関数を呼び出せば良い
 */
function getEmployeeSkillMap() {
  var sheet = getSkillSheet();
  var lastRow = sheet.getLastRow();
  var skillMapByEmpId = {};
  if (lastRow < 2) {
    return skillMapByEmpId;
  }

  var values = sheet.getRange(2, 1, lastRow - 1, 2 + SKILL_ITEMS.length * 2).getValues();
  var defaultLevel = SKILL_LEVELS[SKILL_LEVELS.length - 1]; // ×
  for (var i = 0; i < values.length; i++) {
    var empId = values[i][0];
    if (empId === '') {
      continue;
    }
    var skills = {};
    for (var s = 0; s < SKILL_ITEMS.length; s++) {
      var rawValue = values[i][getSkillEvalColumn(s) - 1]; // 配列の添字は列番号-1
      skills[SKILL_ITEMS[s]] = (SKILL_LEVELS.indexOf(rawValue) !== -1) ? rawValue : defaultLevel;
    }
    skillMapByEmpId[String(empId)] = skills;
  }
  return skillMapByEmpId;
}

/**
 * 従業員番号ごとのスキル評価の更新日を対応表にする
 * 「スキル管理」シートの各評価列のすぐ右（D,F,H,J,L,N列）を読み、
 * 項目名をキーにした「yyyy/MM/dd」形式の文字列にして従業員番号ごとにまとめる
 * まだ一度も保存されていない項目（日付が入っていない）は空文字にする
 *
 * 店長画面でスキル評価の横に最終更新日を表示するために使う
 */
function getEmployeeSkillUpdatedAtMap() {
  var sheet = getSkillSheet();
  var lastRow = sheet.getLastRow();
  var updatedAtMapByEmpId = {};
  if (lastRow < 2) {
    return updatedAtMapByEmpId;
  }

  var values = sheet.getRange(2, 1, lastRow - 1, 2 + SKILL_ITEMS.length * 2).getValues();
  for (var i = 0; i < values.length; i++) {
    var empId = values[i][0];
    if (empId === '') {
      continue;
    }
    var updatedAt = {};
    for (var s = 0; s < SKILL_ITEMS.length; s++) {
      var rawDate = values[i][getSkillEvalColumn(s)]; // 評価列の配列添字+1 = 更新日列
      updatedAt[SKILL_ITEMS[s]] = formatSkillUpdatedAt(rawDate);
    }
    updatedAtMapByEmpId[String(empId)] = updatedAt;
  }
  return updatedAtMapByEmpId;
}

/**
 * スキル更新日セルの値を「yyyy/MM/dd」形式の文字列にする
 * 日付が入っていない場合（未保存、または不正な入力）は空文字を返す
 */
function formatSkillUpdatedAt(rawDate) {
  if (rawDate instanceof Date) {
    return Utilities.formatDate(rawDate, TIME_ZONE, 'yyyy/MM/dd');
  }
  return '';
}

/**
 * 全スキル項目を「未保存（空文字）」で埋めた更新日オブジェクトを作る
 * 「スキル管理」シートにまだ行が無い従業員の初期値として使う
 */
function getDefaultSkillUpdatedAt() {
  var updatedAt = {};
  for (var i = 0; i < SKILL_ITEMS.length; i++) {
    updatedAt[SKILL_ITEMS[i]] = '';
  }
  return updatedAt;
}

/**
 * 店長画面の「スキル管理」から呼び出される一覧
 * 従業員名簿（getEmployees）を基準に、各従業員のスキル評価と更新日をまとめて返す
 * 「スキル管理」シートにまだ行が無い従業員は、初期値（すべて×、更新日は空）として表示する
 * 名簿に存在しない従業員がスキル管理シートに残っていても一覧には含めない
 */
function getSkillMatrix() {
  var employees = getEmployees();
  var skillMapByEmpId = getEmployeeSkillMap();
  var updatedAtMapByEmpId = getEmployeeSkillUpdatedAtMap();

  var matrix = [];
  for (var i = 0; i < employees.length; i++) {
    var empId = employees[i]['id'];
    var key = String(empId);
    matrix.push({
      'id': empId,
      'name': employees[i]['name'],
      'skills': skillMapByEmpId[key] || getDefaultSkills(),
      'updatedAt': updatedAtMapByEmpId[key] || getDefaultSkillUpdatedAt()
    });
  }
  return matrix;
}

/**
 * 「スキル管理」シートの中から、指定した従業員番号の行番号を探す
 * 見つからない場合は -1 を返す
 */
function findSkillRowByEmpId(sheet, empId) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return -1;
  }
  var empIdValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < empIdValues.length; i++) {
    if (String(empIdValues[i][0]) === String(empId)) {
      return i + 2; // ヘッダー行の分だけシート上の行番号は配列の添字より2大きい
    }
  }
  return -1;
}

/**
 * 1人分のスキル評価をまとめて保存する（店長画面の各行の「保存」ボタンから呼び出される）
 * skillValues は SKILL_ITEMS と同じ並び順の配列（接客, レジ, バー, 厨房, 肉の仕込み, 締め作業）
 *
 * データの識別には氏名ではなく従業員番号を使う。氏名は表示用として書き込むだけで、
 * 検索・照合には使わない
 * findSkillRowByEmpId() で既存行が見つかった場合はその行を更新し、新しい行は追加しない
 * （＝従業員1人につき常に1行）
 *
 * 更新日は、実際に評価が変わった項目だけ今日の日付に変更する
 * 評価が変わっていない項目の更新日はそのまま残す。行が新規作成された場合も、
 * 画面上の初期値である「×」を基準にして判定する（＝×のままの項目は初回保存でも
 * 「まだ評価されていない初期状態」として更新日を空欄のままにし、×以外に変更された
 * 項目だけ更新日を今日にする）
 *
 * 戻り値には保存後の各項目の更新日（updatedAt）を含める。店長画面はこれを使って
 * 保存直後の表示（更新日）をサーバーへ再問い合わせせずに更新する
 */
function saveEmployeeSkills(empId, skillValues) {
  var nameByEmpId = getEmployeeNameMap();
  var empName = nameByEmpId[String(empId)];
  if (!empName) {
    throw new Error('従業員名簿に存在しない従業員番号です: ' + empId);
  }
  if (!skillValues || skillValues.length !== SKILL_ITEMS.length) {
    throw new Error('スキルの項目数が不正です');
  }
  for (var i = 0; i < skillValues.length; i++) {
    if (SKILL_LEVELS.indexOf(skillValues[i]) === -1) {
      throw new Error('不正なスキル評価値です: ' + skillValues[i]);
    }
  }

  var sheet = getSkillSheet();
  var targetRow = findSkillRowByEmpId(sheet, empId);
  var isNewRow = (targetRow === -1);
  if (isNewRow) {
    targetRow = sheet.getLastRow() + 1;
    sheet.getRange(targetRow, 1).setValue(empId);
  }
  sheet.getRange(targetRow, 2).setValue(empName);

  // 既存行の場合のみ、現在の評価・更新日を読んで変更の有無を判定する
  // 新規行の場合は、画面の初期値と同じ「×」を現在の評価とみなして判定する
  var currentRowValues = isNewRow ? null : sheet.getRange(targetRow, 3, 1, SKILL_ITEMS.length * 2).getValues()[0];
  var initialLevel = SKILL_LEVELS[SKILL_LEVELS.length - 1]; // ×
  var today = new Date();
  var updatedAt = {};

  for (var s = 0; s < SKILL_ITEMS.length; s++) {
    var evalCol = getSkillEvalColumn(s);
    var newValue = skillValues[s];
    var currentValue = isNewRow ? initialLevel : currentRowValues[evalCol - 3]; // 評価列の配列添字（列3=添字0基準）
    var currentDate = isNewRow ? null : currentRowValues[evalCol - 3 + 1];
    var isChanged = (currentValue !== newValue);

    // 評価自体は、新規行なら（×のままの項目も含めて）常に書き込む
    if (isNewRow || isChanged) {
      sheet.getRange(targetRow, evalCol).setValue(newValue);
    }

    if (isChanged) {
      sheet.getRange(targetRow, evalCol + 1).setValue(today);
      updatedAt[SKILL_ITEMS[s]] = formatSkillUpdatedAt(today);
    } else {
      updatedAt[SKILL_ITEMS[s]] = formatSkillUpdatedAt(currentDate);
    }
  }

  return {
    'message': empName + 'のスキルを保存しました',
    'updatedAt': updatedAt
  };
}

/**
 * 「スキル管理」シートを新しい列構成（評価と更新日を別列にするレイアウト）へ整える保守用の関数
 * 動作確認のたびに使うものではなく、レイアウト変更を反映するときにスクリプトエディタから手動で実行する
 *
 * ・シートが空、またはヘッダーのみの場合：1行目を SKILL_SHEET_HEADER で書き込むだけ
 * ・旧レイアウト（A:従業員ID B:氏名 C〜H:評価のみ、更新日列なし）のデータが残っている場合：
 *   各評価値はそのままの列位置から新しい評価列へ移し、更新日列は空欄にする
 *   （＝旧データはいつ最後に変更されたか記録が無いため「更新日不明」として扱う。
 *   　次にその項目を保存した時点で今日の日付が入る）
 * ・既に新レイアウトで保存されている場合は、ヘッダーを書き直すだけで データはそのまま
 */
function migrateSkillSheetToDatedColumns() {
  var sheet = getSheet(SHEET_NAMES.SKILLS);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 2) {
    sheet.getRange(1, 1, 1, SKILL_SHEET_HEADER.length).setValues([SKILL_SHEET_HEADER]);
    return 'ヘッダーのみ更新しました（既存データはありません）';
  }

  var newColumnCount = 2 + SKILL_ITEMS.length * 2;
  if (lastCol >= newColumnCount) {
    // 既に評価・更新日の列数が揃っている＝新レイアウト済みとみなし、ヘッダーだけ書き直す
    sheet.getRange(1, 1, 1, SKILL_SHEET_HEADER.length).setValues([SKILL_SHEET_HEADER]);
    return 'ヘッダーを更新しました（データは新レイアウト済みのため変更していません）';
  }

  // 旧レイアウト（A:従業員ID B:氏名 C〜H:評価のみ）とみなして変換する
  var oldValues = sheet.getRange(2, 1, lastRow - 1, 2 + SKILL_ITEMS.length).getValues();
  var newValues = [];
  for (var i = 0; i < oldValues.length; i++) {
    var row = oldValues[i];
    var newRow = [row[0], row[1]]; // 従業員ID, 氏名
    for (var s = 0; s < SKILL_ITEMS.length; s++) {
      newRow.push(row[2 + s]); // 評価はそのまま引き継ぐ
      newRow.push(''); // 更新日は不明なので空欄
    }
    newValues.push(newRow);
  }

  sheet.getRange(1, 1, 1, SKILL_SHEET_HEADER.length).setValues([SKILL_SHEET_HEADER]);
  sheet.getRange(2, 1, newValues.length, newColumnCount).setValues(newValues);
  return newValues.length + '件のデータを新しい列構成（評価・更新日を別列）に変換しました';
}

/**
 * ============================================================
 * STEP4: 本日の業務（ポジション・確認事項）
 * ============================================================
 * 店長がダッシュボードから各従業員の「今日の主担当ポジション」と「確認事項」を設定し、
 * 従業員は出勤直後にそれを閲覧できる機能
 */

/**
 * ポジション候補
 * HTML側（店長ダッシュボードのselect）へは view_home.html のテンプレート内で
 * JSON.stringify(POSITION_ITEMS) として埋め込み、この配列と別々にハードコードしない
 * 「未設定」はこの配列に含めず、空文字で表す
 */
var POSITION_ITEMS = ['ホール', '厨房', '肉の仕込み', '締め作業'];

/**
 * 「本日の業務」シートの1行目（ヘッダー）
 */
var TODAY_WORK_SHEET_HEADER = ['日付', '従業員ID', 'ポジション', '確認事項'];

/**
 * 「本日の業務」シートを取得する。シートが存在しない場合は作成し、1行目へヘッダーを設定する
 * 既に存在し、ヘッダーが正しい場合は2行目以降の既存データに一切触れない
 * getSheet() は見つからない場合にエラーを投げる仕様のため使わず、ここでは自動作成のために
 * SpreadsheetApp を直接扱う（getSheet() 自体の仕様は変更しない）
 */
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

/**
 * 今日の日付を「本日の業務」シートの日付列と比較するためのキー（'yyyy/MM/dd'）にする
 */
function getTodayDateKey() {
  return Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy/MM/dd');
}

/**
 * 「本日の業務」シートの中から、指定した日付＋従業員番号（複合キー）の行番号を探す
 * 見つからない場合は -1 を返す
 */
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

/**
 * 指定した日付（'yyyy/MM/dd'）の「本日の業務」設定を、従業員番号をキーにした対応表にする
 * { position: 'ホール', noticesText: 'カルビ残量：少\n牛タン消費期限：明日' } の形
 * getTodayWorkSettings()（店長ダッシュボード）と getTodayWorkInfo()（従業員画面）の両方から使う
 */
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

/**
 * 確認事項のテキスト（textareaの1行につき1件）を改行で分割し、箇条書き用の配列にする
 * 前後の空白を取り除いたうえで、空行は結果に含めない
 */
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

/**
 * 店長ダッシュボードの「本日の業務設定」から呼び出される一覧
 * 従業員名簿（getEmployees）を基準に、本日すでに設定済みの従業員は保存済みの内容を、
 * まだ設定していない従業員はポジション未設定・確認事項空欄として返す
 */
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

/**
 * 1人分の本日の業務（ポジション・確認事項）を保存する（店長ダッシュボードの各行の「保存」ボタンから呼び出される）
 *
 * データの識別には氏名ではなく従業員番号を使う。「今日の日付＋従業員番号」を複合キーとして
 * findTodayWorkRow() で既存行を探し、見つかればその行のポジション・確認事項だけを上書きする
 * （＝新しい行を追加しない）。見つからない場合だけ新しい行を追加する
 * （＝同じ日・同じ従業員につき常に1行）
 */
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

/**
 * 従業員画面から、出勤直後に1人分の本日の業務情報を取得する
 * 手入力の確認事項（noticesText）を箇条書き配列に分割して返す
 *
 * 将来的に食材管理の在庫アラートや業務引継ぎ事項もここへ合流させる想定のため、
 * notices は「手入力の確認事項だけの配列」ではなく「表示すべき確認事項すべての配列」
 * として扱えるよう、この関数の戻り値の形（position, notices）を維持したまま
 * 中身を積み増していく設計にする。今回は手入力分のみを notices に入れる
 */
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
