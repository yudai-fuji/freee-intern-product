// 店長向けのシフト作成（日別のシフト割当・必要人数条件・確認事項）を管理する

// 画面とGASの間でやり取りする日付キー（yyyy/MM/dd）と時刻（HH:mm）の形式
const SHIFT_DATE_KEY_PATTERN = /^\d{4}\/\d{2}\/\d{2}$/;
const SHIFT_TIME_PATTERN = /^([01][0-9]|2[0-3]):[0-5][0-9]$/;

// シフトシートの列番号（1始まり。SHIFT_SHEET_HEADER の並びと揃える）
const SHIFT_COL_DATE = 1;
const SHIFT_COL_EMP_ID = 2;
const SHIFT_COL_START_TIME = 3;
const SHIFT_COL_END_TIME = 4;
const SHIFT_COL_POSITION = 5;
const SHIFT_COL_NOTE = 6;

// シフト必要人数シートの列番号（1始まり。SHIFT_REQUIREMENT_SHEET_HEADER の並びと揃える）
const SHIFT_REQ_COL_DATE = 1;
const SHIFT_REQ_COL_START_TIME = 2;
const SHIFT_REQ_COL_END_TIME = 3;
const SHIFT_REQ_COL_POSITION = 4;
const SHIFT_REQ_COL_REQUIRED_COUNT = 5;

/**
 * シフト関連のシートが無ければ作成し、ヘッダーを既定の列構成に揃える
 * 「17:00」のような時刻をシート側が時刻値へ自動変換すると読み書きがずれるため、時刻の2列は文字列として扱う
 */
function ensureShiftSheetByName(sheetName, header, firstTimeColumn) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  let isNewSheet = false;
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    isNewSheet = true;
  }

  const headerRange = sheet.getRange(1, 1, 1, header.length);
  const currentHeader = headerRange.getValues()[0];
  for (let i = 0; i < header.length; i++) {
    if (currentHeader[i] !== header[i]) {
      headerRange.setValues([header]);
      break;
    }
  }

  // 作成直後だけ時刻列全体を文字列書式にする（既存シートへの追記時は appendShiftRows() 側で書式を指定する）
  const dataRowCount = sheet.getMaxRows() - 1;
  if (isNewSheet && dataRowCount > 0) {
    sheet.getRange(2, firstTimeColumn, dataRowCount, 2).setNumberFormat('@');
  }
  return sheet;
}

function ensureShiftSheet() {
  return ensureShiftSheetByName(SHEET_NAMES.SHIFTS, SHIFT_SHEET_HEADER, SHIFT_COL_START_TIME);
}

function ensureShiftRequirementSheet() {
  return ensureShiftSheetByName(SHEET_NAMES.SHIFT_REQUIREMENTS, SHIFT_REQUIREMENT_SHEET_HEADER, SHIFT_REQ_COL_START_TIME);
}

/**
 * シートに入っている日付（Date でも文字列でも）を 'yyyy/MM/dd' の日付キーに揃える
 */
function toShiftDateKey(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, TIME_ZONE, 'yyyy/MM/dd');
  }
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

/**
 * 画面やシートから受け取った時刻を 'HH:mm' に揃える
 * シートが時刻値へ変換してしまった Date や、先頭の0が落ちた「9:05」のような文字列も受け付ける
 */
function toShiftTimeText(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, TIME_ZONE, 'HH:mm');
  }
  if (value === null || value === undefined) {
    return '';
  }
  const text = String(value).trim();
  const matched = text.match(/^(\d{1,2}):(\d{2})/);
  if (!matched) {
    return text;
  }
  return ('0' + matched[1]).slice(-2) + ':' + matched[2];
}

/**
 * 'HH:mm' を0時からの分数に変換する。時刻として読めない場合は null を返す
 */
function shiftTimeToMinutes(timeText) {
  if (!SHIFT_TIME_PATTERN.test(timeText)) {
    return null;
  }
  return Number(timeText.slice(0, 2)) * 60 + Number(timeText.slice(3, 5));
}

function formatShiftMinutes(totalMinutes) {
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  return ('0' + hours).slice(-2) + ':' + ('0' + minutes).slice(-2);
}

function formatShiftTimeRange(startMinutes, endMinutes) {
  return formatShiftMinutes(startMinutes) + '〜' + formatShiftMinutes(endMinutes);
}

/**
 * 日付キーが 'yyyy/MM/dd' 形式で、かつ実在する日付なら Date を返し、そうでなければ null を返す
 * 実在しない日付（2026/02/31 など）は Date が翌月へ繰り上がるため、作り直して突き合わせる
 */
function parseShiftDateKey(dateKey) {
  if (!SHIFT_DATE_KEY_PATTERN.test(dateKey)) {
    return null;
  }
  const parts = dateKey.split('/');
  const year = Number(parts[0]);
  const month = Number(parts[1]);
  const day = Number(parts[2]);
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
    return null;
  }
  return date;
}

/**
 * シフトシートの全行を読み、日付キーごとの割当一覧にまとめる
 */
function readShiftAssignmentsByDate() {
  const sheet = ensureShiftSheet();
  const lastRow = sheet.getLastRow();
  const assignmentsByDate = {};
  if (lastRow < 2) {
    return assignmentsByDate;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, SHIFT_SHEET_HEADER.length).getValues();
  for (let i = 0; i < values.length; i++) {
    const dateKey = toShiftDateKey(values[i][SHIFT_COL_DATE - 1]);
    const empId = values[i][SHIFT_COL_EMP_ID - 1];
    if (dateKey === '' || empId === '') {
      continue; // 日付か従業員IDが空の行は読み飛ばす
    }
    if (!assignmentsByDate[dateKey]) {
      assignmentsByDate[dateKey] = [];
    }
    assignmentsByDate[dateKey].push({
      'employeeId': String(empId),
      'startTime': toShiftTimeText(values[i][SHIFT_COL_START_TIME - 1]),
      'endTime': toShiftTimeText(values[i][SHIFT_COL_END_TIME - 1]),
      'position': String(values[i][SHIFT_COL_POSITION - 1] || ''),
      'note': String(values[i][SHIFT_COL_NOTE - 1] || '')
    });
  }
  return assignmentsByDate;
}

/**
 * シフト必要人数シートの全行を読み、日付キーごとの必要人数条件一覧にまとめる
 */
function readShiftRequirementsByDate() {
  const sheet = ensureShiftRequirementSheet();
  const lastRow = sheet.getLastRow();
  const requirementsByDate = {};
  if (lastRow < 2) {
    return requirementsByDate;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, SHIFT_REQUIREMENT_SHEET_HEADER.length).getValues();
  for (let i = 0; i < values.length; i++) {
    const dateKey = toShiftDateKey(values[i][SHIFT_REQ_COL_DATE - 1]);
    const position = String(values[i][SHIFT_REQ_COL_POSITION - 1] || '');
    if (dateKey === '' || position === '') {
      continue; // 日付かポジションが空の行は読み飛ばす
    }
    if (!requirementsByDate[dateKey]) {
      requirementsByDate[dateKey] = [];
    }
    requirementsByDate[dateKey].push({
      'startTime': toShiftTimeText(values[i][SHIFT_REQ_COL_START_TIME - 1]),
      'endTime': toShiftTimeText(values[i][SHIFT_REQ_COL_END_TIME - 1]),
      'position': position,
      'requiredCount': Number(values[i][SHIFT_REQ_COL_REQUIRED_COUNT - 1]) || 0
    });
  }
  return requirementsByDate;
}

/**
 * 割当・必要人数条件のどちらかにデータがある日付キーを、重複なく集めて並べ替える
 */
function collectShiftDateKeys(assignmentsByDate, requirementsByDate) {
  const dateKeys = [];
  const seen = {};
  const maps = [assignmentsByDate, requirementsByDate];
  for (let m = 0; m < maps.length; m++) {
    for (const dateKey in maps[m]) {
      if (!seen[dateKey]) {
        seen[dateKey] = true;
        dateKeys.push(dateKey);
      }
    }
  }
  dateKeys.sort();
  return dateKeys;
}

/**
 * 1日分の割当と必要人数条件から確認事項（保存はできるが店長に確認してほしい内容）を作る
 */
function buildShiftWarnings(assignments, requirements, nameByEmpId) {
  const warnings = [];
  addShortStaffWarnings(assignments, requirements, warnings);
  addOverlapWarnings(assignments, nameByEmpId, warnings);
  addBrokenRowWarnings(assignments, requirements, nameByEmpId, warnings);

  // 吹き出しの中で時間帯の早い順に読めるようにする
  warnings.sort(function (a, b) {
    if (a.time === b.time) {
      return 0;
    }
    return a.time < b.time ? -1 : 1;
  });
  return warnings;
}

/**
 * 必要人数条件の時間帯を最初から最後まで担当でき、かつ同じポジションの従業員を数え、不足していれば確認事項にする
 */
function addShortStaffWarnings(assignments, requirements, warnings) {
  for (let r = 0; r < requirements.length; r++) {
    const requirement = requirements[r];
    const requiredStart = shiftTimeToMinutes(requirement.startTime);
    const requiredEnd = shiftTimeToMinutes(requirement.endTime);
    if (requiredStart === null || requiredEnd === null || requiredStart >= requiredEnd) {
      continue; // 時刻として読めない条件は addBrokenRowWarnings() が別途確認事項にする
    }
    if (!(requirement.requiredCount >= 1)) {
      continue; // 必要人数が読めない条件も同様
    }

    let assignedCount = 0;
    for (let a = 0; a < assignments.length; a++) {
      const assignment = assignments[a];
      if (assignment.position !== requirement.position) {
        continue;
      }
      const startMinutes = shiftTimeToMinutes(assignment.startTime);
      const endMinutes = shiftTimeToMinutes(assignment.endTime);
      if (startMinutes === null || endMinutes === null) {
        continue;
      }
      // 条件の時間帯を通して勤務できる人だけを配置人数として数える（途中出勤・途中退勤は数えない）
      if (startMinutes <= requiredStart && endMinutes >= requiredEnd) {
        assignedCount++;
      }
    }

    if (assignedCount < requirement.requiredCount) {
      warnings.push({
        'time': formatShiftTimeRange(requiredStart, requiredEnd),
        'title': requirement.position + '：' + (requirement.requiredCount - assignedCount) + '人不足',
        'reason': '必要' + requirement.requiredCount + '人／配置' + assignedCount + '人'
      });
    }
  }
}

/**
 * 同じ従業員が同じ日に重なる時間帯で割り当てられていないかを調べる
 */
function addOverlapWarnings(assignments, nameByEmpId, warnings) {
  for (let i = 0; i < assignments.length; i++) {
    for (let j = i + 1; j < assignments.length; j++) {
      const first = assignments[i];
      const second = assignments[j];
      if (String(first.employeeId) !== String(second.employeeId)) {
        continue;
      }
      const firstStart = shiftTimeToMinutes(first.startTime);
      const firstEnd = shiftTimeToMinutes(first.endTime);
      const secondStart = shiftTimeToMinutes(second.startTime);
      const secondEnd = shiftTimeToMinutes(second.endTime);
      if (firstStart === null || firstEnd === null || secondStart === null || secondEnd === null) {
        continue;
      }

      const overlapStart = Math.max(firstStart, secondStart);
      const overlapEnd = Math.min(firstEnd, secondEnd);
      if (overlapStart >= overlapEnd) {
        continue; // 接しているだけ（17:00終わりと17:00始まり）は重複として扱わない
      }

      warnings.push({
        'time': formatShiftTimeRange(overlapStart, overlapEnd),
        'title': getShiftEmployeeLabel(first.employeeId, nameByEmpId) + '：勤務時間の重複',
        'reason': first.startTime + '〜' + first.endTime + ' と ' + second.startTime + '〜' + second.endTime + ' が重なっています'
      });
    }
  }
}

/**
 * 時刻として読めない行・前後が逆の行・存在しない従業員やポジションの行を確認事項にする
 * 画面から保存するときは validateShiftInput() がエラーにするため、主にシートを直接編集された場合への備え
 */
function addBrokenRowWarnings(assignments, requirements, nameByEmpId, warnings) {
  for (let i = 0; i < assignments.length; i++) {
    const assignment = assignments[i];
    const startMinutes = shiftTimeToMinutes(assignment.startTime);
    const endMinutes = shiftTimeToMinutes(assignment.endTime);
    const employeeLabel = getShiftEmployeeLabel(assignment.employeeId, nameByEmpId);

    if (startMinutes === null || endMinutes === null) {
      warnings.push({
        'time': '時刻不明',
        'title': employeeLabel + '：勤務時間を読み取れません',
        'reason': '出勤「' + assignment.startTime + '」／退勤「' + assignment.endTime + '」'
      });
    } else if (startMinutes >= endMinutes) {
      warnings.push({
        'time': formatShiftTimeRange(startMinutes, endMinutes),
        'title': employeeLabel + '：出勤時刻が退勤時刻以降です',
        'reason': assignment.startTime + '〜' + assignment.endTime
      });
    }

    if (!nameByEmpId[String(assignment.employeeId)]) {
      warnings.push({
        'time': '要修正',
        'title': '従業員名簿にない従業員番号です',
        'reason': '従業員番号 ' + assignment.employeeId
      });
    }
    if (POSITION_ITEMS.indexOf(assignment.position) === -1) {
      warnings.push({
        'time': '要修正',
        'title': employeeLabel + '：ポジションが不正です',
        'reason': 'ポジション「' + assignment.position + '」'
      });
    }
  }

  for (let r = 0; r < requirements.length; r++) {
    const requirement = requirements[r];
    const startMinutes = shiftTimeToMinutes(requirement.startTime);
    const endMinutes = shiftTimeToMinutes(requirement.endTime);

    if (startMinutes === null || endMinutes === null) {
      warnings.push({
        'time': '時刻不明',
        'title': requirement.position + '：必要人数条件の時刻を読み取れません',
        'reason': '開始「' + requirement.startTime + '」／終了「' + requirement.endTime + '」'
      });
    } else if (startMinutes >= endMinutes) {
      warnings.push({
        'time': formatShiftTimeRange(startMinutes, endMinutes),
        'title': requirement.position + '：開始時刻が終了時刻以降です',
        'reason': requirement.startTime + '〜' + requirement.endTime
      });
    }

    if (!(requirement.requiredCount >= 1) || Math.floor(requirement.requiredCount) !== requirement.requiredCount) {
      warnings.push({
        'time': '要修正',
        'title': requirement.position + '：必要人数が不正です',
        'reason': '必要人数「' + requirement.requiredCount + '」'
      });
    }
    if (POSITION_ITEMS.indexOf(requirement.position) === -1) {
      warnings.push({
        'time': '要修正',
        'title': '必要人数条件のポジションが不正です',
        'reason': 'ポジション「' + requirement.position + '」'
      });
    }
  }
}

/**
 * 確認事項の文言に使う従業員の呼び名を作る。名簿にない従業員IDでも表示できるようにする
 */
function getShiftEmployeeLabel(employeeId, nameByEmpId) {
  return nameByEmpId[String(employeeId)] || ('従業員番号 ' + employeeId);
}

/**
 * 1日分の状態を決める
 * 割当も必要人数条件も無ければ未作成、確認事項があれば要確認、それ以外は作成済み・問題なし
 */
function decideShiftStatus(assignments, requirements, warnings) {
  if (assignments.length === 0 && requirements.length === 0) {
    return 'uncreated';
  }
  return warnings.length > 0 ? 'warning' : 'ok';
}

/**
 * 指定した年月の、日付ごとの状態と確認事項を返す（カレンダー表示用）
 * year・month を渡さない場合は今日を含む月を対象にする
 */
function getShiftMonth(year, month, authToken) {
  requireManagerRole_(authToken);

  let targetYear = Number(year);
  let targetMonth = Number(month);
  if (!targetYear || !targetMonth) {
    const now = new Date();
    targetYear = Number(Utilities.formatDate(now, TIME_ZONE, 'yyyy'));
    targetMonth = Number(Utilities.formatDate(now, TIME_ZONE, 'MM'));
  }
  if (targetYear < 1900 || targetYear > 2999 || targetMonth < 1 || targetMonth > 12) {
    throw new Error('表示する年月が不正です: ' + year + '年' + month + '月');
  }

  const assignmentsByDate = readShiftAssignmentsByDate();
  const requirementsByDate = readShiftRequirementsByDate();
  const nameByEmpId = getEmployeeNameMap();

  // 「2026/09/」で始まる日付キーだけを対象月のデータとして返す
  const monthPrefix = targetYear + '/' + ('0' + targetMonth).slice(-2) + '/';
  const dateKeys = collectShiftDateKeys(assignmentsByDate, requirementsByDate);
  const days = {};
  for (let i = 0; i < dateKeys.length; i++) {
    const dateKey = dateKeys[i];
    if (dateKey.indexOf(monthPrefix) !== 0) {
      continue;
    }
    const assignments = assignmentsByDate[dateKey] || [];
    const requirements = requirementsByDate[dateKey] || [];
    const warnings = buildShiftWarnings(assignments, requirements, nameByEmpId);
    days[dateKey] = {
      'status': decideShiftStatus(assignments, requirements, warnings),
      'warnings': warnings
    };
  }

  return {
    'year': targetYear,
    'month': targetMonth,
    'today': getTodayDateKey(), // 今日の判定はブラウザの時計ではなくサーバー側の日付に揃える
    'days': days
  };
}

/**
 * 指定日のシフト割当・必要人数条件・従業員一覧・確認事項を返す（日付ごとの編集欄用）
 */
function getShiftDay(dateKey, authToken) {
  requireManagerRole_(authToken);

  const normalizedDateKey = toShiftDateKey(dateKey);
  if (parseShiftDateKey(normalizedDateKey) === null) {
    throw new Error('日付の形式が不正です: ' + dateKey);
  }

  const assignments = readShiftAssignmentsByDate()[normalizedDateKey] || [];
  const requirements = readShiftRequirementsByDate()[normalizedDateKey] || [];
  const nameByEmpId = getEmployeeNameMap();

  return {
    'date': normalizedDateKey,
    'employees': getEmployees(),
    'assignments': assignments,
    'requirements': requirements,
    'warnings': buildShiftWarnings(assignments, requirements, nameByEmpId)
  };
}

/**
 * 画面から送られてきた値を文字列に揃える（未入力は空文字にする）
 */
function toShiftText(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

/**
 * 画面から送られてきた入力値を検証し、保存できる形に整える
 * 保存できない入力（時刻の前後関係・未選択・必要人数の不正など）はここでエラーにする
 * 人員不足や勤務時間の重複は保存できる確認事項なので、ここではエラーにしない
 */
function validateShiftInput(dateKey, assignments, requirements) {
  const normalizedDateKey = toShiftText(dateKey);
  const dateValue = parseShiftDateKey(normalizedDateKey);
  if (dateValue === null) {
    throw new Error('日付の形式が不正です: ' + dateKey);
  }

  // 画面から送られてきた従業員IDを、従業員名簿に入っている値そのもの（数値なら数値）へ戻すための対応表
  const employees = getEmployees();
  const rawEmpIdByKey = {};
  for (let i = 0; i < employees.length; i++) {
    rawEmpIdByKey[String(employees[i]['id'])] = employees[i]['id'];
  }

  const assignmentList = assignments || [];
  const normalizedAssignments = [];
  for (let i = 0; i < assignmentList.length; i++) {
    const label = 'シフト割当' + (i + 1) + '行目：';
    const assignment = assignmentList[i] || {};

    const employeeKey = toShiftText(assignment['employeeId']);
    if (employeeKey === '') {
      throw new Error(label + '従業員を選択してください');
    }
    if (!Object.prototype.hasOwnProperty.call(rawEmpIdByKey, employeeKey)) {
      throw new Error(label + '従業員名簿に存在しない従業員番号です: ' + employeeKey);
    }

    const position = toShiftText(assignment['position']);
    if (position === '') {
      throw new Error(label + 'ポジションを選択してください');
    }
    if (POSITION_ITEMS.indexOf(position) === -1) {
      throw new Error(label + '不正なポジションです: ' + position);
    }

    const startTime = toShiftTimeText(assignment['startTime']);
    const endTime = toShiftTimeText(assignment['endTime']);
    const startMinutes = shiftTimeToMinutes(startTime);
    const endMinutes = shiftTimeToMinutes(endTime);
    if (startMinutes === null) {
      throw new Error(label + '出勤時刻を入力してください');
    }
    if (endMinutes === null) {
      throw new Error(label + '退勤時刻を入力してください');
    }
    if (startMinutes >= endMinutes) {
      throw new Error(label + '出勤時刻は退勤時刻より前にしてください');
    }

    normalizedAssignments.push({
      'employeeId': employeeKey,
      'rawEmployeeId': rawEmpIdByKey[employeeKey],
      'startTime': startTime,
      'endTime': endTime,
      'position': position,
      'note': toShiftText(assignment['note'])
    });
  }

  const requirementList = requirements || [];
  const normalizedRequirements = [];
  for (let r = 0; r < requirementList.length; r++) {
    const label = '必要人数条件' + (r + 1) + '行目：';
    const requirement = requirementList[r] || {};

    const position = toShiftText(requirement['position']);
    if (position === '') {
      throw new Error(label + 'ポジションを選択してください');
    }
    if (POSITION_ITEMS.indexOf(position) === -1) {
      throw new Error(label + '不正なポジションです: ' + position);
    }

    const startTime = toShiftTimeText(requirement['startTime']);
    const endTime = toShiftTimeText(requirement['endTime']);
    const startMinutes = shiftTimeToMinutes(startTime);
    const endMinutes = shiftTimeToMinutes(endTime);
    if (startMinutes === null) {
      throw new Error(label + '開始時刻を入力してください');
    }
    if (endMinutes === null) {
      throw new Error(label + '終了時刻を入力してください');
    }
    if (startMinutes >= endMinutes) {
      throw new Error(label + '開始時刻は終了時刻より前にしてください');
    }

    const requiredCount = Number(requirement['requiredCount']);
    if (!isFinite(requiredCount) || Math.floor(requiredCount) !== requiredCount || requiredCount < 1) {
      throw new Error(label + '必要人数は1以上の整数で入力してください');
    }

    normalizedRequirements.push({
      'startTime': startTime,
      'endTime': endTime,
      'position': position,
      'requiredCount': requiredCount
    });
  }

  return {
    'dateKey': normalizedDateKey,
    'dateValue': dateValue,
    'assignments': normalizedAssignments,
    'requirements': normalizedRequirements
  };
}

/**
 * 指定した日付キーに一致する行だけをシートから削除する
 * 行を削除すると以降の行番号がずれるため、下の行から順に削除する
 */
function deleteShiftRowsForDate(sheet, dateColumn, dateKey) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return;
  }
  const values = sheet.getRange(2, dateColumn, lastRow - 1, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (toShiftDateKey(values[i][0]) === dateKey) {
      sheet.deleteRow(i + 2); // ヘッダー行の分だけシート上の行番号は配列の添字より2大きい
    }
  }
}

/**
 * シートの末尾に行をまとめて追記する
 * 時刻の2列がシート側で時刻値に変換されないよう、書き込む前に文字列書式にしておく
 */
function appendShiftRows(sheet, rows, firstTimeColumn) {
  if (rows.length === 0) {
    return;
  }
  const startRow = sheet.getLastRow() + 1;

  // シートの行数が足りないと範囲を取れないため、必要な分だけ行を増やしてから書き込む
  const shortageRows = startRow + rows.length - 1 - sheet.getMaxRows();
  if (shortageRows > 0) {
    sheet.insertRowsAfter(sheet.getMaxRows(), shortageRows);
  }

  sheet.getRange(startRow, firstTimeColumn, rows.length, 2).setNumberFormat('@');
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

/**
 * 指定日のシフト割当と必要人数条件をまとめて保存する
 * 対象日の行だけを入れ替えるので、別の日付のデータは変更しない
 */
function saveShiftDay(dateKey, assignments, requirements, authToken) {
  requireManagerRole_(authToken);

  const validated = validateShiftInput(dateKey, assignments, requirements);
  const shiftSheet = ensureShiftSheet();
  const requirementSheet = ensureShiftRequirementSheet();
  const updatedAt = new Date();

  const assignmentRows = [];
  for (let i = 0; i < validated.assignments.length; i++) {
    const assignment = validated.assignments[i];
    assignmentRows.push([
      validated.dateValue,
      assignment.rawEmployeeId,
      assignment.startTime,
      assignment.endTime,
      assignment.position,
      assignment.note,
      updatedAt
    ]);
  }
  deleteShiftRowsForDate(shiftSheet, SHIFT_COL_DATE, validated.dateKey);
  appendShiftRows(shiftSheet, assignmentRows, SHIFT_COL_START_TIME);

  const requirementRows = [];
  for (let r = 0; r < validated.requirements.length; r++) {
    const requirement = validated.requirements[r];
    requirementRows.push([
      validated.dateValue,
      requirement.startTime,
      requirement.endTime,
      requirement.position,
      requirement.requiredCount,
      updatedAt
    ]);
  }
  deleteShiftRowsForDate(requirementSheet, SHIFT_REQ_COL_DATE, validated.dateKey);
  appendShiftRows(requirementSheet, requirementRows, SHIFT_REQ_COL_START_TIME);

  const warnings = buildShiftWarnings(validated.assignments, validated.requirements, getEmployeeNameMap());
  return {
    'date': validated.dateKey,
    'status': decideShiftStatus(validated.assignments, validated.requirements, warnings),
    'warnings': warnings,
    'message': validated.dateKey + ' のシフトを保存しました'
  };
}