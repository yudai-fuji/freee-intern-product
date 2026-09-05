// 従業員のスキル評価・更新日を管理する

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

function getSkillSheet() {
  var sheet = getSheet(SHEET_NAMES.SKILLS);
  ensureSkillSheetHeader(sheet);
  return sheet;
}

function getDefaultSkills() {
  var skills = {};
  var defaultLevel = SKILL_LEVELS[SKILL_LEVELS.length - 1]; // ×
  for (var i = 0; i < SKILL_ITEMS.length; i++) {
    skills[SKILL_ITEMS[i]] = defaultLevel;
  }
  return skills;
}

function getSkillEvalColumn(skillIndex) {
  return 3 + skillIndex * 2; // 接客=C(3), レジ=E(5), バー=G(7), 厨房=I(9), 肉の仕込み=K(11), 締め作業=M(13)
}

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

function formatSkillUpdatedAt(rawDate) {
  if (rawDate instanceof Date) {
    return Utilities.formatDate(rawDate, TIME_ZONE, 'yyyy/MM/dd');
  }
  return '';
}

function getDefaultSkillUpdatedAt() {
  var updatedAt = {};
  for (var i = 0; i < SKILL_ITEMS.length; i++) {
    updatedAt[SKILL_ITEMS[i]] = '';
  }
  return updatedAt;
}

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