function ensureSkillSheetHeader(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, SKILL_SHEET_HEADER.length);
  const currentHeader = headerRange.getValues()[0];
  for (let i = 0; i < SKILL_SHEET_HEADER.length; i++) {
    if (currentHeader[i] !== SKILL_SHEET_HEADER[i]) {
      headerRange.setValues([SKILL_SHEET_HEADER]);
      return;
    }
  }
}

function getSkillSheet() {
  const sheet = getSheet(SHEET_NAMES.SKILLS);
  ensureSkillSheetHeader(sheet);
  return sheet;
}

function getDefaultSkills() {
  const skills = {};
  for (let i = 0; i < SKILL_ITEMS.length; i++) {
    skills[SKILL_ITEMS[i]] = SKILL_LEVEL_DEFAULT;
  }
  return skills;
}

function normalizeSkillLevel_(rawValue) {
  const text = String(rawValue === null || rawValue === undefined ? '' : rawValue).trim();
  if (SKILL_LEVELS.indexOf(text) !== -1) {
    return text;
  }
  if (Object.prototype.hasOwnProperty.call(SKILL_LEGACY_LEVEL_MAP, text)) {
    return SKILL_LEGACY_LEVEL_MAP[text];
  }
  return SKILL_LEVEL_DEFAULT;
}

function toSkillValueList_(skills) {
  const skillValues = [];
  for (let i = 0; i < SKILL_ITEMS.length; i++) {
    skillValues.push(skills[SKILL_ITEMS[i]]);
  }
  return skillValues;
}

function getSkillEvalColumn(skillIndex) {
  return 3 + skillIndex * 2;
}

function getEmployeeSkillMap() {
  const sheet = getSkillSheet();
  const lastRow = sheet.getLastRow();
  const skillMapByEmpId = {};
  if (lastRow < 2) {
    return skillMapByEmpId;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 2 + SKILL_ITEMS.length * 2).getValues();
  for (let i = 0; i < values.length; i++) {
    const empId = values[i][0];
    if (empId === '') {
      continue;
    }
    const skills = {};
    for (let s = 0; s < SKILL_ITEMS.length; s++) {
      const rawValue = values[i][getSkillEvalColumn(s) - 1];
      skills[SKILL_ITEMS[s]] = normalizeSkillLevel_(rawValue);
    }
    skillMapByEmpId[String(empId)] = skills;
  }
  return skillMapByEmpId;
}

function getEmployeeSkillUpdatedAtMap() {
  const sheet = getSkillSheet();
  const lastRow = sheet.getLastRow();
  const updatedAtMapByEmpId = {};
  if (lastRow < 2) {
    return updatedAtMapByEmpId;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 2 + SKILL_ITEMS.length * 2).getValues();
  for (let i = 0; i < values.length; i++) {
    const empId = values[i][0];
    if (empId === '') {
      continue;
    }
    const updatedAt = {};
    for (let s = 0; s < SKILL_ITEMS.length; s++) {
      const rawDate = values[i][getSkillEvalColumn(s)];
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
  const updatedAt = {};
  for (let i = 0; i < SKILL_ITEMS.length; i++) {
    updatedAt[SKILL_ITEMS[i]] = '';
  }
  return updatedAt;
}

function getSkillMatrix(authToken) {
  requireManagerRole_(authToken);

  const employees = getEmployees();
  const skillMapByEmpId = getEmployeeSkillMap();
  const updatedAtMapByEmpId = getEmployeeSkillUpdatedAtMap();

  const matrix = [];
  for (let i = 0; i < employees.length; i++) {
    const empId = employees[i]['id'];
    const key = String(empId);
    matrix.push({
      'id': empId,
      'name': employees[i]['name'],
      'number': employees[i]['number'],
      'skills': skillMapByEmpId[key] || getDefaultSkills(),
      'updatedAt': updatedAtMapByEmpId[key] || getDefaultSkillUpdatedAt()
    });
  }
  return matrix;
}

function findSkillRowByEmpId(sheet, empId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return -1;
  }
  const empIdValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < empIdValues.length; i++) {
    if (String(empIdValues[i][0]) === String(empId)) {
      return i + 2;
    }
  }
  return -1;
}

function saveEmployeeSkills(empId, skillValues, authToken) {
  requireManagerRole_(authToken);
  return withAppLock_(function () { return saveEmployeeSkills_(empId, skillValues); });
}

function validateEmployeeSkills_(empId, skillValues, nameByEmpId) {

  const empName = nameByEmpId[String(empId)];
  if (!empName) {
    throw new Error('この従業員は選べません。画面を開き直してください（従業員番号 ' + empId + '）');
  }
  if (!Array.isArray(skillValues) || skillValues.length !== SKILL_ITEMS.length) {
    throw new Error('スキルの項目数が不正です');
  }
  for (let i = 0; i < skillValues.length; i++) {
    if (SKILL_LEVELS.indexOf(skillValues[i]) === -1) {
      throw new Error('不正なスキル評価値です: ' + skillValues[i]);
    }
  }

  return empName;
}

function saveEmployeeSkillsBatch(changes, authToken) {
  requireManagerRole_(authToken);
  return withAppLock_(function () {
    if (!Array.isArray(changes) || changes.length === 0) { throw new Error('保存する変更がありません'); }
    const names = getActiveEmployeeNameMap();
    const seen = new Set();
    changes.forEach(function (change) {
      if (!change || seen.has(String(change.empId))) { throw new Error('従業員の指定が不正です'); }
      validateEmployeeSkills_(change.empId, change.skillValues, names);
      seen.add(String(change.empId));
    });

    const sheet = getSkillSheet();
    const allowance = ensureEmployeeSkillAllowanceColumn_();
    const employees = readEmployeeRows_();
    let nextRow = sheet.getLastRow() + 1;
    const backups = [];
    function remember(range) {
      const values = range.getValues();
      const formulas = range.getFormulas();
      backups.push({ range: range, values: values.map(function (row, r) {
        return row.map(function (value, c) { return formulas[r][c] || value; });
      }) });
    }
    changes.forEach(function (change) {
      const existingRow = findSkillRowByEmpId(sheet, change.empId);
      const targetRow = existingRow === -1 ? nextRow++ : existingRow;
      if (targetRow > sheet.getMaxRows()) { sheet.insertRowsAfter(sheet.getMaxRows(), targetRow - sheet.getMaxRows()); }
      remember(sheet.getRange(targetRow, 1, 1, SKILL_SHEET_HEADER.length));
      const employee = findEmployeeRow_(employees, change.empId);
      if (!employee) { throw new Error('従業員が見つかりません'); }
      remember(allowance.sheet.getRange(employee.rowNumber, allowance.column));
    });

    try {
      const saved = changes.map(function (change) {
        const result = saveEmployeeSkills_(change.empId, change.skillValues);
        return { empId: change.empId, skillValues: change.skillValues, updatedAt: result.updatedAt };
      });
      SpreadsheetApp.flush();
      return { employees: saved, message: saved.length + '人分の変更を保存しました' };
    } catch (error) {

      let restoreFailed = false;
      backups.forEach(function (backup) {
        try { backup.range.setValues(backup.values); } catch (restoreError) { restoreFailed = true; }
      });
      try { SpreadsheetApp.flush(); } catch (restoreError) { restoreFailed = true; }
      throw new Error(error.message + (restoreFailed ? '（復元できなかった可能性があります。保存済みデータを確認してください）' : ''));
    }
  });
}

function saveEmployeeSkills_(empId, skillValues) {
  const empName = validateEmployeeSkills_(empId, skillValues, getActiveEmployeeNameMap());
  const sheet = getSkillSheet();
  let targetRow = findSkillRowByEmpId(sheet, empId);
  const isNewRow = (targetRow === -1);
  if (isNewRow) {
    targetRow = sheet.getLastRow() + 1;
    sheet.getRange(targetRow, 1).setValue(empId);
  }
  sheet.getRange(targetRow, 2).setValue(empName);

  const currentRowValues = isNewRow ? null : sheet.getRange(targetRow, 3, 1, SKILL_ITEMS.length * 2).getValues()[0];
  const today = new Date();
  const updatedAt = {};

  for (let s = 0; s < SKILL_ITEMS.length; s++) {
    const evalCol = getSkillEvalColumn(s);
    const newValue = skillValues[s];
    const currentRaw = isNewRow ? SKILL_LEVEL_DEFAULT : currentRowValues[evalCol - 3];
    const currentDate = isNewRow ? null : currentRowValues[evalCol - 3 + 1];

    const isChanged = (normalizeSkillLevel_(currentRaw) !== newValue);

    if (isNewRow || currentRaw !== newValue) {
      sheet.getRange(targetRow, evalCol).setValue(newValue);
    }

    if (isChanged) {
      sheet.getRange(targetRow, evalCol + 1).setValue(today);
      updatedAt[SKILL_ITEMS[s]] = formatSkillUpdatedAt(today);
    } else {
      updatedAt[SKILL_ITEMS[s]] = formatSkillUpdatedAt(currentDate);
    }
  }

  promoteSkillAllowance_(empId, skillValues);

  return {
    'message': empName + 'のスキルを保存しました',
    'updatedAt': updatedAt
  };
}

function countMasteredSkills_(skillValues) {
  let masteredCount = 0;
  for (let i = 0; i < skillValues.length; i++) {
    const level = normalizeSkillLevel_(skillValues[i]);
    if (level === SKILL_LEVEL_MASTERED) {
      masteredCount++;
    }
  }
  return masteredCount;
}

function calcEarnedSkillAllowance_(masteredCount) {
  let earnedAmount = SKILL_ALLOWANCE_DEFAULT;
  for (let i = 0; i < SKILL_ALLOWANCE_TIERS.length; i++) {
    const tier = SKILL_ALLOWANCE_TIERS[i];
    if (masteredCount >= tier.masteredCount && tier.amount > earnedAmount) {
      earnedAmount = tier.amount;
    }
  }
  return earnedAmount;
}

function promoteSkillAllowance_(empId, skillValues) {
  const savedAmount = getEmployeeSkillAllowance(empId);
  const earnedAmount = calcEarnedSkillAllowance_(countMasteredSkills_(skillValues));
  const nextAmount = Math.max(savedAmount, earnedAmount);

  saveEmployeeSkillAllowance_(empId, nextAmount);
  return nextAmount;
}

function migrateSkillSheetToDatedColumns() {
  const sheet = getSheet(SHEET_NAMES.SKILLS);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) {
    sheet.getRange(1, 1, 1, SKILL_SHEET_HEADER.length).setValues([SKILL_SHEET_HEADER]);
    return 'ヘッダーのみ更新しました（既存データはありません）';
  }

  const newColumnCount = 2 + SKILL_ITEMS.length * 2;
  if (lastCol >= newColumnCount) {

    sheet.getRange(1, 1, 1, SKILL_SHEET_HEADER.length).setValues([SKILL_SHEET_HEADER]);
    return 'ヘッダーを更新しました（データは新レイアウト済みのため変更していません）';
  }

  const oldValues = sheet.getRange(2, 1, lastRow - 1, 2 + SKILL_ITEMS.length).getValues();
  const newValues = [];
  for (let i = 0; i < oldValues.length; i++) {
    const row = oldValues[i];
    const newRow = [row[0], row[1]];
    for (let s = 0; s < SKILL_ITEMS.length; s++) {
      newRow.push(row[2 + s]);
      newRow.push('');
    }
    newValues.push(newRow);
  }

  sheet.getRange(1, 1, 1, SKILL_SHEET_HEADER.length).setValues([SKILL_SHEET_HEADER]);
  sheet.getRange(2, 1, newValues.length, newColumnCount).setValues(newValues);
  return newValues.length + '件のデータを新しい列構成（評価・更新日を別列）に変換しました';
}

function migrateSkillRatingsToText() {
  const sheet = getSkillSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return 'スキル管理シートに変換対象のデータがありません';
  }

  let convertedCount = 0;
  let unknownCount = 0;

  for (let s = 0; s < SKILL_ITEMS.length; s++) {
    const evalRange = sheet.getRange(2, getSkillEvalColumn(s), lastRow - 1, 1);
    const evalValues = evalRange.getValues();
    let isColumnChanged = false;

    for (let r = 0; r < evalValues.length; r++) {
      const rawValue = evalValues[r][0];
      const text = String(rawValue === null || rawValue === undefined ? '' : rawValue).trim();
      if (text === '' || SKILL_LEVELS.indexOf(text) !== -1) {
        continue;
      }
      if (!Object.prototype.hasOwnProperty.call(SKILL_LEGACY_LEVEL_MAP, text)) {
        unknownCount++;
        continue;
      }
      evalValues[r][0] = SKILL_LEGACY_LEVEL_MAP[text];
      convertedCount++;
      isColumnChanged = true;
    }

    if (isColumnChanged) {
      evalRange.setValues(evalValues);
    }
  }

  const allowanceMessage = recalcSkillAllowances();
  const unknownMessage = (unknownCount === 0) ? '' :
    '（判定できない値 ' + unknownCount + '件はそのまま残しました）';
  return '旧スキル評価を ' + convertedCount + '件 F/B/Aへ変換しました' + unknownMessage + ' / ' + allowanceMessage;
}

function recalcSkillAllowances() {
  const columns = ensureEmployeeSkillAllowanceColumn_();
  const sheet = columns.sheet;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return 'スキル手当を更新する従業員がいません';
  }

  const skillMapByEmpId = getEmployeeSkillMap();
  const empIdValues = sheet.getRange(2, getEmployeeIdentityColumns_(sheet).id, lastRow - 1, 1).getValues();
  const allowanceRange = sheet.getRange(2, columns.column, lastRow - 1, 1);
  const allowanceValues = allowanceRange.getValues();

  let raisedCount = 0;
  let isChanged = false;
  for (let i = 0; i < empIdValues.length; i++) {
    const empId = empIdValues[i][0];
    if (empId === '') {
      continue;
    }
    const savedAmount = normalizeSkillAllowance_(allowanceValues[i][0]);
    const skills = skillMapByEmpId[String(empId)] || getDefaultSkills();
    const earnedAmount = calcEarnedSkillAllowance_(countMasteredSkills_(toSkillValueList_(skills)));
    const nextAmount = Math.max(savedAmount, earnedAmount);

    if (nextAmount > savedAmount) {
      raisedCount++;
    }
    if (allowanceValues[i][0] !== nextAmount) {
      allowanceValues[i][0] = nextAmount;
      isChanged = true;
    }
  }

  if (isChanged) {
    allowanceRange.setValues(allowanceValues);
  }
  return 'スキル手当を ' + raisedCount + '人分 引き上げました';
}
