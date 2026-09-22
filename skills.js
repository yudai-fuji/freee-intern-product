// 従業員のスキル評価・更新日を管理する

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

/**
 * シートから読んだ旧評価もF/B/A/Sへ読み替える。
 * 表にない値はFとして扱い、Sは自動生成しない。
 * @param {*} rawValue スキル管理シートの評価列に入っている値
 * @return {string} F/B/A/Sの評価文字列
 */
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

/**
 * スキルの対応表（{ 接客: 'A', ... }）を、SKILL_ITEMS の並び順の配列にする
 * @param {Object} skills スキル名をキーにした評価の対応表
 * @return {Array} SKILL_ITEMS と同じ並びの評価の配列
 */
function toSkillValueList_(skills) {
  const skillValues = [];
  for (let i = 0; i < SKILL_ITEMS.length; i++) {
    skillValues.push(skills[SKILL_ITEMS[i]]);
  }
  return skillValues;
}

function getSkillEvalColumn(skillIndex) {
  return 3 + skillIndex * 2; // 接客=C(3), レジ=E(5), バー=G(7), 厨房=I(9), 肉の仕込み=K(11), 締め作業=M(13)
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
      const rawValue = values[i][getSkillEvalColumn(s) - 1]; // 配列の添字は列番号-1
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
      const rawDate = values[i][getSkillEvalColumn(s)]; // 評価列の配列添字+1 = 更新日列
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
      'number': employees[i]['number'], // 一覧・詳細の両方に従業員番号を出すため
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
      return i + 2; // ヘッダー行の分だけシート上の行番号は配列の添字より2大きい
    }
  }
  return -1;
}

function saveEmployeeSkills(empId, skillValues, authToken) {
  requireManagerRole_(authToken);
  return withAppLock_(function () { return saveEmployeeSkills_(empId, skillValues); });
}

function saveEmployeeSkills_(empId, skillValues) {
  // 有効な従業員だけを編集対象にする（削除済みの人のスキル記録は残すが、書き換えはしない）
  const nameByEmpId = getActiveEmployeeNameMap();
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

  const sheet = getSkillSheet();
  let targetRow = findSkillRowByEmpId(sheet, empId);
  const isNewRow = (targetRow === -1);
  if (isNewRow) {
    targetRow = sheet.getLastRow() + 1;
    sheet.getRange(targetRow, 1).setValue(empId);
  }
  sheet.getRange(targetRow, 2).setValue(empName);

  // 既存行の場合のみ、現在の評価・更新日を読んで変更の有無を判定する
  // 新規行の場合は、画面の初期値と同じFを現在の評価とみなして判定する
  const currentRowValues = isNewRow ? null : sheet.getRange(targetRow, 3, 1, SKILL_ITEMS.length * 2).getValues()[0];
  const today = new Date();
  const updatedAt = {};

  for (let s = 0; s < SKILL_ITEMS.length; s++) {
    const evalCol = getSkillEvalColumn(s);
    const newValue = skillValues[s];
    const currentRaw = isNewRow ? SKILL_LEVEL_DEFAULT : currentRowValues[evalCol - 3]; // 評価列の配列添字（列3=添字0基準）
    const currentDate = isNewRow ? null : currentRowValues[evalCol - 3 + 1];
    // 旧表記の○・習得済とAは同じ評価として比較し、表記変更だけでは更新日を動かさない。
    const isChanged = (normalizeSkillLevel_(currentRaw) !== newValue);

    // 評価自体は、新規行のときと、シートの文字列が新しい書き方と違うときに書き込む
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

  // 保存したスキルをもとにスキル手当を見直す（上がることはあっても、下がることはない）
  promoteSkillAllowance_(empId, skillValues);

  return {
    'message': empName + 'のスキルを保存しました',
    'updatedAt': updatedAt
  };
}

/**
 * 習得済みとして扱うA・Sの数を数える（B・Fは含めない）
 * @param {Array} skillValues SKILL_ITEMS と同じ並びの評価の配列
 * @return {number} 習得済の数（0〜6）
 */
function countMasteredSkills_(skillValues) {
  let masteredCount = 0;
  for (let i = 0; i < skillValues.length; i++) {
    const level = normalizeSkillLevel_(skillValues[i]);
    if (level === SKILL_LEVEL_MASTERED || level === SKILL_LEVEL_EXPERT) {
      masteredCount++;
    }
  }
  return masteredCount;
}

/**
 * 習得済の数から、今回到達したスキル手当（円/時）を求める
 * 4個で50円、6個で100円。到達していない段階は金額に含めない
 * @param {number} masteredCount 習得済の数
 * @return {number} 今回到達したスキル手当
 */
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

/**
 * 従業員1人のスキル手当を見直す
 * 新しい手当額 = max(名簿に保存されている手当額, 今回到達した手当額) で決めるので、
 * あとから評価を下げても、一度到達した手当が下がることはない
 * @param {*} empId 従業員番号
 * @param {Array} skillValues SKILL_ITEMS と同じ並びの評価の配列
 * @return {number} 見直したあとのスキル手当（円/時）
 */
function promoteSkillAllowance_(empId, skillValues) {
  const savedAmount = getEmployeeSkillAllowance(empId);
  const earnedAmount = calcEarnedSkillAllowance_(countMasteredSkills_(skillValues));
  const nextAmount = Math.max(savedAmount, earnedAmount);
  // 金額が変わらないときも書き込む。列を足した直後の空欄に 0 が入り、
  // 名簿のスキル手当列がいつも 0・50・100 のどれかになる
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
    // 既に評価・更新日の列数が揃っている＝新レイアウト済みとみなし、ヘッダーだけ書き直す
    sheet.getRange(1, 1, 1, SKILL_SHEET_HEADER.length).setValues([SKILL_SHEET_HEADER]);
    return 'ヘッダーを更新しました（データは新レイアウト済みのため変更していません）';
  }

  // 旧レイアウト（A:従業員ID B:氏名 C〜H:評価のみ）とみなして変換する
  const oldValues = sheet.getRange(2, 1, lastRow - 1, 2 + SKILL_ITEMS.length).getValues();
  const newValues = [];
  for (let i = 0; i < oldValues.length; i++) {
    const row = oldValues[i];
    const newRow = [row[0], row[1]]; // 従業員ID, 氏名
    for (let s = 0; s < SKILL_ITEMS.length; s++) {
      newRow.push(row[2 + s]); // 評価はそのまま引き継ぐ
      newRow.push(''); // 更新日は不明なので空欄
    }
    newValues.push(newRow);
  }

  sheet.getRange(1, 1, 1, SKILL_SHEET_HEADER.length).setValues([SKILL_SHEET_HEADER]);
  sheet.getRange(2, 1, newValues.length, newColumnCount).setValues(newValues);
  return newValues.length + '件のデータを新しい列構成（評価・更新日を別列）に変換しました';
}

// ここから下は一括変換・再計算のための関数
// 初回起動の移行から呼ぶ。Apps Scriptエディタから再実行することもできる。
// どちらも何度実行しても結果が変わらない（変換済みの文字列や、すでに到達した手当は壊さない）

/**
 * スキル管理シートの旧評価をF/B/Aへ一括変換する（Sは自動生成しない）。
 * 未習得・× → F ／ 習得中・△ → B ／ 習得済・○・◎ → A
 *
 * ・すでにF/B/A/Sになっている値と空欄はそのまま残す
 * ・対応表にない値は書き換えず、人が確認できるように残す
 * ・書き込むのは評価列だけなので、更新日列の日付は一切変わらない
 *
 * 変換のあと、その内容にあわせてスキル手当も計算し直す（recalcSkillAllowances）
 * @return {string} 実行結果の説明（Apps Script のログに出る）
 */
function migrateSkillRatingsToText() {
  const sheet = getSkillSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return 'スキル管理シートに変換対象のデータがありません';
  }

  let convertedCount = 0;
  let unknownCount = 0;

  // 評価列だけを1列ずつ読み書きする。更新日列には触れないので、記録済みの日付はそのまま残る
  for (let s = 0; s < SKILL_ITEMS.length; s++) {
    const evalRange = sheet.getRange(2, getSkillEvalColumn(s), lastRow - 1, 1);
    const evalValues = evalRange.getValues();
    let isColumnChanged = false;

    for (let r = 0; r < evalValues.length; r++) {
      const rawValue = evalValues[r][0];
      const text = String(rawValue === null || rawValue === undefined ? '' : rawValue).trim();
      if (text === '' || SKILL_LEVELS.indexOf(text) !== -1) {
        continue; // 空欄と変換済みの評価は触らない
      }
      if (!Object.prototype.hasOwnProperty.call(SKILL_LEGACY_LEVEL_MAP, text)) {
        unknownCount++; // 想定外の値は勝手に書き換えない
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

/**
 * スキル管理シートの現在の評価をもとに、従業員名簿のスキル手当を計算し直す
 * 一度到達した手当は下げないので、書き込むのは max(保存済みの手当, 今回到達した手当) で上がるときだけ
 * 記号の一括変換より前に作られた評価にも、あとからスキル手当を付けるために使う
 * @return {string} 実行結果の説明（Apps Script のログに出る）
 */
function recalcSkillAllowances() {
  const columns = ensureEmployeeSkillAllowanceColumn_();
  const sheet = columns.sheet;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return 'スキル手当を更新する従業員がいません';
  }

  const skillMapByEmpId = getEmployeeSkillMap(); // 旧評価もF/B/Aへ読み替えて返す
  const empIdValues = sheet.getRange(2, EMPLOYEE_COL_ID, lastRow - 1, 1).getValues();
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
      allowanceValues[i][0] = nextAmount; // 空欄の従業員には 0 を入れて、手当なしを明示する
      isChanged = true;
    }
  }

  if (isChanged) {
    allowanceRange.setValues(allowanceValues);
  }
  return 'スキル手当を ' + raisedCount + '人分 引き上げました';
}
