// 従業員の取得・追加・削除・氏名管理を行う
// 削除は名簿の行や氏名を消さず「状態」列を『削除済み』にするだけなので、
// 過去の勤怠・スキルの記録と氏名はそのまま残り、従業員番号も別の人に使い回されない

/**
 * 従業員名簿に「状態」「削除日時」の列を用意する
 * 既存の列・データは書き換えず、見出しが見つからないときだけ名簿の右端に足す
 * @return {Object} { sheet: 従業員名簿シート, statusColumn: 状態列の列番号, deletedAtColumn: 削除日時列の列番号 }
 */
function ensureEmployeeStatusColumns_() {
  const sheet = getSheet(SHEET_NAMES.EMPLOYEES);

  // 見出し行を読み、状態・削除日時の列がすでにあるかを調べる
  const lastColumn = Math.max(sheet.getLastColumn(), EMPLOYEE_COL_PART_TIME_MARK);
  const header = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  let statusColumn = 0;
  let deletedAtColumn = 0;
  for (let i = 0; i < header.length; i++) {
    const title = String(header[i] === null || header[i] === undefined ? '' : header[i]).trim();
    if (title === EMPLOYEE_STATUS_HEADER) {
      statusColumn = i + 1;
    } else if (title === EMPLOYEE_DELETED_AT_HEADER) {
      deletedAtColumn = i + 1;
    }
  }

  // 見つからなかった列だけを、既存の列より右側に新しく割り当てる
  let nextColumn = lastColumn;
  if (statusColumn === 0) {
    nextColumn++;
    statusColumn = nextColumn;
  }
  if (deletedAtColumn === 0) {
    nextColumn++;
    deletedAtColumn = nextColumn;
  }

  // 割り当てた列がシートの列数を超える場合は、足りない分だけ列を増やしてから見出しを書く
  if (nextColumn > sheet.getMaxColumns()) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), nextColumn - sheet.getMaxColumns());
  }
  if (String(sheet.getRange(1, statusColumn).getValue()).trim() !== EMPLOYEE_STATUS_HEADER) {
    sheet.getRange(1, statusColumn).setValue(EMPLOYEE_STATUS_HEADER);
  }
  if (String(sheet.getRange(1, deletedAtColumn).getValue()).trim() !== EMPLOYEE_DELETED_AT_HEADER) {
    sheet.getRange(1, deletedAtColumn).setValue(EMPLOYEE_DELETED_AT_HEADER);
  }

  return { 'sheet': sheet, 'statusColumn': statusColumn, 'deletedAtColumn': deletedAtColumn };
}

/**
 * 従業員名簿の全行を、削除済みかどうかも含めて読み取る
 * 一覧・氏名の対応表・追加・削除は、すべてこの関数が読んだ内容をもとに判断する
 * @return {Array} [{ id: 従業員番号, name: 氏名, rowNumber: シート上の行番号, isDeleted: 削除済みなら true }, ...]
 */
function readEmployeeRows_() {
  const columns = ensureEmployeeStatusColumns_();
  const sheet = columns.sheet;
  const lastRow = sheet.getLastRow();
  const rows = [];
  if (lastRow < 2) {
    return rows; // 見出し行しかない
  }

  const width = Math.max(columns.statusColumn, columns.deletedAtColumn);
  const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
  for (let i = 0; i < values.length; i++) {
    const empId = values[i][EMPLOYEE_COL_ID - 1];
    if (empId === '') {
      continue; // 従業員番号が無い行はそもそも使わない
    }
    // 状態列が空欄の従業員は、これまでどおり有効として扱う
    const status = String(values[i][columns.statusColumn - 1] || '').trim();
    rows.push({
      'id': empId,
      'name': String(values[i][EMPLOYEE_COL_NAME - 1] || '').trim(),
      'rowNumber': i + 2, // 見出し行の分だけ、シート上の行番号は配列の添字より2大きい
      'isDeleted': (status === EMPLOYEE_STATUS_DELETED)
    });
  }
  return rows;
}

/**
 * 画面に表示する従業員番号の文字列を作る
 * 名簿の値そのものは変えず、表示のときだけ「01」「001」のゼロ埋めを外して「1」にする
 * @param {*} empId 名簿に入っている従業員番号
 * @return {string} 表示用の従業員番号
 */
function formatEmployeeNumber(empId) {
  const text = String(empId === null || empId === undefined ? '' : empId).trim();
  if (/^[0-9]+$/.test(text)) {
    return String(Number(text));
  }
  return text;
}

/**
 * 有効な従業員（削除済みでなく、氏名が入っている人）の一覧を取得する
 * 従業員画面の一覧・本日の業務設定・スキル管理の対象は、すべてこの一覧に揃える
 * @return {Array} [{ id: 従業員番号, name: 氏名, number: 表示用の従業員番号 }, ...]
 */
function getEmployees() {
  const rows = readEmployeeRows_();
  const employees = [];
  for (let i = 0; i < rows.length; i++) {
    if (rows[i].isDeleted || rows[i].name === '') {
      continue; // 削除済みの人と、まだ氏名が入っていない番号は一覧に出さない
    }
    employees.push({
      'id': rows[i].id,
      'name': rows[i].name,
      'number': formatEmployeeNumber(rows[i].id)
    });
  }
  return employees;
}

/**
 * 画面用に有効な従業員の一覧を返す（ログインしている人だけが取得できる）
 * 追加・削除のあとも、この関数で取り直した一覧を表示すれば最新の内容になる
 * @param {string} authToken ログイントークン
 * @return {Array} getEmployees() と同じ形式の一覧
 */
function getActiveEmployeeList(authToken) {
  requireLoggedIn_(authToken);
  return getEmployees();
}

/**
 * 従業員番号と氏名の対応表をつくる（削除済みの人も含む）
 * 過去の勤怠に残っている記録の氏名を表示するために使う
 * @return {Object} { '3': '川村 祐大', ... }
 */
function getEmployeeNameMap() {
  const rows = readEmployeeRows_();
  const nameByEmpId = {};
  for (let i = 0; i < rows.length; i++) {
    if (rows[i].name === '') {
      continue;
    }
    nameByEmpId[String(rows[i].id)] = rows[i].name;
  }
  return nameByEmpId;
}

/**
 * 有効な従業員だけの、従業員番号と氏名の対応表をつくる
 * 打刻・本日の業務・スキルの保存先として選んでよい人かどうかの確認に使う
 * @return {Object} { '3': '川村 祐大', ... }
 */
function getActiveEmployeeNameMap() {
  const employees = getEmployees();
  const nameByEmpId = {};
  for (let i = 0; i < employees.length; i++) {
    nameByEmpId[String(employees[i].id)] = employees[i].name;
  }
  return nameByEmpId;
}

/**
 * 従業員名簿から、指定した従業員番号の行を探す
 * @param {Array} rows readEmployeeRows_() が返した一覧
 * @param {*} empId 探したい従業員番号
 * @return {Object} 見つかった行。見つからなければ null
 */
function findEmployeeRow_(rows, empId) {
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i].id) === String(empId)) {
      return rows[i];
    }
  }
  return null;
}

/**
 * 従業員を追加する（店長だけが実行できる）
 * 氏名がまだ入っていない番号の行に書き込む。削除済みの行は飛ばすので、番号が別の人に使い回されない
 * @param {string} lastName 姓
 * @param {string} firstName 名
 * @param {string} authToken ログイントークン
 * @return {Object} { message: 画面に出す文言, employees: 追加後の有効な従業員一覧 }
 */
function addEmployee(lastName, firstName, authToken) {
  requireManagerRole_(authToken);

  // .trim() で文字列に含まれる余分な空白を削除する
  lastName = String(lastName === null || lastName === undefined ? '' : lastName).trim();
  firstName = String(firstName === null || firstName === undefined ? '' : firstName).trim();
  if (lastName === '' || firstName === '') {
    throw new Error('姓と名を入力してください');
  }

  const columns = ensureEmployeeStatusColumns_();
  const rows = readEmployeeRows_();

  // 「従業員番号があり、氏名が空欄で、削除済みでない」行だけを追加先にする
  let targetRow = null;
  for (let i = 0; i < rows.length; i++) {
    if (rows[i].name === '' && !rows[i].isDeleted) {
      targetRow = rows[i];
      break;
    }
  }
  if (targetRow === null) {
    throw new Error('空いている従業員番号がありません');
  }

  const fullName = lastName + ' ' + firstName;
  const partTimeMark = '○'; // 〇のUnicode
  columns.sheet.getRange(targetRow.rowNumber, EMPLOYEE_COL_NAME).setValue(fullName);
  columns.sheet.getRange(targetRow.rowNumber, EMPLOYEE_COL_PART_TIME_MARK).setValue(partTimeMark);
  columns.sheet.getRange(targetRow.rowNumber, columns.statusColumn).setValue(EMPLOYEE_STATUS_ACTIVE);

  return {
    'message': fullName + ' を追加しました（従業員番号 : ' + formatEmployeeNumber(targetRow.id) + '）',
    'employees': getEmployees()
  };
}

/**
 * 削除する前に、対象の従業員と削除の影響を確認するための情報を返す（店長だけが実行できる）
 * @param {*} empId 削除しようとしている従業員番号
 * @param {string} authToken ログイントークン
 * @return {Object} { id, name, number }
 */
function getEmployeeDeletionPreview(empId, authToken) {
  requireManagerRole_(authToken);

  const targetRow = findEmployeeRow_(readEmployeeRows_(), empId);
  if (targetRow === null || targetRow.name === '') {
    throw new Error('従業員名簿に存在しない従業員番号です: ' + empId);
  }
  if (targetRow.isDeleted) {
    throw new Error('この従業員はすでに削除されています');
  }

  return {
    'id': targetRow.id,
    'name': targetRow.name,
    'number': formatEmployeeNumber(targetRow.id)
  };
}

/**
 * 従業員を削除済みにする（店長だけが実行できる）
 * 名簿の行と氏名はそのまま残し、状態列と削除日時列だけを書き換える
 * @param {*} empId 削除する従業員番号
 * @param {string} authToken ログイントークン
 * @return {Object} { message: 画面に出す文言, employees: 削除後の有効な従業員一覧 }
 */
function deleteEmployee(empId, authToken) {
  requireManagerRole_(authToken);

  const columns = ensureEmployeeStatusColumns_();
  const targetRow = findEmployeeRow_(readEmployeeRows_(), empId);
  if (targetRow === null || targetRow.name === '') {
    throw new Error('従業員名簿に存在しない従業員番号です: ' + empId);
  }
  if (targetRow.isDeleted) {
    throw new Error('この従業員はすでに削除されています');
  }

  columns.sheet.getRange(targetRow.rowNumber, columns.statusColumn).setValue(EMPLOYEE_STATUS_DELETED);
  columns.sheet.getRange(targetRow.rowNumber, columns.deletedAtColumn).setValue(new Date());

  return {
    'message': targetRow.name + ' を削除しました（従業員番号 : ' + formatEmployeeNumber(targetRow.id) + '）',
    'employees': getEmployees()
  };
}

// ここから下はスキル手当（円/時）を従業員名簿で持つための処理
// スキル手当は「一度到達したら下げない」ため、そのときどきの習得済の数から計算し直すのではなく、
// 到達した金額そのものを名簿に残しておく必要がある

/**
 * 従業員名簿に「スキル手当」の列を用意する
 * 状態・削除日時と同じ考え方で、既存の列・データは書き換えず、見出しが見つからないときだけ名簿の右端に足す
 * @return {Object} { sheet: 従業員名簿シート, column: スキル手当列の列番号 }
 */
function ensureEmployeeSkillAllowanceColumn_() {
  // 先に状態・削除日時の列を用意しておくことで、どの処理が先に動いても列の並びが変わらない
  const statusColumns = ensureEmployeeStatusColumns_();
  const sheet = statusColumns.sheet;

  // 見出し行を読み、スキル手当の列がすでにあるかを調べる
  const lastColumn = Math.max(sheet.getLastColumn(), statusColumns.deletedAtColumn);
  const header = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  let allowanceColumn = 0;
  for (let i = 0; i < header.length; i++) {
    const title = String(header[i] === null || header[i] === undefined ? '' : header[i]).trim();
    if (title === EMPLOYEE_SKILL_ALLOWANCE_HEADER) {
      allowanceColumn = i + 1;
      break;
    }
  }

  // 見つからなかったときだけ、既存の列より右側に新しく割り当てる
  if (allowanceColumn === 0) {
    allowanceColumn = lastColumn + 1;
    if (allowanceColumn > sheet.getMaxColumns()) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), allowanceColumn - sheet.getMaxColumns());
    }
    sheet.getRange(1, allowanceColumn).setValue(EMPLOYEE_SKILL_ALLOWANCE_HEADER);
  }

  return { 'sheet': sheet, 'column': allowanceColumn };
}

/**
 * スキル手当として読んだ値を金額（数値）に直す
 * 列を足した直後の空欄や、数値として読めない値は「まだ手当なし」の 0 円として扱う
 * @param {*} rawValue 名簿のスキル手当列に入っている値
 * @return {number} スキル手当（円/時）
 */
function normalizeSkillAllowance_(rawValue) {
  const amount = Number(rawValue);
  if (!isFinite(amount) || amount <= 0) {
    return SKILL_ALLOWANCE_DEFAULT;
  }
  return amount;
}

/**
 * 従業員番号ごとのスキル手当（円/時）の対応表をつくる
 * 給与計算と、スキル保存時の「今より下げない」判定の両方でこの対応表を使う
 * @return {Object} { '3': 50, ... }
 */
function getEmployeeSkillAllowanceMap() {
  const columns = ensureEmployeeSkillAllowanceColumn_();
  const sheet = columns.sheet;
  const lastRow = sheet.getLastRow();
  const allowanceByEmpId = {};
  if (lastRow < 2) {
    return allowanceByEmpId; // 見出し行しかない
  }

  const empIdValues = sheet.getRange(2, EMPLOYEE_COL_ID, lastRow - 1, 1).getValues();
  const allowanceValues = sheet.getRange(2, columns.column, lastRow - 1, 1).getValues();
  for (let i = 0; i < empIdValues.length; i++) {
    if (empIdValues[i][0] === '') {
      continue;
    }
    allowanceByEmpId[String(empIdValues[i][0])] = normalizeSkillAllowance_(allowanceValues[i][0]);
  }
  return allowanceByEmpId;
}

/**
 * 従業員1人のスキル手当（円/時）を取得する
 * @param {*} empId 従業員番号
 * @return {number} 名簿に保存されているスキル手当。未設定なら 0
 */
function getEmployeeSkillAllowance(empId) {
  const amount = getEmployeeSkillAllowanceMap()[String(empId)];
  return (amount === undefined) ? SKILL_ALLOWANCE_DEFAULT : amount;
}

/**
 * 従業員1人のスキル手当（円/時）を名簿に書き込む
 * 金額を決めるのは呼び出し側なので、ここでは下げないための判定はしない
 * @param {*} empId 従業員番号
 * @param {number} amount 保存するスキル手当
 * @return {boolean} 書き込めたら true、名簿に行が見つからなければ false
 */
function saveEmployeeSkillAllowance_(empId, amount) {
  const columns = ensureEmployeeSkillAllowanceColumn_();
  const targetRow = findEmployeeRow_(readEmployeeRows_(), empId);
  if (targetRow === null) {
    return false;
  }
  columns.sheet.getRange(targetRow.rowNumber, columns.column).setValue(amount);
  return true;
}
