/**
 * 従業員名簿に状態列と削除日時列を用意する。
 * 戻り値 - {sheet: GoogleAppsScript.Spreadsheet.Sheet, statusColumn: number, deletedAtColumn: number}
 */
function ensureEmployeeStatusColumns_() {
  const sheet = getSheet(SHEET_NAMES.EMPLOYEES);

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

  let nextColumn = lastColumn;
  if (statusColumn === 0) {
    nextColumn++;
    statusColumn = nextColumn;
  }
  if (deletedAtColumn === 0) {
    nextColumn++;
    deletedAtColumn = nextColumn;
  }

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
 * 名簿の見出しからID・氏名・雇用区分の列位置を調べる。
 * sheet - GoogleAppsScript.Spreadsheet.Sheet
 * 戻り値 - {id: number, name: number, partTimeMark: number}
 */
function getEmployeeIdentityColumns_(sheet) {
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (value) {
    return String(value).trim();
  });
  /**
   * 見出し候補のどれかに一致する列を探す。
   * names - string[]
   * fallback - number
   * 戻り値 - number
   */
  function findColumn(names, fallback) {
    const index = header.findIndex(function (title) { return names.indexOf(title) !== -1; });
    return index === -1 ? fallback : index + 1;
  }
  const id = findColumn(['従業員番号', '従業員ID'], EMPLOYEE_COL_ID);
  const name = findColumn(['氏名'], EMPLOYEE_COL_NAME);

  return {
    id: id,
    name: name,
    partTimeMark: findColumn(['雇用区分', '雇用区分の印'],
      id === EMPLOYEE_COL_ID && name === EMPLOYEE_COL_NAME ? EMPLOYEE_COL_PART_TIME_MARK : 0)
  };
}

/**
 * 従業員名簿の全行を、削除済みかどうかも含めて読み取る 一覧・氏名の対応表・追加・削除は、すべてこの関数が読んだ内容をもとに判断する。
 * 戻り値 - ({id: unknown, name: string, rowNumber: number, isDeleted: boolean})[]
 */
function readEmployeeRows_() {
  const columns = ensureEmployeeStatusColumns_();
  const sheet = columns.sheet;
  const lastRow = sheet.getLastRow();
  const rows = [];
  if (lastRow < 2) {
    return rows;
  }

  const identity = getEmployeeIdentityColumns_(sheet);
  const width = sheet.getLastColumn();
  const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
  for (let i = 0; i < values.length; i++) {
    const empId = values[i][identity.id - 1];
    if (empId === '') {
      continue;
    }

    const status = String(values[i][columns.statusColumn - 1] || '').trim();
    rows.push({
      'id': empId,
      'name': String(values[i][identity.name - 1] || '').trim(),
      'rowNumber': i + 2,
      'isDeleted': (status === EMPLOYEE_STATUS_DELETED)
    });
  }
  return rows;
}

/**
 * 画面に表示する従業員番号の文字列を作る 名簿の値そのものは変えず、表示のときだけ「01」「001」のゼロ埋めを外して「1」にする。
 * empId - unknown
 * 戻り値 - string
 */
function formatEmployeeNumber(empId) {
  const text = String(empId === null || empId === undefined ? '' : empId).trim();
  if (/^[0-9]+$/.test(text)) {
    return String(Number(text));
  }
  return text;
}

/**
 * 削除済みでない有効な従業員の一覧を取得する。
 * 戻り値 - ({id: unknown, name: string, number: string})[]
 */
function getEmployees() {
  const rows = readEmployeeRows_();
  const employees = [];
  for (let i = 0; i < rows.length; i++) {
    if (rows[i].isDeleted || rows[i].name === '') {
      continue;
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
 * ログイン状態を確認し、有効な従業員の一覧を返す。
 * authToken - string
 * 戻り値 - ({id: unknown, name: string, number: string})[]
 */
function getActiveEmployeeList(authToken) {
  requireLoggedIn_(authToken);
  return getEmployees();
}

/**
 * 有効な従業員にスキルと時給を組み合わせ、店長の従業員一覧用に返す。
 * authToken - string
 * 戻り値 - ({id: unknown, number: string, name: string, skills: {[key: string]: string}, hourlyWage: number})[]
 */
function getEmployeeDirectory(authToken) {
  requireManagerRole_(authToken);
  const employees = getEmployees();
  const skillMap = getEmployeeSkillMap();
  const rateMap = getEmployeeRateMap();
  return employees.map(function (employee) {
    const rate = rateMap[String(employee.id)];
    return {
      id: employee.id,
      number: employee.number,
      name: employee.name,
      skills: skillMap[String(employee.id)] || getDefaultSkills(),
      hourlyWage: rate.base + rate.skillAllowance
    };
  });
}

/**
 * 従業員番号と氏名の対応表をつくる（削除済みの人も含む） 過去の勤怠に残っている記録の氏名を表示するために使う。
 * 戻り値 - {[key: string]: string}
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
 * 有効な従業員だけの、従業員番号と氏名の対応表をつくる 打刻・本日の業務・スキルの保存先として選んでよい人かどうかの確認に使う。
 * 戻り値 - {[key: string]: string}
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
 * 従業員名簿から、指定した従業員番号の行を探す。
 * rows - ({id: unknown, name: string, rowNumber: number, isDeleted: boolean})[]
 * empId - unknown
 * 戻り値 - null または {id: unknown, name: string, rowNumber: number, isDeleted: boolean}
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
 * PINは名簿の末尾に文字列として保存する。
 * 戻り値 - {sheet: GoogleAppsScript.Spreadsheet.Sheet, column: number}
 */
function ensureEmployeePinColumn_() {
  const sheet = ensureEmployeeStatusColumns_().sheet;
  const lastColumn = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  for (let i = 0; i < header.length; i++) {
    if (String(header[i]).trim() === EMPLOYEE_PIN_HEADER) return { sheet: sheet, column: i + 1 };
  }
  const column = lastColumn + 1;
  if (column > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), column - sheet.getMaxColumns());
  sheet.getRange(1, column).setValue(EMPLOYEE_PIN_HEADER);
  return { sheet: sheet, column: column };
}

/**
 * 名簿にPIN列を用意し、未設定の従業員に初期PINを入れる初期移行処理。
 * 戻り値 - なし
 */
function migrateEmployeePins_() {
  const columns = ensureEmployeePinColumn_();
  const rows = readEmployeeRows_();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i].name === '') continue;
    const cell = columns.sheet.getRange(rows[i].rowNumber, columns.column);
    if (cell.getValue() === '') cell.setNumberFormat('@').setValue(EMPLOYEE_PIN_DEFAULT);
  }
}

/**
 * 有効な従業員かを確かめ、名簿に保存されたPINを文字列として取得する。
 * empId - unknown
 * 戻り値 - string
 */
function readEmployeePin_(empId) {
  const row = findEmployeeRow_(readEmployeeRows_(), empId);
  if (!row || row.isDeleted || row.name === '') throw new Error('この従業員は選べません');
  const columns = ensureEmployeePinColumn_();
  return String(columns.sheet.getRange(row.rowNumber, columns.column).getValue());
}

/**
 * 店長権限・入力形式・確認欄の一致を確かめ、対象従業員のPINを書き換える。
 * empId - string
 * newPin - string
 * confirmPin - string
 * authToken - string
 * 戻り値 - {message: string}
 */
function changeEmployeePin(empId, newPin, confirmPin, authToken) {
  requireManagerRole_(authToken);
  requireValidEmployeePin_(newPin);
  requireValidEmployeePin_(confirmPin);
  if (newPin !== confirmPin) throw new Error('新しいPINと確認欄が一致しません');
  return withAppLock_(function () {
    const row = findEmployeeRow_(readEmployeeRows_(), empId);
    if (!row || row.isDeleted || row.name === '') throw new Error('この従業員は選べません。対象を選び直してください');
    const columns = ensureEmployeePinColumn_();
    columns.sheet.getRange(row.rowNumber, columns.column).setNumberFormat('@').setValue(newPin);
    return { message: row.name + 'のPINを変更しました' };
  });
}

/**
 * 店長権限を確認し、新しい従業員を登録する。
 * lastName - string
 * firstName - string
 * pin - string
 * authToken - string
 * 戻り値 - {message: string, employees: ({id: unknown, name: string, number: string})[]}
 */
function addEmployee(lastName, firstName, pin, authToken) {
  requireManagerRole_(authToken);
  requireValidEmployeePin_(pin);
  return withAppLock_(function () { return addEmployee_(lastName, firstName, pin); });
}

/**
 * 氏名を検証し、全従業員の最大ID+1で新しい名簿行を作る。
 * lastName - string
 * firstName - string
 * pin - string
 * 戻り値 - {message: string, employees: ({id: unknown, name: string, number: string})[]}
 */
function addEmployee_(lastName, firstName, pin) {

  lastName = String(lastName === null || lastName === undefined ? '' : lastName).trim();
  firstName = String(firstName === null || firstName === undefined ? '' : firstName).trim();
  if (lastName === '' || firstName === '') {
    throw new Error('姓と名を入力してください');
  }

  const columns = ensureEmployeeStatusColumns_();
  const rows = readEmployeeRows_();

  let maxId = 0;
  for (let i = 0; i < rows.length; i++) {
    const value = String(rows[i].id).trim();
    if (/^[0-9]+$/.test(value)) maxId = Math.max(maxId, Number(value));
  }
  const employeeId = maxId + 1;
  if (!Number.isSafeInteger(employeeId)) throw new Error('従業員番号の上限に達しました');

  const fullName = lastName + ' ' + firstName;
  const pinColumn = ensureEmployeePinColumn_();
  const allowanceColumn = ensureEmployeeSkillAllowanceColumn_();
  const sheet = columns.sheet;
  const identity = getEmployeeIdentityColumns_(sheet);
  const rowNumber = sheet.getLastRow() + 1;
  if (rowNumber > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 1);
  const values = new Array(sheet.getLastColumn()).fill('');
  values[identity.id - 1] = employeeId;
  values[identity.name - 1] = fullName;
  if (identity.partTimeMark) values[identity.partTimeMark - 1] = '○';
  values[columns.statusColumn - 1] = EMPLOYEE_STATUS_ACTIVE;
  values[pinColumn.column - 1] = pin;
  values[allowanceColumn.column - 1] = SKILL_ALLOWANCE_DEFAULT;
  sheet.getRange(rowNumber, pinColumn.column).setNumberFormat('@');
  sheet.getRange(rowNumber, 1, 1, values.length).setValues([values]);

  return {
    'message': fullName + ' を追加しました（従業員番号 : ' + formatEmployeeNumber(employeeId) + '）',
    'employees': getEmployees()
  };
}

/**
 * 削除する前に、対象の従業員と削除の影響を確認するための情報を返す（店長だけが実行できる）。
 * empId - unknown
 * authToken - string
 * 戻り値 - {id: unknown, name: string, number: string}
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
 * 従業員を削除済みにする（店長だけが実行できる） 名簿の行と氏名はそのまま残し、状態列と削除日時列だけを書き換える。
 * empId - unknown
 * authToken - string
 * 戻り値 - {message: string, employees: ({id: unknown, name: string, number: string})[]}
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

/**
 * 従業員名簿にスキル手当列を用意する。
 * 戻り値 - {sheet: GoogleAppsScript.Spreadsheet.Sheet, column: number}
 */
function ensureEmployeeSkillAllowanceColumn_() {

  const statusColumns = ensureEmployeeStatusColumns_();
  const sheet = statusColumns.sheet;

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
 * スキル手当として読んだ値を金額（数値）に直す 列を足した直後の空欄や、数値として読めない値は「まだ手当なし」の 0 円として扱う。
 * rawValue - unknown
 * 戻り値 - number
 */
function normalizeSkillAllowance_(rawValue) {
  const amount = Number(rawValue);
  if (!isFinite(amount) || amount <= 0) {
    return SKILL_ALLOWANCE_DEFAULT;
  }
  return amount;
}

/**
 * 従業員番号ごとのスキル手当（円/時）の対応表をつくる 給与計算と、スキル保存時の「今より下げない」判定の両方でこの対応表を使う。
 * 戻り値 - {[key: string]: number}
 */
function getEmployeeSkillAllowanceMap() {
  const columns = ensureEmployeeSkillAllowanceColumn_();
  const sheet = columns.sheet;
  const lastRow = sheet.getLastRow();
  const allowanceByEmpId = {};
  if (lastRow < 2) {
    return allowanceByEmpId;
  }

  const empIdValues = sheet.getRange(2, getEmployeeIdentityColumns_(sheet).id, lastRow - 1, 1).getValues();
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
 * 従業員1人のスキル手当（円/時）を取得する。
 * empId - unknown
 * 戻り値 - number
 */
function getEmployeeSkillAllowance(empId) {
  const amount = getEmployeeSkillAllowanceMap()[String(empId)];
  return (amount === undefined) ? SKILL_ALLOWANCE_DEFAULT : amount;
}

/**
 * 従業員1人のスキル手当（円/時）を名簿に書き込む 金額を決めるのは呼び出し側なので、ここでは下げないための判定はしない。
 * empId - unknown
 * amount - number
 * 戻り値 - boolean
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
