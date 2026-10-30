 
 
 

 




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

function getEmployeeIdentityColumns_(sheet) {
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (value) {
    return String(value).trim();
  });
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

 





function formatEmployeeNumber(empId) {
  const text = String(empId === null || empId === undefined ? '' : empId).trim();
  if (/^[0-9]+$/.test(text)) {
    return String(Number(text));
  }
  return text;
}

 




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

 





function getActiveEmployeeList(authToken) {
  requireLoggedIn_(authToken);
  return getEmployees();
}

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

 




function getActiveEmployeeNameMap() {
  const employees = getEmployees();
  const nameByEmpId = {};
  for (let i = 0; i < employees.length; i++) {
    nameByEmpId[String(employees[i].id)] = employees[i].name;
  }
  return nameByEmpId;
}

 





function findEmployeeRow_(rows, empId) {
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i].id) === String(empId)) {
      return rows[i];
    }
  }
  return null;
}

 
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

function migrateEmployeePins_() {
  const columns = ensureEmployeePinColumn_();
  const rows = readEmployeeRows_();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i].name === '') continue;
    const cell = columns.sheet.getRange(rows[i].rowNumber, columns.column);
    if (cell.getValue() === '') cell.setNumberFormat('@').setValue(EMPLOYEE_PIN_DEFAULT);
  }
}

function readEmployeePin_(empId) {
  const row = findEmployeeRow_(readEmployeeRows_(), empId);
  if (!row || row.isDeleted || row.name === '') throw new Error('この従業員は選べません');
  const columns = ensureEmployeePinColumn_();
  return String(columns.sheet.getRange(row.rowNumber, columns.column).getValue());
}

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

 








function addEmployee(lastName, firstName, pin, authToken) {
  requireManagerRole_(authToken);
  requireValidEmployeePin_(pin);
  return withAppLock_(function () { return addEmployee_(lastName, firstName, pin); });
}

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

 





function normalizeSkillAllowance_(rawValue) {
  const amount = Number(rawValue);
  if (!isFinite(amount) || amount <= 0) {
    return SKILL_ALLOWANCE_DEFAULT;
  }
  return amount;
}

 




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

 




function getEmployeeSkillAllowance(empId) {
  const amount = getEmployeeSkillAllowanceMap()[String(empId)];
  return (amount === undefined) ? SKILL_ALLOWANCE_DEFAULT : amount;
}

 






function saveEmployeeSkillAllowance_(empId, amount) {
  const columns = ensureEmployeeSkillAllowanceColumn_();
  const targetRow = findEmployeeRow_(readEmployeeRows_(), empId);
  if (targetRow === null) {
    return false;
  }
  columns.sheet.getRange(targetRow.rowNumber, columns.column).setValue(amount);
  return true;
}
