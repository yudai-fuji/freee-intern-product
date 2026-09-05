// 従業員の取得・追加・氏名管理を行う
//console.log()

//　従業員一覧を取得する関数
function getEmployees() {
  //従業員名簿のシートを取得
  const empSheet = getSheet(SHEET_NAMES.EMPLOYEES)
  
  //データが入っている最後の行番号を取得
  const last_row = empSheet.getLastRow()
  
  //範囲取得(2行目から、A列から、~行分、2列分)
  const empRange = empSheet.getRange(2, 1, last_row - 1, 2)

  //全員分の[従業員番号, 名前]を2次元配列で取得
  const empValues = empRange.getValues();// の配列をまとめて取得

  //従業員の最終結果を格納する配列を用意(letじゃない理由はpushは配列を入れ替えるのではなく、加えるだけだから)
  const employees = [];

  //[{id: 1, name: 'freee　花子' },...]のようにオブジェクト要素を持つ配列を作成。またIdか名前が空白のものは加えない。
  for (let i = 0; i < empValues.length; i++) {
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
  console.log(empValues)
  console.log(employees)
  return employees
}

// 従業員を追加
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

// 従業員IDと名前の対応表をつくる
function getEmployeeNameMap() {
  var employees = getEmployees();
  var nameByEmpId = {};
  for (var i = 0; i < employees.length; i++) {
    nameByEmpId[String(employees[i]['id'])] = employees[i]['name'];
  }
  return nameByEmpId;
}