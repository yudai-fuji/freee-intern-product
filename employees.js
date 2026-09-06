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
    const empId = empValues[i][0];
    const empName = empValues[i][1];
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

// 従業員を追加
function addEmployee(lastName, firstName) {
  //.trim()で文字列に含まれる余分な空白を削除
  lastName = String(lastName).trim();
  firstName = String(firstName).trim();

  if (lastName === "" || firstName === "") {
    throw new Error('姓と名を入力してください')
  }

  //'従業員名簿'をスプシから取得
  const empSheet = getSheet(SHEET_NAMES.EMPLOYEES)

  //データが入っている最後の行番号を取得
  const last_row = empSheet.getLastRow()

  //(開始行, 開始列, 行数, 列数)
  const empRange = empSheet.getRange(2, 1, last_row - 1, 2);

  //全員分の[従業員番号, 名前]を2次元配列で取得
  const empValues = empRange.getValues();
  console.log(empValues)
  console.log(empValues.length)
  
  //このあとに使う変数、-1なのはまだ見つかってないことを明示する一般的な使い方らしい
  let targetRow = -1

  for (var i = 0; i < empValues.length; i++) {
    const empId = empValues[i][0];
    const empName = empValues[i][1];
    if (empId !== "" && empName === "") {
      targetRow = i + 2;
      break;
    }
  }

  if (targetRow === -1) {
    throw new Error('空いている従業員番号がありません')
  }

  //〇のUniコードを代入
  const partTimeMark = '\u25CB'

  //対象の行の2列目に名前、3列目に〇をセット
  empSheet.getRange(targetRow, 2).setValue(lastName + ' ' + firstName)
  empSheet.getRange(targetRow, 3).setValue(partTimeMark)

  return '従業員を追加しました'
}

// 従業員IDと名前の対応表をつくる
function getEmployeeNameMap() {
  //上の関数から取得するものを参照して
  const employees = getEmployees();

  //空のオブジェクトを作成
  const nameByEmpId = {};

  //[employees[i].id]というキーに、名前を対応させる。ちなみにキーは文字列なのでString()を使用。
  for (let i = 0; i < employees.length; i++) {
    nameByEmpId[String(employees[i].id)] = employees[i].name;
  }
  
  return nameByEmpId;
}