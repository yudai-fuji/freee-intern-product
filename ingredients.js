// 発注管理が使う食材・業者・発注記録のマスタ（スプレッドシートの読み書きと、共通の書式変換）を扱う
//
// 扱い方の前提
// ・食材マスタは「食材名・発注単位・紐づく業者」を持つ。発注画面の選択肢はここから作る
// ・業者マスタは連絡先と発注条件（受注締切・営業日・リードタイム・臨時休業日）を持つ
// ・未設定の条件は推測で埋めない。設定がそろうまで「設定が必要です」と表示する
// ・単位の異なる数量は合算・換算しない

// 画面とGASの間でやり取りする日付・日時の形式
const INGREDIENT_DATE_PATTERN = /^\d{4}\/\d{2}\/\d{2}$/;
const INGREDIENT_DATE_TIME_PATTERN = /^\d{4}\/\d{2}\/\d{2} \d{2}:\d{2}$/;
const INGREDIENT_TIME_PATTERN = /^([01][0-9]|2[0-3]):[0-5][0-9]$/;

// 目視の概数として受け付ける数量の上限（入力ミスで極端な値が入るのを防ぐ）
const INGREDIENT_MAX_QUANTITY = 100000;

// 納品可能日を探すときに見る日数の上限
const INGREDIENT_PLAN_SEARCH_DAYS = 120;

/**
 * 食材・業者・発注のシートが無ければ作成し、ヘッダーを既定の列構成に揃える
 * 日付や「0.5」などが勝手に別の型へ変換されないよう、新規作成時は全列を文字列書式にする
 * @param {string} sheetName シート名
 * @param {Array<string>} header 見出し行
 * @return {Sheet} 対象シート
 */
function ensureIngredientSheet(sheetName, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  let isNewSheet = false;
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    isNewSheet = true;
  }

  // 見出しの数よりシートの列が少ないと見出し行を書けないので、足りない分だけ列を足す
  const shortageColumns = header.length - sheet.getMaxColumns();
  if (shortageColumns > 0) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), shortageColumns);
  }

  const headerRange = sheet.getRange(1, 1, 1, header.length);
  const currentHeader = headerRange.getValues()[0];
  for (let i = 0; i < header.length; i++) {
    if (currentHeader[i] !== header[i]) {
      headerRange.setValues([header]); // 既存データの行には触れず、見出し行だけを揃える
      break;
    }
  }

  const dataRowCount = sheet.getMaxRows() - 1;
  if (isNewSheet && dataRowCount > 0) {
    sheet.getRange(2, 1, dataRowCount, header.length).setNumberFormat('@');
  }
  return sheet;
}

/**
 * 食材マスタのシートを用意する（新規作成したときだけ初期候補の食材名を入れる）
 * すでに行があるシートには追記しない
 */
function ensureIngredientMasterSheet() {
  const sheet = ensureIngredientSheet(SHEET_NAMES.INGREDIENT_MASTERS, INGREDIENT_MASTER_SHEET_HEADER);
  if (sheet.getLastRow() >= 2) {
    return sheet;
  }
  const rows = [];
  const updatedAt = formatIngredientDateTime(new Date());
  for (let i = 0; i < INGREDIENT_DEFAULT_MASTERS.length; i++) {
    const master = INGREDIENT_DEFAULT_MASTERS[i];
    rows.push([
      newIngredientId('ING'), master.name, master.category, master.unit,
      '', '', '', '', '○', updatedAt,
      '' // 発注単位は食材マスタで管理する
    ]);
  }
  appendIngredientRows(sheet, rows, INGREDIENT_MASTER_SHEET_HEADER.length);
  return sheet;
}

function ensureIngredientEventSheet() {
  return ensureIngredientSheet(SHEET_NAMES.INGREDIENT_EVENTS, INGREDIENT_EVENT_SHEET_HEADER);
}

function ensureIngredientSupplierSheet() {
  return ensureIngredientSheet(SHEET_NAMES.INGREDIENT_SUPPLIERS, INGREDIENT_SUPPLIER_SHEET_HEADER);
}

function ensureIngredientOrderSheet() {
  return ensureIngredientSheet(SHEET_NAMES.INGREDIENT_ORDERS, INGREDIENT_ORDER_SHEET_HEADER);
}

/**
 * シートの末尾に行をまとめて追記する
 * 追記先の行数が足りないときは必要な分だけ行を増やし、書き込む前に文字列書式にしておく
 * @return {number} 追記した先頭行のシート上の行番号（1行も無いときは -1）
 */
function appendIngredientRows(sheet, rows, columnCount) {
  if (rows.length === 0) {
    return -1;
  }
  const startRow = sheet.getLastRow() + 1;
  const shortageRows = startRow + rows.length - 1 - sheet.getMaxRows();
  if (shortageRows > 0) {
    sheet.insertRowsAfter(sheet.getMaxRows(), shortageRows);
  }
  const range = sheet.getRange(startRow, 1, rows.length, columnCount);
  range.setNumberFormat('@');
  range.setValues(rows);
  return startRow;
}

/**
 * シートの2行目以降を「見出し名をキーにしたオブジェクト」の配列として読み込む
 * 値はすべて文字列に揃えるので、読み出し側で Date か文字列かを気にしなくてよい
 * @return {Array<Object>} { rowNumber: シート上の行番号, values: {見出し名: 文字列} } の配列
 */
function readIngredientRows(sheet, header) {
  const lastRow = sheet.getLastRow();
  const rows = [];
  if (lastRow < 2) {
    return rows;
  }
  const values = sheet.getRange(2, 1, lastRow - 1, header.length).getValues();
  for (let i = 0; i < values.length; i++) {
    const record = {};
    let hasValue = false;
    for (let c = 0; c < header.length; c++) {
      const text = toIngredientCellText(values[i][c]);
      record[header[c]] = text;
      if (text !== '') {
        hasValue = true;
      }
    }
    if (!hasValue) {
      continue; // 空行は読み飛ばす
    }
    rows.push({ 'rowNumber': i + 2, 'values': record });
  }
  return rows;
}

/**
 * シートのセル値を画面で扱う文字列に揃える
 * Date が入っていた場合は日付または日時の文字列へ変換する（手作業で入力された行でも読めるようにするため）
 */
function toIngredientCellText(value) {
  if (value === null || value === undefined) {
    return '';
  }
  if (value instanceof Date) {
    const hasTime = (value.getHours() !== 0 || value.getMinutes() !== 0);
    return hasTime ? formatIngredientDateTime(value) : formatIngredientDate(value);
  }
  return String(value).trim();
}

function formatIngredientDate(date) {
  return Utilities.formatDate(date, TIME_ZONE, 'yyyy/MM/dd');
}

function formatIngredientDateTime(date) {
  return Utilities.formatDate(date, TIME_ZONE, 'yyyy/MM/dd HH:mm');
}

/**
 * 新しいIDを作る（種類が分かるよう接頭辞を付ける）
 */
function newIngredientId(prefix) {
  return prefix + '-' + Utilities.getUuid().replace(/-/g, '').substring(0, 12);
}

function normalizeIngredientText(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

/**
 * 画面から来た日付を 'yyyy/MM/dd' に揃える（未入力は空文字）
 * input[type=date] の 'yyyy-MM-dd' も受け付ける
 * @return {string} 正規化した日付。形式が不正なら例外を投げる
 */
function normalizeIngredientDateInput(value, label) {
  const raw = normalizeIngredientText(value);
  const text = raw.replace(/-/g, '/');
  if (text === '') {
    return '';
  }
  if (!INGREDIENT_DATE_PATTERN.test(text)) {
    throw new Error(label + 'の日付の形式が不正です: ' + raw);
  }
  if (!parseIngredientDate(text)) {
    throw new Error(label + 'に存在しない日付が指定されています: ' + text);
  }
  return text;
}

/**
 * 'yyyy/MM/dd' を日本時間の Date（0時0分）に変換する。変換できなければ null
 */
function parseIngredientDate(dateText) {
  if (!INGREDIENT_DATE_PATTERN.test(dateText)) {
    return null;
  }
  const parts = dateText.split('/');
  const year = Number(parts[0]);
  const month = Number(parts[1]);
  const day = Number(parts[2]);
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
    return null; // 2026/02/31 のような存在しない日付を弾く
  }
  return date;
}

/**
 * 'yyyy/MM/dd HH:mm' を日本時間の Date に変換する。変換できなければ null
 */
function parseIngredientDateTime(dateTimeText) {
  if (!INGREDIENT_DATE_TIME_PATTERN.test(dateTimeText)) {
    return null;
  }
  const date = parseIngredientDate(dateTimeText.substring(0, 10));
  if (!date) {
    return null;
  }
  date.setHours(Number(dateTimeText.substring(11, 13)));
  date.setMinutes(Number(dateTimeText.substring(14, 16)));
  return date;
}

/**
 * 目視の概数として数量を受け取る
 * @param {*} value 画面から来た数量
 * @param {string} label エラー文に出す項目名
 * @param {boolean} allowZero 0 を許すか（残量の更新では 0 を許し、移動量では許さない）
 * @return {number} 小数第3位までに丸めた数量
 */
function parseIngredientQuantity(value, label, allowZero) {
  const text = normalizeIngredientText(value);
  if (text === '') {
    throw new Error(label + 'を入力してください');
  }
  const quantity = Number(text);
  if (!isFinite(quantity)) {
    throw new Error(label + 'は数値で入力してください: ' + text);
  }
  if (quantity < 0) {
    throw new Error(label + 'は0以上で入力してください');
  }
  if (!allowZero && quantity === 0) {
    throw new Error(label + 'は0より大きい値を入力してください');
  }
  if (quantity > INGREDIENT_MAX_QUANTITY) {
    throw new Error(label + 'が大きすぎます（' + INGREDIENT_MAX_QUANTITY + 'までで入力してください）');
  }
  return roundIngredientQuantity(quantity);
}

/**
 * 目視の概数どうしの足し引きで誤差が出ないよう、小数第3位までに丸める
 */
function roundIngredientQuantity(quantity) {
  return Math.round(quantity * 1000) / 1000;
}

/**
 * 数量を画面表示用の文字列にする（1.000 のような余分な0を付けない）
 */
function formatIngredientQuantity(quantity) {
  return String(roundIngredientQuantity(quantity));
}

function addIngredientDays(date, days) {
  const result = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  result.setDate(result.getDate() + days);
  return result;
}

/**
 * 日付に曜日を添えた表示文字列を作る（曜日は固定値ではなく日付から求める）
 */
function formatIngredientDateWithWeekday(date) {
  return formatIngredientDate(date) + '（' + INGREDIENT_WEEKDAY_LABELS[date.getDay()] + '）';
}

// ここから下はシートの読み取り

/**
 * 食材マスタを読み込む
 * @return {Array<Object>} 食材の設定（名前・発注単位・業者など）の配列
 */
function getIngredientMasters() {
  const sheet = ensureIngredientMasterSheet();
  const rows = readIngredientRows(sheet, INGREDIENT_MASTER_SHEET_HEADER);
  const masters = [];
  for (let i = 0; i < rows.length; i++) {
    const values = rows[i].values;
    if (values['食材ID'] === '' || values['食材名'] === '') {
      continue;
    }
    masters.push({
      'id': values['食材ID'],
      'name': values['食材名'],
      'category': values['区分'] || INGREDIENT_CATEGORIES[0],
      'defaultUnit': values['既定単位'],
      'supplierId': values['業者ID'],
      'active': values['有効'] !== '',
      // 発注単位が空欄なら既定単位をそのまま使う（発注単位を設定する前の行でも発注できるようにする）
      'orderUnit': values['発注単位'] !== '' ? values['発注単位'] : values['既定単位'],
      'rowNumber': rows[i].rowNumber
    });
  }
  return masters;
}

/** 旧マスタの廃止済み指示を表示・本文追記へ持ち込まない。シートは書き換えない。 */
function normalizeIngredientSupplierNote_(value) {
  const retiredInstruction = /商品コード|product\s*code|item\s*code/i;
  return normalizeIngredientText(value).split(/\r\n|\r|\n/).map(function (line) {
    if (!retiredInstruction.test(line)) {
      return line; // 本文追記などの有効な注意事項は、句読点を含めてそのまま保つ。
    }
    const parts = line.match(/[^。；;、,]+[。；;、,]?/g) || [];
    return parts.filter(function (part) {
      return !retiredInstruction.test(part);
    }).join('').replace(/[；;、,]+\s*$/, '').trim();
  }).filter(function (line) { return line.trim() !== ''; }).join('\n');
}

/** 業者と発注条件を読み込む。 */
function getIngredientSuppliers() {
  const sheet = ensureIngredientSupplierSheet();
  const rows = readIngredientRows(sheet, INGREDIENT_SUPPLIER_SHEET_HEADER);
  const suppliers = [];
  for (let i = 0; i < rows.length; i++) {
    const values = rows[i].values;
    if (values['業者ID'] === '') {
      continue;
    }
    const leadTimeText = values['納品リードタイム日数'];
    suppliers.push({
      'id': values['業者ID'],
      'name': values['業者名'],
      'deadlineTime': values['受注締切時刻'],
      'orderDays': parseIngredientWeekdayList(values['受注営業日']),
      'shipDays': parseIngredientWeekdayList(values['出荷営業日']),
      'deliveryDays': parseIngredientWeekdayList(values['納品営業日']),
      'leadTimeDays': (leadTimeText === '' || !isFinite(Number(leadTimeText))) ? null : Number(leadTimeText),
      'leadTimeCount': values['リードタイムの数え方'],
      'closedDates': parseIngredientDateList(values['臨時休業日']),
      'pickupDays': parseIngredientWeekdayList(values['店舗受取可能日']),
      'note': normalizeIngredientSupplierNote_(values['備考']),
      'address': values['住所'],
      'phone': values['電話番号'],
      'email': values['メールアドレス'],
      'businessHours': values['営業時間'],
      'closedDays': values['定休日'],
      'mainItems': values['主な取扱商品'],
      'orderNote': normalizeIngredientSupplierNote_(values['発注時の注意事項']),
      'rowNumber': rows[i].rowNumber
    });
  }
  return suppliers;
}

/**
 * 発注の記録を読み込む
 */
function getIngredientOrders() {
  const sheet = ensureIngredientOrderSheet();
  const rows = readIngredientRows(sheet, INGREDIENT_ORDER_SHEET_HEADER);
  const orders = [];
  for (let i = 0; i < rows.length; i++) {
    const values = rows[i].values;
    if (values['発注ID'] === '') {
      continue;
    }
    const quantity = Number(values['数量']);
    orders.push({
      'id': values['発注ID'],
      'ingredientId': values['食材ID'],
      'name': values['食材名'],
      'supplierId': values['業者ID'],
      'quantity': isFinite(quantity) ? roundIngredientQuantity(quantity) : 0,
      'unit': values['単位'],
      'orderedAt': values['発注日時'],
      'expectedDate': values['入荷予定日'],
      'status': values['状態'],
      'note': values['備考'],
      // ここから下は発注管理で足した列。発注管理を使い始める前に記録された行では空欄になる
      'orderNo': values['発注番号'],
      'orderedBy': values['発注者'],
      'mailTo': values['送信先メール'],
      'rowNumber': rows[i].rowNumber
    });
  }
  return orders;
}

/**
 * '月,火,水' のような文字列を曜日番号（0=日）の配列にする
 */
function parseIngredientWeekdayList(text) {
  const days = [];
  if (normalizeIngredientText(text) === '') {
    return days;
  }
  const parts = String(text).split(/[,、\s]+/);
  for (let i = 0; i < parts.length; i++) {
    const index = INGREDIENT_WEEKDAY_LABELS.indexOf(parts[i].trim());
    if (index !== -1 && days.indexOf(index) === -1) {
      days.push(index);
    }
  }
  days.sort(function (a, b) { return a - b; });
  return days;
}

/**
 * 曜日番号の配列を '月・火' のような表示文字列にする
 */
function formatIngredientWeekdayList(days) {
  const labels = [];
  for (let i = 0; i < days.length; i++) {
    labels.push(INGREDIENT_WEEKDAY_LABELS[days[i]]);
  }
  return labels.join('・');
}

/**
 * '2026/09/16,2026/09/17' のような文字列を日付文字列の配列にする
 */
function parseIngredientDateList(text) {
  const dates = [];
  if (normalizeIngredientText(text) === '') {
    return dates;
  }
  const parts = String(text).split(/[,、\s]+/);
  for (let i = 0; i < parts.length; i++) {
    const dateText = parts[i].trim().replace(/-/g, '/');
    if (INGREDIENT_DATE_PATTERN.test(dateText) && parseIngredientDate(dateText)) {
      dates.push(dateText);
    }
  }
  return dates;
}

// ここから下は業者の発注条件から「次に納品できる日」と「その締切」を求める処理

/**
 * その日が業者の臨時休業日かどうか
 */
function isIngredientClosedDate(supplier, date) {
  return supplier.closedDates.indexOf(formatIngredientDate(date)) !== -1;
}

/**
 * 発注してから納品できる日を求める
 * 出荷できる最初の日を探し、そこからリードタイムを数え、納品できる日かつ店舗が受け取れる日まで後ろへずらす
 * @return {Date|null} 納品日。条件を満たす日が見つからなければ null
 */
function findIngredientDeliveryDate(supplier, orderDate) {
  let shipDate = null;
  for (let i = 0; i < INGREDIENT_PLAN_SEARCH_DAYS; i++) {
    const candidate = addIngredientDays(orderDate, i);
    if (supplier.shipDays.indexOf(candidate.getDay()) !== -1 && !isIngredientClosedDate(supplier, candidate)) {
      shipDate = candidate;
      break;
    }
  }
  if (!shipDate) {
    return null;
  }

  let deliveryDate = null;
  if (supplier.leadTimeCount === '暦日') {
    deliveryDate = addIngredientDays(shipDate, supplier.leadTimeDays);
  } else if (supplier.leadTimeDays === 0) {
    deliveryDate = shipDate;
  } else {
    // 営業日で数える場合は、出荷日の翌日から「納品できる曜日かつ臨時休業日でない日」だけを数える
    let remaining = supplier.leadTimeDays;
    for (let i = 1; i <= INGREDIENT_PLAN_SEARCH_DAYS; i++) {
      const candidate = addIngredientDays(shipDate, i);
      if (supplier.deliveryDays.indexOf(candidate.getDay()) === -1 || isIngredientClosedDate(supplier, candidate)) {
        continue;
      }
      remaining--;
      if (remaining === 0) {
        deliveryDate = candidate;
        break;
      }
    }
    if (!deliveryDate) {
      return null;
    }
  }

  for (let i = 0; i < INGREDIENT_PLAN_SEARCH_DAYS; i++) {
    const candidate = addIngredientDays(deliveryDate, i);
    if (supplier.deliveryDays.indexOf(candidate.getDay()) === -1) {
      continue;
    }
    if (isIngredientClosedDate(supplier, candidate)) {
      continue;
    }
    if (supplier.pickupDays.indexOf(candidate.getDay()) === -1) {
      continue; // 店舗が受け取れない日は納品日にしない
    }
    return candidate;
  }
  return null;
}

/**
 * 業者ごとに「次に納品できる日」と「それに間に合う発注締切」を求める
 * 条件が未設定のときは推測せず、設定が必要な項目を返す
 * @param {Object} supplier 業者の設定
 * @param {Date} now 現在日時
 * @return {Object} { available, missing, text, orderDate, deadline, deliveryDate }
 */
function buildIngredientSupplierPlan(supplier, now) {
  const missing = [];
  if (!INGREDIENT_TIME_PATTERN.test(supplier.deadlineTime)) {
    missing.push('受注締切');
  }
  if (supplier.orderDays.length === 0) {
    missing.push('受注営業日');
  }
  if (supplier.shipDays.length === 0) {
    missing.push('出荷営業日');
  }
  if (supplier.deliveryDays.length === 0) {
    missing.push('納品営業日');
  }
  if (supplier.leadTimeDays === null || supplier.leadTimeDays < 0) {
    missing.push('納品までの日数');
  }
  if (INGREDIENT_LEAD_TIME_COUNTS.indexOf(supplier.leadTimeCount) === -1) {
    missing.push('日数の数え方');
  }
  if (supplier.pickupDays.length === 0) {
    missing.push('店舗の受取可能日');
  }
  if (missing.length > 0) {
    return {
      'available': false,
      'missing': missing,
      'text': supplier.name + 'の発注条件が未設定です（' + missing.join('・') + '）。設定が必要です'
    };
  }

  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  for (let i = 0; i < INGREDIENT_PLAN_SEARCH_DAYS; i++) {
    const orderDate = addIngredientDays(today, i);
    if (supplier.orderDays.indexOf(orderDate.getDay()) === -1 || isIngredientClosedDate(supplier, orderDate)) {
      continue;
    }
    const deadline = parseIngredientDateTime(formatIngredientDate(orderDate) + ' ' + supplier.deadlineTime);
    if (!deadline || deadline.getTime() <= now.getTime()) {
      continue; // 締切を過ぎた発注日は選べない
    }
    const deliveryDate = findIngredientDeliveryDate(supplier, orderDate);
    if (!deliveryDate) {
      continue;
    }
    return {
      'available': true,
      'missing': [],
      'orderDate': formatIngredientDateWithWeekday(orderDate),
      'deadline': formatIngredientDateWithWeekday(orderDate) + ' ' + supplier.deadlineTime,
      'deliveryDate': formatIngredientDateWithWeekday(deliveryDate),
      'closedWeekdayText': buildIngredientClosedWeekdayText(supplier),
      'text': '次に納品できるのは ' + formatIngredientDateWithWeekday(deliveryDate) +
        ' です。この分の発注締切は ' + formatIngredientDateWithWeekday(orderDate) + ' ' + supplier.deadlineTime + ' です'
    };
  }
  return {
    'available': false,
    'missing': [],
    'text': supplier.name + 'は設定された条件では納品できる日が見つかりませんでした。発注条件を確認してください'
  };
}

/**
 * 納品できない曜日を説明する文言を作る（曜日は設定から求め、固定値にしない）
 */
function buildIngredientClosedWeekdayText(supplier) {
  const closedDays = [];
  for (let day = 0; day < INGREDIENT_WEEKDAY_LABELS.length; day++) {
    if (supplier.deliveryDays.indexOf(day) === -1) {
      closedDays.push(day);
    }
  }
  if (closedDays.length === 0) {
    return '';
  }
  return formatIngredientWeekdayList(closedDays) + 'は納品休業です';
}

// ここから下は「食材在庫履歴」シートへの操作記録（発注管理の二重送信を防ぐために使う）

/**
 * 操作キーから履歴の行を探す（再送の判定に使う）
 */
function findIngredientEventByKey(eventSheet, operationKey) {
  const rows = readIngredientRows(eventSheet, INGREDIENT_EVENT_SHEET_HEADER);
  for (let i = rows.length - 1; i >= 0; i--) {
    if (rows[i].values['操作キー'] === operationKey) {
      return rows[i];
    }
  }
  return null;
}

/**
 * 履歴に「処理中」の行を先に作る
 * 途中で失敗した操作をあとから見分けられるようにするため、保存の前に記録する
 * @return {number} 追加した行番号
 */
function startIngredientEvent(eventSheet, operationKey, type) {
  const row = [
    newIngredientId('EVT'), operationKey, formatIngredientDateTime(new Date()), type,
    '', '', '', '', '', '', '', '処理中'
  ];
  return appendIngredientRows(eventSheet, [row], INGREDIENT_EVENT_SHEET_HEADER.length);
}

/**
 * 履歴の行を「完了」にし、何をどれだけ動かしたかを書き残す
 */
function completeIngredientEvent(eventSheet, rowNumber, result) {
  const values = [
    result.lotId || '', result.partnerLotId || '', result.ingredientId || '',
    (result.quantity === undefined || result.quantity === null) ? '' : formatIngredientQuantity(result.quantity),
    result.unit || '', result.message, result.actor || '', '完了'
  ];
  const range = eventSheet.getRange(rowNumber, 5, 1, values.length);
  range.setNumberFormat('@');
  range.setValues([values]);
}

/**
 * 履歴の行を「失敗」にする
 * 記録は残していないので、同じ操作キーでもう一度実行できる状態に戻す
 */
function failIngredientEvent(eventSheet, rowNumber, error) {
  const range = eventSheet.getRange(rowNumber, 10, 1, 3); // J:内容 K:実施者 L:処理状態
  range.setNumberFormat('@');
  range.setValues([['保存できませんでした: ' + error.message, '', '失敗']]);
}

/**
 * 食材IDから食材マスタの1件を取り出す
 */
function findIngredientMasterById(masters, ingredientId) {
  const id = normalizeIngredientText(ingredientId);
  for (let i = 0; i < masters.length; i++) {
    if (masters[i].id === id) {
      return masters[i];
    }
  }
  throw new Error('食材マスタに存在しない食材です: ' + id);
}
