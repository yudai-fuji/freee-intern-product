// 食材管理（目視での残量把握・入荷・仕込み・冷凍/解凍・期限・発注条件）を扱う
//
// 扱い方の前提
// ・在庫は「明細（ロット）」単位で管理する。同じ食材でも入荷・仕込み・期限・冷凍日時が違えば別の明細にする
// ・「原材料／仕込み済み」（区分）と「冷蔵／冷凍／解凍中」（保存状態）は別の属性として持つ
// ・ラベルの期限と、現在の状態に対する使用期限は別の列に持ち、どちらも自動では延長しない
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
 * 食材管理用のシートが無ければ作成し、ヘッダーを既定の列構成に揃える
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
 * 在庫数量・期限は入れない。すでに行があるシートには追記しない
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
      '', '', '', '', '○', updatedAt
    ]);
  }
  appendIngredientRows(sheet, rows, INGREDIENT_MASTER_SHEET_HEADER.length);
  return sheet;
}

function ensureIngredientLotSheet() {
  return ensureIngredientSheet(SHEET_NAMES.INGREDIENT_LOTS, INGREDIENT_LOT_SHEET_HEADER);
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
 * 画面から来た日時を 'yyyy/MM/dd HH:mm' に揃える（未入力は空文字）
 * input[type=datetime-local] の 'yyyy-MM-ddTHH:mm' も受け付ける
 */
function normalizeIngredientDateTimeInput(value, label) {
  const raw = normalizeIngredientText(value);
  let text = raw.replace(/-/g, '/').replace('T', ' ');
  if (text === '') {
    return '';
  }
  if (text.length === 19) {
    text = text.substring(0, 16); // 秒まで送られてきた場合は分までにそろえる
  }
  if (!INGREDIENT_DATE_TIME_PATTERN.test(text)) {
    throw new Error(label + 'の日時の形式が不正です: ' + raw);
  }
  if (!parseIngredientDateTime(text)) {
    throw new Error(label + 'に存在しない日時が指定されています: ' + text);
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

/**
 * 選択肢に含まれる値かを確かめる
 */
function assertIngredientChoice(value, choices, label) {
  const text = normalizeIngredientText(value);
  if (choices.indexOf(text) === -1) {
    throw new Error('不正な' + label + 'です: ' + text);
  }
  return text;
}

/**
 * 日付の差を「日数」で返す（時刻は見ない）
 */
function diffIngredientDays(fromDate, toDate) {
  const from = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
  const to = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
  return Math.round((to.getTime() - from.getTime()) / (24 * 60 * 60 * 1000));
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

// ここから下はシートの読み取りと、画面に渡す情報の組み立て

/**
 * 食材マスタを読み込む
 * @return {Array<Object>} 食材の設定（区分・既定単位・最低残量・業者など）の配列
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
    const minQuantityText = values['最低残量'];
    masters.push({
      'id': values['食材ID'],
      'name': values['食材名'],
      'category': values['区分'] || INGREDIENT_CATEGORIES[0],
      'defaultUnit': values['既定単位'],
      'minQuantity': (minQuantityText === '' || !isFinite(Number(minQuantityText))) ? null : Number(minQuantityText),
      'minUnit': values['最低残量単位'],
      'notifyExpirySoon': values['期限接近通知'] !== '',
      'supplierId': values['業者ID'],
      'active': values['有効'] !== '',
      'rowNumber': rows[i].rowNumber
    });
  }
  return masters;
}

/**
 * 在庫明細（ロット）を読み込む
 * @return {Array<Object>} 明細の配列（行状態が '保留' や '終了' のものも含む）
 */
function getIngredientLots() {
  const sheet = ensureIngredientLotSheet();
  const rows = readIngredientRows(sheet, INGREDIENT_LOT_SHEET_HEADER);
  const lots = [];
  for (let i = 0; i < rows.length; i++) {
    if (rows[i].values['在庫ID'] === '') {
      continue;
    }
    lots.push(toIngredientLot(rows[i].values, rows[i].rowNumber));
  }
  return lots;
}

/**
 * シートの1行分を在庫明細のオブジェクトに変換する
 */
function toIngredientLot(values, rowNumber) {
  const quantity = Number(values['数量']);
  const version = Number(values['版数']);
  return {
    'id': values['在庫ID'],
    'ingredientId': values['食材ID'],
    'name': values['食材名'],
    'category': values['区分'],
    'quantity': isFinite(quantity) ? roundIngredientQuantity(quantity) : 0,
    'unit': values['単位'],
    'storageState': values['保存状態'],
    'labelExpiryKind': values['ラベル期限種別'],
    'labelExpiryDate': values['ラベル期限'],
    'useByDate': values['使用期限'],
    'receivedDate': values['入荷日'],
    'preparedAt': values['仕込み日時'],
    'preparedAtUnknown': values['仕込み日時不明'] !== '',
    'frozenAt': values['冷凍日時'],
    'frozenKind': values['冷凍区分'],
    'thawStartedAt': values['解凍開始日時'],
    'thawCompletedAt': values['解凍完了日時'],
    'sourceLotId': values['由来在庫ID'],
    'checkedAt': values['確認日時'],
    'checkedBy': values['確認者'],
    'note': values['備考'],
    'status': values['行状態'] || '有効',
    'version': isFinite(version) ? version : 1,
    'updatedAt': values['更新日時'],
    'rowNumber': rowNumber
  };
}

/**
 * 在庫明細のオブジェクトをシートの1行分（文字列の配列）に戻す
 */
function toIngredientLotRow(lot) {
  return [
    lot.id, lot.ingredientId, lot.name, lot.category,
    formatIngredientQuantity(lot.quantity), lot.unit, lot.storageState,
    lot.labelExpiryKind, lot.labelExpiryDate, lot.useByDate,
    lot.receivedDate, lot.preparedAt, lot.preparedAtUnknown ? '○' : '',
    lot.frozenAt, lot.frozenKind, lot.thawStartedAt, lot.thawCompletedAt, lot.sourceLotId,
    lot.checkedAt, lot.checkedBy, lot.note, lot.status, String(lot.version), lot.updatedAt
  ];
}

/**
 * 業者と発注条件を読み込む
 */
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
      'note': values['備考'],
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
      'lotId': values['入荷在庫ID'],
      'note': values['備考'],
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

/**
 * その明細について、いま管理すべき期限が「ラベルの期限」か「店舗で決めた使用期限」かを判定する
 * 開封・仕込み・店内冷凍・解凍をした品はラベルの期限では管理できないため、使用期限が無ければ要確認にする
 * @return {Object} { kind: 期限の種類, date: 期限日, state: 状態, label: 表示文, days: 残り日数 }
 */
function buildIngredientExpiryStatus(lot, today) {
  let kind = '';
  let dateText = '';

  if (lot.useByDate !== '') {
    kind = '使用期限';
    dateText = lot.useByDate;
  } else if (canUseIngredientLabelExpiry(lot)) {
    kind = lot.labelExpiryKind;
    dateText = lot.labelExpiryDate;
  }

  if (dateText === '') {
    return {
      'kind': '', 'date': '', 'state': '要確認', 'days': null,
      'label': lot.labelExpiryKind === '不明' ? '期限不明（要確認）' : '期限未設定（要確認）'
    };
  }

  const expiryDate = parseIngredientDate(dateText);
  if (!expiryDate) {
    return { 'kind': kind, 'date': dateText, 'state': '要確認', 'days': null, 'label': '期限の内容を確認してください（' + dateText + '）' };
  }

  const days = diffIngredientDays(today, expiryDate);
  if (days < 0) {
    return { 'kind': kind, 'date': dateText, 'state': '期限切れ', 'days': days, 'label': kind + ' ' + dateText + '（期限切れ）' };
  }
  if (days <= INGREDIENT_EXPIRY_SOON_DAYS) {
    const remainText = (days === 0) ? '本日まで' : ('あと' + days + '日');
    return { 'kind': kind, 'date': dateText, 'state': '期限接近', 'days': days, 'label': kind + ' ' + dateText + '（' + remainText + '）' };
  }
  return { 'kind': kind, 'date': dateText, 'state': '余裕あり', 'days': days, 'label': kind + ' ' + dateText };
}

/**
 * ラベルの期限をそのまま使ってよい明細かを判定する
 * 未開封で表示条件どおりに保存している原材料（冷蔵のままの品と、納品時から冷凍の品）だけを対象にする
 */
function canUseIngredientLabelExpiry(lot) {
  if (lot.labelExpiryDate === '' || lot.labelExpiryKind === '' || lot.labelExpiryKind === '不明') {
    return false;
  }
  if (lot.category !== '原材料' || lot.preparedAt !== '') {
    return false;
  }
  if (lot.storageState === '冷蔵') {
    return lot.frozenKind === '' && lot.thawStartedAt === '';
  }
  if (lot.storageState === '冷凍') {
    return lot.frozenKind === '納品時冷凍';
  }
  return false; // 解凍中はラベルの表示条件から外れるため、店舗で決めた使用期限が要る
}

/**
 * 明細ごとの「確認事項」を組み立てる（色だけでなく文言でも状態が分かるようにするため）
 * @return {Array<string>} 確認事項の文言
 */
function buildIngredientLotIssues(lot, expiry) {
  const issues = [];
  if (lot.status === '保留') {
    issues.push('保存が完了していない明細です（残量に数えていません）');
  }
  if (expiry.state === '期限切れ') {
    issues.push('期限切れ');
  } else if (expiry.state === '期限接近') {
    issues.push('期限が近い');
  } else if (expiry.state === '要確認') {
    issues.push(lot.labelExpiryKind === '不明' ? '期限が不明' : '期限が未設定');
  }
  if (lot.preparedAtUnknown) {
    issues.push('仕込み日時が不明');
  }
  if (lot.storageState === '解凍中') {
    issues.push('解凍中（完了を記録してください）');
  }
  if (lot.quantity === 0) {
    issues.push('残量0（使い切り済みか確認してください）');
  }
  if (expiry.state !== '期限切れ' && lot.note.indexOf('期限切れ') !== -1) {
    issues.push('冷凍・解凍の前に期限切れの記録があります（備考を確認）');
  }
  return issues;
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

// ここから下は画面上部に出す確認事項の組み立て

/**
 * 食材ごとの残量を単位別に集計する（単位が違うものは合算しない）
 * @return {Object} 食材IDごとの { fresh: {単位: 数量}, frozen: {単位: 数量}, thawing: {単位: 数量} }
 */
function summarizeIngredientQuantities(lots) {
  const summary = {};
  for (let i = 0; i < lots.length; i++) {
    const lot = lots[i];
    if (lot.status !== '有効') {
      continue; // 保存が完了していない行や終了した行は残量に数えない
    }
    if (!summary[lot.ingredientId]) {
      summary[lot.ingredientId] = { 'fresh': {}, 'frozen': {}, 'thawing': {} };
    }
    const target = (lot.storageState === '冷凍') ? 'frozen' : ((lot.storageState === '解凍中') ? 'thawing' : 'fresh');
    const bucket = summary[lot.ingredientId][target];
    bucket[lot.unit] = roundIngredientQuantity((bucket[lot.unit] || 0) + lot.quantity);
  }
  return summary;
}

/**
 * 単位別の数量を「1.5袋／300g」のような表示文字列にする（換算はしない）
 */
function formatIngredientQuantityMap(quantityMap) {
  const parts = [];
  for (const unit in quantityMap) {
    if (Object.prototype.hasOwnProperty.call(quantityMap, unit)) {
      parts.push(formatIngredientQuantity(quantityMap[unit]) + unit);
    }
  }
  return parts.length === 0 ? '0' : parts.join('／');
}

/**
 * 画面上部に出す確認事項（期限未設定・期限接近・残量不足・発注）を組み立てる
 * 根拠のない予測はせず、記録済みの内容と設定条件から言えることだけを文言にする
 * @return {Object} 種類ごとの確認事項の配列
 */
function buildIngredientAlerts(lots, masters, suppliers, orders, now) {
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const alerts = { 'expiryUnset': [], 'expiryNear': [], 'lowStock': [], 'order': [], 'pending': [] };

  const supplierById = {};
  for (let i = 0; i < suppliers.length; i++) {
    supplierById[suppliers[i].id] = suppliers[i];
  }
  const planBySupplierId = {};

  for (let i = 0; i < lots.length; i++) {
    const lot = lots[i];
    if (lot.status === '終了') {
      continue;
    }
    if (lot.status === '保留') {
      alerts.pending.push({
        'title': lot.name + '（' + formatIngredientQuantity(lot.quantity) + lot.unit + '）',
        'detail': '保存が途中で終わった可能性がある明細です。実物を確認し、必要なら残量を更新してください',
        'lotId': lot.id
      });
      continue;
    }
    const expiry = buildIngredientExpiryStatus(lot, today);
    const lotTitle = lot.name + '（' + formatIngredientQuantity(lot.quantity) + lot.unit + '・' + lot.storageState + '）';
    if (expiry.state === '要確認') {
      alerts.expiryUnset.push({
        'title': lotTitle,
        'detail': expiry.label + '。実物を確認して使用期限を登録してください',
        'lotId': lot.id
      });
    } else if (expiry.state === '期限切れ') {
      alerts.expiryNear.push({
        'title': lotTitle,
        'detail': expiry.kind + ' ' + expiry.date + ' を過ぎています。使用前に必ず確認してください',
        'lotId': lot.id
      });
    } else if (expiry.state === '期限接近') {
      alerts.expiryNear.push({
        'title': lotTitle,
        'detail': expiry.kind + ' ' + expiry.date + ' まで' + (expiry.days === 0 ? '本日' : 'あと' + expiry.days + '日') + 'です',
        'lotId': lot.id
      });
    }
    if (lot.preparedAtUnknown) {
      alerts.expiryUnset.push({
        'title': lotTitle,
        'detail': '仕込み日時が不明のままです。分かった時点で登録してください',
        'lotId': lot.id
      });
    }
  }

  const summary = summarizeIngredientQuantities(lots);
  const ordersByIngredientId = {};
  for (let i = 0; i < orders.length; i++) {
    if (orders[i].status !== '発注済み') {
      continue;
    }
    if (!ordersByIngredientId[orders[i].ingredientId]) {
      ordersByIngredientId[orders[i].ingredientId] = [];
    }
    ordersByIngredientId[orders[i].ingredientId].push(orders[i]);
  }

  for (let i = 0; i < masters.length; i++) {
    const master = masters[i];
    if (!master.active || master.minQuantity === null || master.minUnit === '') {
      continue; // 最低残量が未設定の食材は通知しない（推測で判定しない）
    }
    const totals = summary[master.id] || { 'fresh': {}, 'frozen': {}, 'thawing': {} };
    const freshQuantity = roundIngredientQuantity((totals.fresh[master.minUnit] || 0) + (totals.thawing[master.minUnit] || 0));
    const frozenQuantity = totals.frozen[master.minUnit] || 0;
    const detailParts = [];

    if (frozenQuantity > 0) {
      detailParts.push('冷凍在庫が ' + formatIngredientQuantity(frozenQuantity) + master.minUnit + 'あります（解凍・仕込みが必要です）');
    }
    const otherUnitText = buildIngredientOtherUnitText(totals, master.minUnit);
    if (otherUnitText !== '') {
      detailParts.push('単位の違う在庫があります（合算していません）: ' + otherUnitText);
    }
    const pendingOrders = ordersByIngredientId[master.id] || [];
    for (let p = 0; p < pendingOrders.length; p++) {
      detailParts.push('発注済み ' + formatIngredientQuantity(pendingOrders[p].quantity) + pendingOrders[p].unit +
        '（入荷予定 ' + (pendingOrders[p].expectedDate === '' ? '未設定' : pendingOrders[p].expectedDate) + '・入荷を記録するまで在庫には加えません）');
    }

    const supplier = supplierById[master.supplierId];
    if (!supplier) {
      detailParts.push('業者が未設定です。発注締切・納品日を出すには業者と発注条件の設定が必要です');
    } else {
      if (!planBySupplierId[supplier.id]) {
        planBySupplierId[supplier.id] = buildIngredientSupplierPlan(supplier, now);
      }
      const plan = planBySupplierId[supplier.id];
      if (plan.available && plan.closedWeekdayText !== '') {
        detailParts.push(plan.closedWeekdayText);
      }
      detailParts.push(plan.text);
    }

    if (freshQuantity < master.minQuantity) {
      alerts.lowStock.push({
        'title': master.name,
        'detail': master.name + 'の残量が設定した最低量を下回っています（いま ' + formatIngredientQuantity(freshQuantity) + master.minUnit +
          '／最低 ' + formatIngredientQuantity(master.minQuantity) + master.minUnit + '）。' + detailParts.join('。'),
        'ingredientId': master.id
      });
    } else if (master.notifyExpirySoon && hasIngredientExpiringLot(lots, master.id, today)) {
      alerts.lowStock.push({
        'title': master.name,
        'detail': master.name + 'は最低量以上ありますが、期限が近い在庫があります（いま ' + formatIngredientQuantity(freshQuantity) + master.minUnit +
          '）。使い切りか追加発注かを判断してください。' + detailParts.join('。'),
        'ingredientId': master.id
      });
    }
  }

  for (let i = 0; i < orders.length; i++) {
    const order = orders[i];
    if (order.status !== '発注済み') {
      continue;
    }
    const title = order.name + '（' + formatIngredientQuantity(order.quantity) + order.unit + '）';
    if (order.expectedDate === '') {
      alerts.order.push({ 'title': title, 'detail': '発注済みですが入荷予定日が未設定です', 'orderId': order.id });
      continue;
    }
    const expectedDate = parseIngredientDate(order.expectedDate);
    if (expectedDate && diffIngredientDays(today, expectedDate) < 0) {
      alerts.order.push({
        'title': title,
        'detail': '入荷予定日 ' + order.expectedDate + ' を過ぎていますが入荷が記録されていません。届いていれば入荷を記録してください',
        'orderId': order.id
      });
    } else {
      alerts.order.push({
        'title': title,
        'detail': '発注済み（入荷予定 ' + order.expectedDate + '）。入荷を記録するまで在庫には加わりません',
        'orderId': order.id
      });
    }
  }

  return alerts;
}

/**
 * 最低残量の単位以外で持っている在庫を説明する文字列を作る
 */
function buildIngredientOtherUnitText(totals, minUnit) {
  const parts = [];
  const buckets = [totals.fresh, totals.thawing, totals.frozen];
  const labels = ['冷蔵', '解凍中', '冷凍'];
  for (let i = 0; i < buckets.length; i++) {
    for (const unit in buckets[i]) {
      if (Object.prototype.hasOwnProperty.call(buckets[i], unit) && unit !== minUnit && buckets[i][unit] > 0) {
        parts.push(labels[i] + ' ' + formatIngredientQuantity(buckets[i][unit]) + unit);
      }
    }
  }
  return parts.join('・');
}

/**
 * その食材に期限が近い（または切れた）在庫があるかを調べる
 */
function hasIngredientExpiringLot(lots, ingredientId, today) {
  for (let i = 0; i < lots.length; i++) {
    if (lots[i].ingredientId !== ingredientId || lots[i].status !== '有効') {
      continue;
    }
    const state = buildIngredientExpiryStatus(lots[i], today).state;
    if (state === '期限接近' || state === '期限切れ') {
      return true;
    }
  }
  return false;
}

/**
 * 食材管理画面に必要な情報を一度にまとめて返す（画面から呼ぶ入口）
 * 食材管理は店長画面の機能なので、ここで店長かどうかを確認する
 * @param {string} authToken ログイントークン
 * @return {Object} 設定値・在庫明細・業者・発注・確認事項
 */
function getIngredientBoard(authToken) {
  requireManagerRole_(authToken);
  return buildIngredientBoard_();
}

/**
 * 食材管理画面に必要な情報を組み立てる
 * 権限確認済みの保存処理からも呼ぶため、ここでは権限を確認しない
 * @return {Object} 設定値・在庫明細・業者・発注・確認事項
 */
function buildIngredientBoard_() {
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const masters = getIngredientMasters();
  const lots = getIngredientLots();
  const suppliers = getIngredientSuppliers();
  const orders = getIngredientOrders();

  const supplierNameById = {};
  for (let i = 0; i < suppliers.length; i++) {
    supplierNameById[suppliers[i].id] = suppliers[i].name;
  }

  const lotViews = [];
  for (let i = 0; i < lots.length; i++) {
    const lot = lots[i];
    if (lot.status === '終了' && lot.quantity === 0) {
      continue; // 使い切った明細は一覧には出さない（履歴はシートに残る）
    }
    const expiry = buildIngredientExpiryStatus(lot, today);
    lotViews.push({
      'id': lot.id,
      'ingredientId': lot.ingredientId,
      'name': lot.name,
      'category': lot.category,
      'quantity': lot.quantity,
      'quantityText': formatIngredientQuantity(lot.quantity) + lot.unit,
      'unit': lot.unit,
      'storageState': lot.storageState,
      'labelExpiryKind': lot.labelExpiryKind,
      'labelExpiryDate': lot.labelExpiryDate,
      'useByDate': lot.useByDate,
      'receivedDate': lot.receivedDate,
      'preparedAt': lot.preparedAt,
      'preparedAtUnknown': lot.preparedAtUnknown,
      'preparedAtText': lot.preparedAtUnknown ? '不明（要確認）' : (lot.preparedAt === '' ? NO_VALUE_LABEL : lot.preparedAt),
      'frozenAt': lot.frozenAt,
      'frozenKind': lot.frozenKind,
      // 納品時から冷凍の品は店内で冷凍した日時を持たないので、そのことが分かる表示にする
      'frozenText': lot.frozenAt !== ''
        ? (lot.frozenAt + '（' + (lot.frozenKind || '区分未設定') + '）')
        : (lot.frozenKind === '' ? NO_VALUE_LABEL : lot.frozenKind + '（店内で冷凍した記録なし）'),
      'thawStartedAt': lot.thawStartedAt,
      'thawCompletedAt': lot.thawCompletedAt,
      'sourceLotId': lot.sourceLotId,
      'checkedAt': lot.checkedAt,
      'checkedBy': lot.checkedBy,
      'note': lot.note,
      'status': lot.status,
      'version': lot.version,
      'updatedAt': lot.updatedAt,
      'expiry': expiry,
      'issues': buildIngredientLotIssues(lot, expiry)
    });
  }

  const supplierViews = [];
  for (let i = 0; i < suppliers.length; i++) {
    const supplier = suppliers[i];
    supplierViews.push({
      'id': supplier.id,
      'name': supplier.name,
      'deadlineTime': supplier.deadlineTime,
      'orderDaysText': formatIngredientWeekdayList(supplier.orderDays),
      'shipDaysText': formatIngredientWeekdayList(supplier.shipDays),
      'deliveryDaysText': formatIngredientWeekdayList(supplier.deliveryDays),
      'leadTimeDays': supplier.leadTimeDays,
      'leadTimeCount': supplier.leadTimeCount,
      'closedDatesText': supplier.closedDates.join('、'),
      'pickupDaysText': formatIngredientWeekdayList(supplier.pickupDays),
      'note': supplier.note,
      'plan': buildIngredientSupplierPlan(supplier, now)
    });
  }

  const orderViews = [];
  for (let i = 0; i < orders.length; i++) {
    const order = orders[i];
    orderViews.push({
      'id': order.id,
      'ingredientId': order.ingredientId,
      'name': order.name,
      'supplierId': order.supplierId,
      'supplierName': supplierNameById[order.supplierId] || '',
      'quantity': order.quantity,
      'unit': order.unit,
      'orderedAt': order.orderedAt,
      'expectedDate': order.expectedDate,
      'status': order.status,
      'note': order.note
    });
  }

  return {
    'today': formatIngredientDate(now),
    'nowText': formatIngredientDateTime(now),
    'categories': INGREDIENT_CATEGORIES,
    'storageStates': INGREDIENT_STORAGE_STATES,
    'units': INGREDIENT_UNITS,
    'expiryKinds': INGREDIENT_EXPIRY_KINDS,
    'leadTimeCounts': INGREDIENT_LEAD_TIME_COUNTS,
    'weekdayLabels': INGREDIENT_WEEKDAY_LABELS,
    'expirySoonDays': INGREDIENT_EXPIRY_SOON_DAYS,
    'masters': masters,
    'lots': lotViews,
    'suppliers': supplierViews,
    'orders': orderViews,
    'alerts': buildIngredientAlerts(lots, masters, suppliers, orders, now)
  };
}

// ここから下は保存処理
// 保存はすべて runIngredientOperation() を通し、次の3点を守る
// ・同時に走る保存を締め出す（LockService）
// ・同じ操作キーの再送では在庫を二重に増減しない（食材在庫履歴に操作キーを残す）
// ・明細の版数が画面の表示時と違えば保存せず、古い画面からの上書きを防ぐ

// 在庫明細シートの「行状態」列（1始まり。INGREDIENT_LOT_SHEET_HEADER の並びと揃える）
const INGREDIENT_LOT_COL_STATUS = 22;

/**
 * 保存処理の共通の入れ物
 * @param {string} operationKey 画面が1回の保存ごとに発行する操作キー（再送の判定に使う）
 * @param {string} type 履歴に残す操作の種別
 * @param {Function} executor 実際の保存処理。{ message, lotId, partnerLotId, ingredientId, quantity, unit, actor } を返す
 * @return {Object} { message: 画面に出す文言, board: 最新の食材管理情報 }
 */
function runIngredientOperation(operationKey, type, executor) {
  const key = normalizeIngredientText(operationKey);
  if (key === '') {
    throw new Error('操作キーがありません。画面を読み込み直してからもう一度お試しください');
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    throw new Error('ほかの保存処理が実行中です。少し待ってからもう一度お試しください');
  }

  let message = '';
  try {
    const eventSheet = ensureIngredientEventSheet();
    const existing = findIngredientEventByKey(eventSheet, key);
    let isDuplicated = false;
    if (existing) {
      const status = existing.values['処理状態'];
      if (status === '完了') {
        // 同じ操作の再送。すでに記録済みなので在庫は動かさず、前回の結果をそのまま返す
        message = existing.values['内容'] + '（この操作はすでに記録されています）';
        isDuplicated = true;
      } else if (status === '処理中') {
        // 前回の実行が途中で止まっている。取り返しがつかなくなる前に利用者へ確認してもらう
        throw new Error('前回の保存が完了していない可能性があります。最新の情報を読み込み、在庫の内容を確認してください');
      }
      // '失敗' のときは在庫を変えていないので、同じ操作キーのままやり直せる
    }

    if (!isDuplicated) {
      const eventRowNumber = startIngredientEvent(eventSheet, key, type);
      let result;
      try {
        result = executor();
      } catch (error) {
        // 入力の誤りなどで保存できなかった記録を残す（同じ操作キーでやり直せるようにする）
        failIngredientEvent(eventSheet, eventRowNumber, error);
        throw error;
      }
      SpreadsheetApp.flush(); // 在庫の書き込みを確定させてから履歴を完了にする
      completeIngredientEvent(eventSheet, eventRowNumber, result);
      message = result.message;
    }
  } finally {
    lock.releaseLock();
  }

  return { 'message': message, 'board': buildIngredientBoard_() };
}

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
 * 在庫は変えていないので、同じ操作キーでもう一度保存できる状態に戻す
 */
function failIngredientEvent(eventSheet, rowNumber, error) {
  const range = eventSheet.getRange(rowNumber, 10, 1, 3); // J:内容 K:実施者 L:処理状態
  range.setNumberFormat('@');
  range.setValues([['保存できませんでした: ' + error.message, '', '失敗']]);
}

/**
 * 残量を確認した人を求める
 * 従業員画面で選択中の従業員があればその氏名、無ければ実行中のGoogleアカウントを使う
 * @return {string} 確認者名（分からなければ空文字）
 */
function resolveIngredientActor(actorEmpId) {
  const empId = normalizeIngredientText(actorEmpId);
  if (empId !== '') {
    const nameByEmpId = getEmployeeNameMap();
    if (nameByEmpId[empId]) {
      return nameByEmpId[empId];
    }
  }
  try {
    const email = Session.getActiveUser().getEmail();
    if (email) {
      return email;
    }
  } catch (error) {
    // 実行者を取得できない設定のときは確認者を空のままにする
  }
  return '';
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

/**
 * 在庫IDから明細を1件取り出す（見つからなければ例外）
 */
function findIngredientLotById(lotSheet, lotId) {
  const id = normalizeIngredientText(lotId);
  const rows = readIngredientRows(lotSheet, INGREDIENT_LOT_SHEET_HEADER);
  for (let i = 0; i < rows.length; i++) {
    if (rows[i].values['在庫ID'] === id) {
      return toIngredientLot(rows[i].values, rows[i].rowNumber);
    }
  }
  throw new Error('対象の在庫明細が見つかりません: ' + id);
}

/**
 * 画面が表示していたときから明細が変わっていないかを確かめる（古い画面からの上書きを防ぐ）
 */
function assertIngredientLotVersion(lot, expectedVersion) {
  const version = Number(normalizeIngredientText(expectedVersion));
  if (!isFinite(version) || version !== lot.version) {
    throw new Error(lot.name + 'の明細は別の操作で更新されています。最新の情報を読み込んでからやり直してください');
  }
  if (lot.status === '保留') {
    throw new Error(lot.name + 'の明細は保存が完了していません。内容を確認してから操作してください');
  }
  if (lot.status === '終了') {
    throw new Error(lot.name + 'の明細は使い切り済みです');
  }
}

/**
 * 明細の1行をシートに書き戻す（版数を1つ進める）
 */
function writeIngredientLot(lotSheet, lot) {
  lot.version = lot.version + 1;
  lot.updatedAt = formatIngredientDateTime(new Date());
  const range = lotSheet.getRange(lot.rowNumber, 1, 1, INGREDIENT_LOT_SHEET_HEADER.length);
  range.setNumberFormat('@');
  range.setValues([toIngredientLotRow(lot)]);
}

/**
 * 明細を新しく1行追加する
 * @param {string} status '有効' か '保留'。在庫の移動では先に '保留' で追加し、元の明細を減らしてから '有効' にする
 * @return {number} 追加した行番号
 */
function appendIngredientLot(lotSheet, lot, status) {
  lot.status = status;
  lot.version = 1;
  lot.updatedAt = formatIngredientDateTime(new Date());
  return appendIngredientRows(lotSheet, [toIngredientLotRow(lot)], INGREDIENT_LOT_SHEET_HEADER.length);
}

/**
 * '保留' で追加した明細を '有効' にする
 */
function activateIngredientLot(lotSheet, rowNumber) {
  const range = lotSheet.getRange(rowNumber, INGREDIENT_LOT_COL_STATUS, 1, 1);
  range.setNumberFormat('@');
  range.setValue('有効');
}

/**
 * 在庫を分ける（冷凍・解凍・脂の取り分け）ときの共通処理
 * 先に移動先を '保留' で作り、元の明細を減らしてから '有効' にするので、
 * 途中で失敗しても増えた側だけが残ることがない（'保留' の行は残量に数えず要確認として表示する）
 * @return {string} 追加した明細の在庫ID
 */
function moveIngredientQuantity(lotSheet, sourceLot, movedQuantity, newLot) {
  const newRowNumber = appendIngredientLot(lotSheet, newLot, '保留');
  sourceLot.quantity = roundIngredientQuantity(sourceLot.quantity - movedQuantity);
  writeIngredientLot(lotSheet, sourceLot);
  SpreadsheetApp.flush();
  activateIngredientLot(lotSheet, newRowNumber);
  return newLot.id;
}

/**
 * 移動量が「0より大きく元の数量以下」かを確かめる
 */
function assertIngredientMovableQuantity(lot, quantity, label) {
  if (quantity > lot.quantity) {
    throw new Error(label + 'が在庫の数量を超えています（在庫 ' + formatIngredientQuantity(lot.quantity) + lot.unit + '）');
  }
}

/**
 * 明細を作るときの共通のひな形
 */
function createIngredientLotBase(master, quantity, unit, storageState, actor) {
  const now = formatIngredientDateTime(new Date());
  return {
    'id': newIngredientId('LOT'),
    'ingredientId': master.id,
    'name': master.name,
    'category': master.category,
    'quantity': quantity,
    'unit': unit,
    'storageState': storageState,
    'labelExpiryKind': '',
    'labelExpiryDate': '',
    'useByDate': '',
    'receivedDate': '',
    'preparedAt': '',
    'preparedAtUnknown': false,
    'frozenAt': '',
    'frozenKind': '',
    'thawStartedAt': '',
    'thawCompletedAt': '',
    'sourceLotId': '',
    'checkedAt': now,
    'checkedBy': actor,
    'note': '',
    'status': '有効',
    'version': 1,
    'updatedAt': now
  };
}

/**
 * 備考に1行追記する（元の情報を消さずに残すために使う）
 */
function appendIngredientNote(note, addition) {
  if (addition === '') {
    return note;
  }
  return note === '' ? addition : (note + ' / ' + addition);
}

/**
 * 入荷を記録する
 * 数量・単位・入荷日・保存状態・ラベルの期限と種類を登録し、新しい明細を1件作る
 * @param {Object} input 画面から送られた入力内容
 * @return {Object} { message, board }
 */
function saveIngredientReceive(input) {
  const payload = input || {};
  requireManagerRole_(payload.authToken);
  return runIngredientOperation(payload.operationKey, '入荷', function () {
    const masters = getIngredientMasters();
    const master = findIngredientMasterById(masters, payload.ingredientId);
    const quantity = parseIngredientQuantity(payload.quantity, '入荷数量', false);
    const unit = assertIngredientChoice(payload.unit, INGREDIENT_UNITS, '単位');
    const storageState = assertIngredientChoice(payload.storageState, ['冷蔵', '冷凍'], '保存状態');
    const receivedDate = normalizeIngredientDateInput(payload.receivedDate, '入荷日');
    if (receivedDate === '') {
      throw new Error('入荷日を入力してください');
    }
    const expiryKind = assertIngredientChoice(payload.expiryKind, INGREDIENT_EXPIRY_KINDS, '期限の種類');
    const expiryDate = normalizeIngredientDateInput(payload.expiryDate, 'ラベルの期限');
    if (expiryKind !== '不明' && expiryDate === '') {
      throw new Error('ラベルの期限を入力してください（分からない場合は期限の種類を「不明」にしてください）');
    }
    if (expiryKind === '不明' && expiryDate !== '') {
      throw new Error('期限の種類が「不明」のときは期限日を入力できません');
    }
    const actor = resolveIngredientActor(payload.actorEmpId);

    // 発注との結び付けは書き込みの前に確かめる（在庫だけ増えて発注が残る状態を作らないため）
    const orderId = normalizeIngredientText(payload.orderId);
    const orderRow = (orderId === '') ? null : findIngredientReceivableOrder(orderId, master.id);

    const lotSheet = ensureIngredientLotSheet();
    const lot = createIngredientLotBase(master, quantity, unit, storageState, actor);
    lot.labelExpiryKind = expiryKind;
    lot.labelExpiryDate = expiryDate;
    lot.receivedDate = receivedDate;
    lot.note = normalizeIngredientText(payload.note);
    if (storageState === '冷凍') {
      // 納品時から冷凍の品。店内で冷凍した品と区別するため冷凍日時は入れず、区分だけを残す
      lot.frozenKind = '納品時冷凍';
    }
    appendIngredientLot(lotSheet, lot, '有効');

    let orderText = '';
    if (orderRow) {
      markIngredientOrderReceived(orderRow, lot.id);
      orderText = '。発注分を入荷済みにしました';
    }

    return {
      'message': master.name + ' の入荷を記録しました（' + formatIngredientQuantity(quantity) + unit + '・' + storageState + '）' + orderText,
      'lotId': lot.id,
      'ingredientId': master.id,
      'quantity': quantity,
      'unit': unit,
      'actor': actor
    };
  });
}

/**
 * 入荷として記録できる発注かどうかを確かめ、その行を返す
 * @return {Object} 発注の行（{ rowNumber, values }）
 */
function findIngredientReceivableOrder(orderId, ingredientId) {
  const rows = readIngredientRows(ensureIngredientOrderSheet(), INGREDIENT_ORDER_SHEET_HEADER);
  for (let i = 0; i < rows.length; i++) {
    if (rows[i].values['発注ID'] !== orderId) {
      continue;
    }
    if (rows[i].values['状態'] !== '発注済み') {
      throw new Error('この発注はすでに「' + rows[i].values['状態'] + '」になっています');
    }
    if (rows[i].values['食材ID'] !== ingredientId) {
      throw new Error('発注した食材と入荷した食材が違います。入荷内容を確認してください');
    }
    return rows[i];
  }
  throw new Error('対象の発注が見つかりません: ' + orderId);
}

/**
 * 発注の記録を「入荷済み」にして、入荷した明細と結び付ける
 */
function markIngredientOrderReceived(orderRow, lotId) {
  const range = ensureIngredientOrderSheet().getRange(orderRow.rowNumber, 9, 1, 2); // I:状態 J:入荷在庫ID
  range.setNumberFormat('@');
  range.setValues([['入荷済み', lotId]]);
}

/**
 * 仕込みを記録する
 * 仕込み品の明細を作り、あわせて使った原材料の残量を利用者の入力どおりに更新する
 * 原材料から仕込み品への換算率は仮定しない（残量は利用者が目視で入力する）
 */
function saveIngredientPrep(input) {
  const payload = input || {};
  requireManagerRole_(payload.authToken);
  return runIngredientOperation(payload.operationKey, '仕込み', function () {
    const masters = getIngredientMasters();
    const master = findIngredientMasterById(masters, payload.ingredientId);
    if (master.category !== '仕込み済み') {
      throw new Error(master.name + ' は仕込み品として登録されていません。食材マスタの区分を確認してください');
    }
    const quantity = parseIngredientQuantity(payload.quantity, '仕込んだ数量', false);
    const unit = assertIngredientChoice(payload.unit, INGREDIENT_UNITS, '単位');
    const storageState = assertIngredientChoice(payload.storageState, ['冷蔵', '冷凍'], '保存状態');
    const preparedAtUnknown = (payload.preparedAtUnknown === true || payload.preparedAtUnknown === 'true');
    const preparedAt = normalizeIngredientDateTimeInput(payload.preparedAt, '仕込み日時');
    if (!preparedAtUnknown && preparedAt === '') {
      throw new Error('仕込んだ日時を入力してください（分からない場合は「不明」を選んでください）');
    }
    const useByDate = normalizeIngredientDateInput(payload.useByDate, '使用期限');
    const actor = resolveIngredientActor(payload.actorEmpId);

    const lotSheet = ensureIngredientLotSheet();
    // 使った原材料の入力内容を先に確かめる。書き込んだあとで弾かれて片側だけが残るのを防ぐ
    const consumptions = prepareIngredientConsumptions(lotSheet, payload.consumptions);

    const lot = createIngredientLotBase(master, quantity, unit, storageState, actor);
    lot.preparedAt = preparedAtUnknown ? '' : preparedAt;
    lot.preparedAtUnknown = preparedAtUnknown;
    lot.useByDate = useByDate;
    lot.note = normalizeIngredientText(payload.note);
    if (storageState === '冷凍') {
      lot.frozenKind = '店内冷凍';
      lot.frozenAt = preparedAtUnknown ? '' : preparedAt;
    }
    // 仕込み品を先に '保留' で作り、原材料の残量を更新してから '有効' にする
    const newRowNumber = appendIngredientLot(lotSheet, lot, '保留');
    const consumedTexts = applyIngredientConsumptions(lotSheet, consumptions, actor);
    SpreadsheetApp.flush();
    activateIngredientLot(lotSheet, newRowNumber);
    lot.status = '有効';

    let message = master.name + ' の仕込みを記録しました（' + formatIngredientQuantity(quantity) + unit + '・' +
      (preparedAtUnknown ? '仕込み日時は不明（要確認）' : preparedAt) + '）';
    if (useByDate === '') {
      message = message + '。使用期限が未設定のため要確認です';
    }
    if (consumedTexts.length > 0) {
      message = message + '。使った原材料の残量を更新しました（' + consumedTexts.join('、') + '）';
    }
    return {
      'message': message,
      'lotId': lot.id,
      'ingredientId': master.id,
      'quantity': quantity,
      'unit': unit,
      'actor': actor
    };
  });
}

/**
 * 仕込みで使った原材料の入力内容を、書き込む前にまとめて確かめる
 * @return {Array<Object>} { lot, before, remaining } の配列
 */
function prepareIngredientConsumptions(lotSheet, consumptions) {
  const prepared = [];
  if (!consumptions || consumptions.length === 0) {
    return prepared;
  }
  const usedLotIds = {};
  for (let i = 0; i < consumptions.length; i++) {
    const consumption = consumptions[i] || {};
    const lotId = normalizeIngredientText(consumption.lotId);
    if (lotId === '') {
      continue;
    }
    if (usedLotIds[lotId]) {
      throw new Error('同じ在庫明細が2回指定されています。1件にまとめてください');
    }
    usedLotIds[lotId] = true;

    const lot = findIngredientLotById(lotSheet, lotId);
    assertIngredientLotVersion(lot, consumption.version);
    const remaining = parseIngredientQuantity(consumption.remainingQuantity, lot.name + 'の仕込み後の残量', true);
    if (remaining > lot.quantity) {
      throw new Error(lot.name + 'の仕込み後の残量が元の数量を超えています（元 ' + formatIngredientQuantity(lot.quantity) + lot.unit + '）');
    }
    prepared.push({ 'lot': lot, 'before': lot.quantity, 'remaining': remaining });
  }
  return prepared;
}

/**
 * 確かめ終わった原材料の残量を、利用者が入力した値に更新する
 * @return {Array<string>} 更新内容の説明文
 */
function applyIngredientConsumptions(lotSheet, consumptions, actor) {
  const texts = [];
  const checkedAt = formatIngredientDateTime(new Date());
  for (let i = 0; i < consumptions.length; i++) {
    const lot = consumptions[i].lot;
    lot.quantity = consumptions[i].remaining;
    lot.checkedAt = checkedAt;
    lot.checkedBy = actor;
    writeIngredientLot(lotSheet, lot);
    texts.push(lot.name + ' ' + formatIngredientQuantity(consumptions[i].before) + lot.unit +
      '→' + formatIngredientQuantity(consumptions[i].remaining) + lot.unit);
  }
  return texts;
}

/**
 * 営業中の残量更新
 * その明細の「現在のおおよその残量」だけを更新し、仕込み日時や期限は変更しない
 */
function saveIngredientAmount(input) {
  const payload = input || {};
  requireManagerRole_(payload.authToken);
  return runIngredientOperation(payload.operationKey, '残量更新', function () {
    const lotSheet = ensureIngredientLotSheet();
    const lot = findIngredientLotById(lotSheet, payload.lotId);
    assertIngredientLotVersion(lot, payload.version);

    const quantity = parseIngredientQuantity(payload.quantity, '現在のおおよその残量', true);
    const actor = resolveIngredientActor(payload.actorEmpId);
    const before = lot.quantity;
    lot.quantity = quantity;
    lot.checkedAt = formatIngredientDateTime(new Date());
    lot.checkedBy = actor;
    if (quantity === 0 && (payload.finish === true || payload.finish === 'true')) {
      lot.status = '終了'; // 使い切ったことが確認できた明細だけ一覧から外す
    }
    writeIngredientLot(lotSheet, lot);

    let message = lot.name + ' の残量を更新しました（' + formatIngredientQuantity(before) + lot.unit +
      '→' + formatIngredientQuantity(quantity) + lot.unit + '）';
    if (lot.status === '終了') {
      message = message + '。使い切りとして一覧から外しました';
    }
    return {
      'message': message,
      'lotId': lot.id,
      'ingredientId': lot.ingredientId,
      'quantity': quantity,
      'unit': lot.unit,
      'actor': actor
    };
  });
}

/**
 * 冷凍に移す（全量でも一部でも記録できる）
 * 冷凍した量・実際に冷凍した日時・冷凍後の使用期限を登録し、入荷・仕込み・ラベル期限の情報はそのまま残す
 * 期限は入力された内容だけを使い、「冷凍したから何日延長」といった自動設定はしない
 */
function saveIngredientFreeze(input) {
  const payload = input || {};
  requireManagerRole_(payload.authToken);
  return runIngredientOperation(payload.operationKey, '冷凍', function () {
    const lotSheet = ensureIngredientLotSheet();
    const lot = findIngredientLotById(lotSheet, payload.lotId);
    assertIngredientLotVersion(lot, payload.version);

    if (lot.storageState === '冷凍') {
      throw new Error(lot.name + ' はすでに冷凍です');
    }
    if (lot.storageState === '解凍中') {
      throw new Error(lot.name + ' は解凍中です。解凍の完了を記録してから操作してください');
    }
    const quantity = parseIngredientQuantity(payload.quantity, '冷凍する量', false);
    assertIngredientMovableQuantity(lot, quantity, '冷凍する量');
    const frozenAt = normalizeIngredientDateTimeInput(payload.frozenAt, '冷凍した日時');
    if (frozenAt === '') {
      throw new Error('実際に冷凍した日時を入力してください');
    }
    const useByDate = normalizeIngredientDateInput(payload.useByDate, '冷凍後の使用期限');
    const actor = resolveIngredientActor(payload.actorEmpId);
    // 冷凍前に持っていた使用期限は今の状態の期限ではないので、備考に残したうえで期限欄からは外す
    let keptNote = (lot.useByDate === '') ? '' : ('冷凍前の使用期限: ' + lot.useByDate);
    // 冷凍したというだけで期限切れ・期限不明の警告が消えないよう、冷凍前の状態も備考に残す
    const beforeExpiry = buildIngredientExpiryStatus(lot, new Date());
    if (beforeExpiry.state === '期限切れ') {
      keptNote = appendIngredientNote(keptNote, '冷凍前に期限切れ（' + beforeExpiry.kind + ' ' + beforeExpiry.date + '）');
    } else if (beforeExpiry.state === '要確認') {
      keptNote = appendIngredientNote(keptNote, '冷凍前は期限が未設定・不明');
    }
    const now = formatIngredientDateTime(new Date());

    let partnerLotId = '';
    if (quantity === lot.quantity) {
      lot.storageState = '冷凍';
      lot.frozenAt = frozenAt;
      lot.frozenKind = '店内冷凍';
      lot.useByDate = useByDate;
      lot.note = appendIngredientNote(lot.note, keptNote);
      lot.checkedAt = now;
      lot.checkedBy = actor;
      writeIngredientLot(lotSheet, lot);
    } else {
      const frozenLot = createIngredientLotBase(
        { 'id': lot.ingredientId, 'name': lot.name, 'category': lot.category },
        quantity, lot.unit, '冷凍', actor);
      frozenLot.labelExpiryKind = lot.labelExpiryKind;
      frozenLot.labelExpiryDate = lot.labelExpiryDate;
      frozenLot.useByDate = useByDate;
      frozenLot.receivedDate = lot.receivedDate;
      frozenLot.preparedAt = lot.preparedAt;
      frozenLot.preparedAtUnknown = lot.preparedAtUnknown;
      frozenLot.frozenAt = frozenAt;
      frozenLot.frozenKind = '店内冷凍';
      frozenLot.sourceLotId = lot.id;
      frozenLot.note = appendIngredientNote(lot.note, keptNote);
      partnerLotId = moveIngredientQuantity(lotSheet, lot, quantity, frozenLot);
    }

    let message = lot.name + ' を冷凍にしました（' + formatIngredientQuantity(quantity) + lot.unit + '・冷凍日時 ' + frozenAt + '）';
    if (useByDate === '') {
      message = message + '。冷凍後の使用期限が未設定のため要確認です';
    }
    return {
      'message': message,
      'lotId': lot.id,
      'partnerLotId': partnerLotId,
      'ingredientId': lot.ingredientId,
      'quantity': quantity,
      'unit': lot.unit,
      'actor': actor
    };
  });
}

/**
 * 解凍を開始する（全量でも一部でもよい）
 * 冷凍日時などの履歴は消さずに残し、解凍後の使用期限は入力された内容だけを設定する
 */
function saveIngredientThawStart(input) {
  const payload = input || {};
  requireManagerRole_(payload.authToken);
  return runIngredientOperation(payload.operationKey, '解凍開始', function () {
    const lotSheet = ensureIngredientLotSheet();
    const lot = findIngredientLotById(lotSheet, payload.lotId);
    assertIngredientLotVersion(lot, payload.version);

    if (lot.storageState !== '冷凍') {
      throw new Error(lot.name + ' は冷凍ではありません（現在: ' + lot.storageState + '）');
    }
    const quantity = parseIngredientQuantity(payload.quantity, '解凍する量', false);
    assertIngredientMovableQuantity(lot, quantity, '解凍する量');
    const startedAt = normalizeIngredientDateTimeInput(payload.startedAt, '解凍開始日時');
    if (startedAt === '') {
      throw new Error('解凍を開始した日時を入力してください');
    }
    const useByDate = normalizeIngredientDateInput(payload.useByDate, '解凍後の使用期限');
    const actor = resolveIngredientActor(payload.actorEmpId);
    let keptNote = (lot.useByDate === '') ? '' : ('解凍前の使用期限: ' + lot.useByDate);
    // 解凍したというだけで期限切れ・期限不明の警告が消えないよう、解凍前の状態も備考に残す
    const beforeExpiry = buildIngredientExpiryStatus(lot, new Date());
    if (beforeExpiry.state === '期限切れ') {
      keptNote = appendIngredientNote(keptNote, '解凍前に期限切れ（' + beforeExpiry.kind + ' ' + beforeExpiry.date + '）');
    } else if (beforeExpiry.state === '要確認') {
      keptNote = appendIngredientNote(keptNote, '解凍前は期限が未設定・不明');
    }
    const now = formatIngredientDateTime(new Date());

    let partnerLotId = '';
    if (quantity === lot.quantity) {
      lot.storageState = '解凍中';
      lot.thawStartedAt = startedAt;
      lot.thawCompletedAt = '';
      lot.useByDate = useByDate;
      lot.note = appendIngredientNote(lot.note, keptNote);
      lot.checkedAt = now;
      lot.checkedBy = actor;
      writeIngredientLot(lotSheet, lot);
    } else {
      const thawingLot = createIngredientLotBase(
        { 'id': lot.ingredientId, 'name': lot.name, 'category': lot.category },
        quantity, lot.unit, '解凍中', actor);
      thawingLot.labelExpiryKind = lot.labelExpiryKind;
      thawingLot.labelExpiryDate = lot.labelExpiryDate;
      thawingLot.useByDate = useByDate;
      thawingLot.receivedDate = lot.receivedDate;
      thawingLot.preparedAt = lot.preparedAt;
      thawingLot.preparedAtUnknown = lot.preparedAtUnknown;
      thawingLot.frozenAt = lot.frozenAt;          // 冷凍の履歴は解凍しても残す
      thawingLot.frozenKind = lot.frozenKind;
      thawingLot.thawStartedAt = startedAt;
      thawingLot.sourceLotId = lot.id;
      thawingLot.note = appendIngredientNote(lot.note, keptNote);
      partnerLotId = moveIngredientQuantity(lotSheet, lot, quantity, thawingLot);
    }

    let message = lot.name + ' の解凍を開始しました（' + formatIngredientQuantity(quantity) + lot.unit + '・開始 ' + startedAt + '）';
    if (useByDate === '') {
      message = message + '。解凍後の使用期限が未設定のため要確認です';
    }
    return {
      'message': message,
      'lotId': lot.id,
      'partnerLotId': partnerLotId,
      'ingredientId': lot.ingredientId,
      'quantity': quantity,
      'unit': lot.unit,
      'actor': actor
    };
  });
}

/**
 * 解凍の完了を記録する
 * 保存状態を冷蔵に戻し、冷凍日時・解凍開始日時はそのまま残す
 */
function saveIngredientThawComplete(input) {
  const payload = input || {};
  requireManagerRole_(payload.authToken);
  return runIngredientOperation(payload.operationKey, '解凍完了', function () {
    const lotSheet = ensureIngredientLotSheet();
    const lot = findIngredientLotById(lotSheet, payload.lotId);
    assertIngredientLotVersion(lot, payload.version);

    if (lot.storageState !== '解凍中') {
      throw new Error(lot.name + ' は解凍中ではありません（現在: ' + lot.storageState + '）');
    }
    const completedAt = normalizeIngredientDateTimeInput(payload.completedAt, '解凍完了日時');
    if (completedAt === '') {
      throw new Error('解凍が完了した日時を入力してください');
    }
    if (lot.thawStartedAt !== '' && completedAt < lot.thawStartedAt) {
      throw new Error('解凍完了日時が解凍開始日時（' + lot.thawStartedAt + '）より前になっています');
    }
    const useByDate = normalizeIngredientDateInput(payload.useByDate, '解凍後の使用期限');
    const actor = resolveIngredientActor(payload.actorEmpId);

    lot.storageState = '冷蔵';
    lot.thawCompletedAt = completedAt;
    if (useByDate !== '') {
      lot.useByDate = useByDate; // 未入力なら解凍開始時に登録した使用期限をそのまま使う
    }
    lot.checkedAt = formatIngredientDateTime(new Date());
    lot.checkedBy = actor;
    writeIngredientLot(lotSheet, lot);

    let message = lot.name + ' の解凍完了を記録しました（' + completedAt + '）';
    if (lot.useByDate === '') {
      message = message + '。使用期限が未設定のため要確認です';
    }
    return {
      'message': message,
      'lotId': lot.id,
      'ingredientId': lot.ingredientId,
      'quantity': lot.quantity,
      'unit': lot.unit,
      'actor': actor
    };
  });
}

/**
 * 取り分けた脂を別在庫として登録する
 * 由来が分かる食材名で登録し、あわせて元の肉の残量も入力どおりに更新して二重に数えないようにする
 * 脂の量と肉の残量はどちらも利用者の入力を使い、重量や用途への換算は推測しない
 */
function saveIngredientFat(input) {
  const payload = input || {};
  requireManagerRole_(payload.authToken);
  return runIngredientOperation(payload.operationKey, '脂の取り分け', function () {
    const lotSheet = ensureIngredientLotSheet();
    const sourceLot = findIngredientLotById(lotSheet, payload.lotId);
    assertIngredientLotVersion(sourceLot, payload.version);

    const fatQuantity = parseIngredientQuantity(payload.fatQuantity, '脂の量', false);
    const fatUnit = assertIngredientChoice(payload.fatUnit, INGREDIENT_UNITS, '単位');
    const remaining = parseIngredientQuantity(payload.sourceRemainingQuantity, sourceLot.name + 'の残量', true);
    if (remaining > sourceLot.quantity) {
      throw new Error(sourceLot.name + 'の残量が元の数量を超えています（元 ' + formatIngredientQuantity(sourceLot.quantity) + sourceLot.unit + '）');
    }
    const frozenAt = normalizeIngredientDateTimeInput(payload.frozenAt, '冷凍した日時');
    if (frozenAt === '') {
      throw new Error('脂を冷凍した日時を入力してください');
    }
    const useByDate = normalizeIngredientDateInput(payload.useByDate, '冷凍後の使用期限');
    const actor = resolveIngredientActor(payload.actorEmpId);

    const fatMaster = ensureIngredientFatMaster(sourceLot, fatUnit);
    const fatLot = createIngredientLotBase(fatMaster, fatQuantity, fatUnit, '冷凍', actor);
    fatLot.useByDate = useByDate;
    fatLot.frozenAt = frozenAt;
    fatLot.frozenKind = '店内冷凍';
    fatLot.sourceLotId = sourceLot.id;
    fatLot.note = appendIngredientNote(normalizeIngredientText(payload.note), sourceLot.name + 'から取り分け');

    // 脂は元の肉から取り分けたものなので、肉の残量は利用者が入力した値へ置き換える（自動で差し引かない）
    const newRowNumber = appendIngredientLot(lotSheet, fatLot, '保留');
    const before = sourceLot.quantity;
    sourceLot.quantity = remaining;
    sourceLot.checkedAt = formatIngredientDateTime(new Date());
    sourceLot.checkedBy = actor;
    writeIngredientLot(lotSheet, sourceLot);
    SpreadsheetApp.flush();
    activateIngredientLot(lotSheet, newRowNumber);

    let message = fatMaster.name + ' を冷凍在庫として登録しました（' + formatIngredientQuantity(fatQuantity) + fatUnit + '）。' +
      sourceLot.name + ' の残量を ' + formatIngredientQuantity(before) + sourceLot.unit + '→' +
      formatIngredientQuantity(remaining) + sourceLot.unit + ' に更新しました';
    if (useByDate === '') {
      message = message + '。脂の使用期限が未設定のため要確認です';
    }
    return {
      'message': message,
      'lotId': fatLot.id,
      'partnerLotId': sourceLot.id,
      'ingredientId': fatMaster.id,
      'quantity': fatQuantity,
      'unit': fatUnit,
      'actor': actor
    };
  });
}

/**
 * 「脂（○○由来）」の食材マスタを用意する（無ければ作る）
 * @return {Object} 食材マスタの1件
 */
function ensureIngredientFatMaster(sourceLot, fatUnit) {
  const name = '脂（' + sourceLot.name + '由来）';
  const masters = getIngredientMasters();
  for (let i = 0; i < masters.length; i++) {
    if (masters[i].name === name) {
      return masters[i];
    }
  }
  const category = (INGREDIENT_CATEGORIES.indexOf(sourceLot.category) === -1) ? INGREDIENT_CATEGORIES[0] : sourceLot.category;
  const master = {
    'id': newIngredientId('ING'),
    'name': name,
    'category': category,
    'defaultUnit': fatUnit,
    'minQuantity': null,
    'minUnit': '',
    'notifyExpirySoon': false,
    'supplierId': '',
    'active': true
  };
  const sheet = ensureIngredientMasterSheet();
  appendIngredientRows(sheet, [[
    master.id, master.name, master.category, master.defaultUnit,
    '', '', '', '', '○', formatIngredientDateTime(new Date())
  ]], INGREDIENT_MASTER_SHEET_HEADER.length);
  return master;
}

/**
 * 明細の使用期限を設定・変更する（詳細画面から使う）
 * 数量・仕込み日時・冷凍日時には触れない。空で保存すると「要確認」に戻る
 */
function saveIngredientUseBy(input) {
  const payload = input || {};
  requireManagerRole_(payload.authToken);
  return runIngredientOperation(payload.operationKey, '使用期限の設定', function () {
    const lotSheet = ensureIngredientLotSheet();
    const lot = findIngredientLotById(lotSheet, payload.lotId);
    assertIngredientLotVersion(lot, payload.version);

    const useByDate = normalizeIngredientDateInput(payload.useByDate, '使用期限');
    const actor = resolveIngredientActor(payload.actorEmpId);
    lot.useByDate = useByDate;
    lot.note = normalizeIngredientText(payload.note);
    lot.checkedAt = formatIngredientDateTime(new Date());
    lot.checkedBy = actor;
    writeIngredientLot(lotSheet, lot);

    return {
      'message': lot.name + ' の使用期限を' + (useByDate === '' ? '未設定（要確認）にしました' : useByDate + ' に設定しました'),
      'lotId': lot.id,
      'ingredientId': lot.ingredientId,
      'unit': lot.unit,
      'actor': actor
    };
  });
}

/**
 * 食材マスタ（食材名・区分・既定単位・最低残量・業者）を登録・更新する
 * 食材や仕込み品はここから増やせる
 */
function saveIngredientMaster(input) {
  const payload = input || {};
  requireManagerRole_(payload.authToken);
  return runIngredientOperation(payload.operationKey, '食材設定', function () {
    const sheet = ensureIngredientMasterSheet();
    const rows = readIngredientRows(sheet, INGREDIENT_MASTER_SHEET_HEADER);
    const ingredientId = normalizeIngredientText(payload.ingredientId);
    const name = normalizeIngredientText(payload.name);
    if (name === '') {
      throw new Error('食材名を入力してください');
    }
    const category = assertIngredientChoice(payload.category, INGREDIENT_CATEGORIES, '区分');
    const defaultUnit = assertIngredientChoice(payload.defaultUnit, INGREDIENT_UNITS, '既定単位');

    const minQuantityText = normalizeIngredientText(payload.minQuantity);
    let minQuantity = '';
    let minUnit = '';
    if (minQuantityText !== '') {
      minQuantity = formatIngredientQuantity(parseIngredientQuantity(minQuantityText, '最低残量', true));
      minUnit = assertIngredientChoice(payload.minUnit, INGREDIENT_UNITS, '最低残量の単位');
    }
    const supplierId = normalizeIngredientText(payload.supplierId);
    if (supplierId !== '' && !findIngredientSupplierRow(readIngredientRows(ensureIngredientSupplierSheet(), INGREDIENT_SUPPLIER_SHEET_HEADER), supplierId)) {
      throw new Error('登録されていない業者です: ' + supplierId);
    }

    let targetRow = -1;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i].values['食材名'] === name && rows[i].values['食材ID'] !== ingredientId) {
        throw new Error('同じ名前の食材がすでに登録されています: ' + name);
      }
      if (ingredientId !== '' && rows[i].values['食材ID'] === ingredientId) {
        targetRow = rows[i].rowNumber;
      }
    }
    if (ingredientId !== '' && targetRow === -1) {
      throw new Error('対象の食材が見つかりません: ' + ingredientId);
    }

    const values = [
      (ingredientId === '' ? newIngredientId('ING') : ingredientId), name, category, defaultUnit,
      minQuantity, minUnit,
      (payload.notifyExpirySoon === true || payload.notifyExpirySoon === 'true') ? '○' : '',
      supplierId,
      (payload.active === false || payload.active === 'false') ? '' : '○',
      formatIngredientDateTime(new Date())
    ];
    if (targetRow === -1) {
      appendIngredientRows(sheet, [values], INGREDIENT_MASTER_SHEET_HEADER.length);
    } else {
      const range = sheet.getRange(targetRow, 1, 1, INGREDIENT_MASTER_SHEET_HEADER.length);
      range.setNumberFormat('@');
      range.setValues([values]);
    }
    return {
      'message': name + ' の設定を保存しました',
      'ingredientId': values[0],
      'actor': resolveIngredientActor(payload.actorEmpId)
    };
  });
}

/**
 * 業者IDから業者の行を探す
 */
function findIngredientSupplierRow(rows, supplierId) {
  for (let i = 0; i < rows.length; i++) {
    if (rows[i].values['業者ID'] === supplierId) {
      return rows[i];
    }
  }
  return null;
}

/**
 * 曜日の一覧（画面から来た配列）を検証して「月,火」の形式の保存用文字列にする
 */
function normalizeIngredientWeekdayInput(value, label) {
  if (!value) {
    return '';
  }
  const list = Array.isArray(value) ? value : String(value).split(/[,、\s]+/);
  const labels = [];
  for (let i = 0; i < list.length; i++) {
    const day = normalizeIngredientText(list[i]);
    if (day === '') {
      continue;
    }
    if (INGREDIENT_WEEKDAY_LABELS.indexOf(day) === -1) {
      throw new Error('不正な' + label + 'です: ' + day);
    }
    if (labels.indexOf(day) === -1) {
      labels.push(day);
    }
  }
  return labels.join(',');
}

/**
 * 業者と発注条件（受注締切・営業日・リードタイム・臨時休業日・受取可能日）を登録・更新する
 * 未入力の条件はそのまま未設定として保存し、推測で埋めない
 */
function saveIngredientSupplier(input) {
  const payload = input || {};
  requireManagerRole_(payload.authToken);
  return runIngredientOperation(payload.operationKey, '業者設定', function () {
    const sheet = ensureIngredientSupplierSheet();
    const rows = readIngredientRows(sheet, INGREDIENT_SUPPLIER_SHEET_HEADER);
    const supplierId = normalizeIngredientText(payload.supplierId);
    const name = normalizeIngredientText(payload.name);
    if (name === '') {
      throw new Error('業者名を入力してください');
    }

    const deadlineTime = normalizeIngredientText(payload.deadlineTime);
    if (deadlineTime !== '' && !INGREDIENT_TIME_PATTERN.test(deadlineTime)) {
      throw new Error('受注締切時刻は HH:mm の形式で入力してください: ' + deadlineTime);
    }
    const leadTimeText = normalizeIngredientText(payload.leadTimeDays);
    if (leadTimeText !== '' && !/^\d+$/.test(leadTimeText)) {
      throw new Error('納品までの日数は0以上の整数で入力してください: ' + leadTimeText);
    }
    const leadTimeCount = normalizeIngredientText(payload.leadTimeCount);
    if (leadTimeCount !== '' && INGREDIENT_LEAD_TIME_COUNTS.indexOf(leadTimeCount) === -1) {
      throw new Error('不正な日数の数え方です: ' + leadTimeCount);
    }

    const closedDates = [];
    const closedList = normalizeIngredientText(payload.closedDates) === '' ? [] : String(payload.closedDates).split(/[,、\s\r\n]+/);
    for (let i = 0; i < closedList.length; i++) {
      if (normalizeIngredientText(closedList[i]) === '') {
        continue;
      }
      closedDates.push(normalizeIngredientDateInput(closedList[i], '臨時休業日'));
    }

    let targetRow = -1;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i].values['業者名'] === name && rows[i].values['業者ID'] !== supplierId) {
        throw new Error('同じ名前の業者がすでに登録されています: ' + name);
      }
      if (supplierId !== '' && rows[i].values['業者ID'] === supplierId) {
        targetRow = rows[i].rowNumber;
      }
    }
    if (supplierId !== '' && targetRow === -1) {
      throw new Error('対象の業者が見つかりません: ' + supplierId);
    }

    const values = [
      (supplierId === '' ? newIngredientId('SUP') : supplierId), name, deadlineTime,
      normalizeIngredientWeekdayInput(payload.orderDays, '受注営業日'),
      normalizeIngredientWeekdayInput(payload.shipDays, '出荷営業日'),
      normalizeIngredientWeekdayInput(payload.deliveryDays, '納品営業日'),
      leadTimeText, leadTimeCount, closedDates.join(','),
      normalizeIngredientWeekdayInput(payload.pickupDays, '店舗受取可能日'),
      normalizeIngredientText(payload.note), formatIngredientDateTime(new Date())
    ];
    if (targetRow === -1) {
      appendIngredientRows(sheet, [values], INGREDIENT_SUPPLIER_SHEET_HEADER.length);
    } else {
      const range = sheet.getRange(targetRow, 1, 1, INGREDIENT_SUPPLIER_SHEET_HEADER.length);
      range.setNumberFormat('@');
      range.setValues([values]);
    }
    return {
      'message': name + ' の発注条件を保存しました',
      'actor': resolveIngredientActor(payload.actorEmpId)
    };
  });
}

/**
 * 発注済みの内容（数量・入荷予定日）を記録する
 * 実際の注文送信は行わない。現在庫も増やさず、入荷を記録した時点で在庫に反映する
 */
function saveIngredientOrder(input) {
  const payload = input || {};
  requireManagerRole_(payload.authToken);
  return runIngredientOperation(payload.operationKey, '発注記録', function () {
    const masters = getIngredientMasters();
    const master = findIngredientMasterById(masters, payload.ingredientId);
    const quantity = parseIngredientQuantity(payload.quantity, '発注数量', false);
    const unit = assertIngredientChoice(payload.unit, INGREDIENT_UNITS, '単位');
    const orderedAt = normalizeIngredientDateTimeInput(payload.orderedAt, '発注日時');
    if (orderedAt === '') {
      throw new Error('発注した日時を入力してください');
    }
    const expectedDate = normalizeIngredientDateInput(payload.expectedDate, '入荷予定日');
    const supplierId = normalizeIngredientText(payload.supplierId);
    if (supplierId !== '' && !findIngredientSupplierRow(readIngredientRows(ensureIngredientSupplierSheet(), INGREDIENT_SUPPLIER_SHEET_HEADER), supplierId)) {
      throw new Error('登録されていない業者です: ' + supplierId);
    }

    const sheet = ensureIngredientOrderSheet();
    const orderId = newIngredientId('ORD');
    appendIngredientRows(sheet, [[
      orderId, master.id, master.name, supplierId,
      formatIngredientQuantity(quantity), unit, orderedAt, expectedDate,
      '発注済み', '', normalizeIngredientText(payload.note), formatIngredientDateTime(new Date())
    ]], INGREDIENT_ORDER_SHEET_HEADER.length);

    return {
      'message': master.name + ' の発注を記録しました（' + formatIngredientQuantity(quantity) + unit +
        '・入荷予定 ' + (expectedDate === '' ? '未設定' : expectedDate) + '）。入荷を記録するまで在庫には加えません',
      'ingredientId': master.id,
      'quantity': quantity,
      'unit': unit,
      'actor': resolveIngredientActor(payload.actorEmpId)
    };
  });
}

/**
 * 発注の記録を取り消す（在庫は動かさない）
 */
function cancelIngredientOrder(input) {
  const payload = input || {};
  requireManagerRole_(payload.authToken);
  return runIngredientOperation(payload.operationKey, '発注取消', function () {
    const sheet = ensureIngredientOrderSheet();
    const rows = readIngredientRows(sheet, INGREDIENT_ORDER_SHEET_HEADER);
    const orderId = normalizeIngredientText(payload.orderId);
    for (let i = 0; i < rows.length; i++) {
      if (rows[i].values['発注ID'] !== orderId) {
        continue;
      }
      if (rows[i].values['状態'] !== '発注済み') {
        throw new Error('この発注はすでに「' + rows[i].values['状態'] + '」になっています');
      }
      const range = sheet.getRange(rows[i].rowNumber, 9, 1, 1);
      range.setNumberFormat('@');
      range.setValue('取消');
      return {
        'message': rows[i].values['食材名'] + ' の発注記録を取り消しました',
        'ingredientId': rows[i].values['食材ID'],
        'actor': resolveIngredientActor(payload.actorEmpId)
      };
    }
    throw new Error('対象の発注が見つかりません: ' + orderId);
  });
}
