const INGREDIENT_DATE_PATTERN = /^\d{4}\/\d{2}\/\d{2}$/;
const INGREDIENT_DATE_TIME_PATTERN = /^\d{4}\/\d{2}\/\d{2} \d{2}:\d{2}$/;
const INGREDIENT_TIME_PATTERN = /^([01][0-9]|2[0-3]):[0-5][0-9]$/;

const INGREDIENT_MAX_QUANTITY = 100000;

const INGREDIENT_PLAN_SEARCH_DAYS = 120;

function ensureIngredientSheet(sheetName, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  let isNewSheet = false;
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    isNewSheet = true;
  }

  const shortageColumns = header.length - sheet.getMaxColumns();
  if (shortageColumns > 0) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), shortageColumns);
  }

  const headerRange = sheet.getRange(1, 1, 1, header.length);
  const currentHeader = headerRange.getValues()[0];
  for (let i = 0; i < header.length; i++) {
    if (currentHeader[i] !== header[i]) {
      headerRange.setValues([header]);
      break;
    }
  }

  const dataRowCount = sheet.getMaxRows() - 1;
  if (isNewSheet && dataRowCount > 0) {
    sheet.getRange(2, 1, dataRowCount, header.length).setNumberFormat('@');
  }
  return sheet;
}

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
      ''
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
      continue;
    }
    rows.push({ 'rowNumber': i + 2, 'values': record });
  }
  return rows;
}

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

function newIngredientId(prefix) {
  return prefix + '-' + Utilities.getUuid().replace(/-/g, '').substring(0, 12);
}

function normalizeIngredientText(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

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
    return null;
  }
  return date;
}

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

function roundIngredientQuantity(quantity) {
  return Math.round(quantity * 1000) / 1000;
}

function formatIngredientQuantity(quantity) {
  return String(roundIngredientQuantity(quantity));
}

function addIngredientDays(date, days) {
  const result = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  result.setDate(result.getDate() + days);
  return result;
}

function formatIngredientDateWithWeekday(date) {
  return formatIngredientDate(date) + '（' + INGREDIENT_WEEKDAY_LABELS[date.getDay()] + '）';
}

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

      'orderUnit': values['発注単位'] !== '' ? values['発注単位'] : values['既定単位'],
      'rowNumber': rows[i].rowNumber
    });
  }
  return masters;
}

function normalizeIngredientSupplierNote_(value) {
  const retiredInstruction = /商品コード|product\s*code|item\s*code/i;
  return normalizeIngredientText(value).split(/\r\n|\r|\n/).map(function (line) {
    if (!retiredInstruction.test(line)) {
      return line;
    }
    const parts = line.match(/[^。；;、,]+[。；;、,]?/g) || [];
    return parts.filter(function (part) {
      return !retiredInstruction.test(part);
    }).join('').replace(/[；;、,]+\s*$/, '').trim();
  }).filter(function (line) { return line.trim() !== ''; }).join('\n');
}

function normalizeOrderSupplierName_(name) {
  if (name === '○○青果') { return 'freee青果'; }
  if (name === '△△食品') { return 'freee食品'; }
  return name;
}

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
      'name': normalizeOrderSupplierName_(values['業者名']),
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

      'orderNo': values['発注番号'],
      'orderedBy': values['発注者'],
      'mailTo': values['送信先メール'],
      'rowNumber': rows[i].rowNumber
    });
  }
  return orders;
}

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

function formatIngredientWeekdayList(days) {
  const labels = [];
  for (let i = 0; i < days.length; i++) {
    labels.push(INGREDIENT_WEEKDAY_LABELS[days[i]]);
  }
  return labels.join('・');
}

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

function isIngredientClosedDate(supplier, date) {
  return supplier.closedDates.indexOf(formatIngredientDate(date)) !== -1;
}

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
      continue;
    }
    return candidate;
  }
  return null;
}

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
      continue;
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
      'deliveryDateKey': formatIngredientDate(deliveryDate),
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

function findIngredientEventByKey(eventSheet, operationKey) {
  const rows = readIngredientRows(eventSheet, INGREDIENT_EVENT_SHEET_HEADER);
  for (let i = rows.length - 1; i >= 0; i--) {
    if (rows[i].values['操作キー'] === operationKey) {
      return rows[i];
    }
  }
  return null;
}

function startIngredientEvent(eventSheet, operationKey, type) {
  const row = [
    newIngredientId('EVT'), operationKey, formatIngredientDateTime(new Date()), type,
    '', '', '', '', '', '', '', '処理中'
  ];
  return appendIngredientRows(eventSheet, [row], INGREDIENT_EVENT_SHEET_HEADER.length);
}

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

function failIngredientEvent(eventSheet, rowNumber, error) {
  const range = eventSheet.getRange(rowNumber, 10, 1, 3);
  range.setNumberFormat('@');
  range.setValues([['保存できませんでした: ' + error.message, '', '失敗']]);
}

function findIngredientMasterById(masters, ingredientId) {
  const id = normalizeIngredientText(ingredientId);
  for (let i = 0; i < masters.length; i++) {
    if (masters[i].id === id) {
      return masters[i];
    }
  }
  throw new Error('食材マスタに存在しない食材です: ' + id);
}
