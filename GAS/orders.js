const ORDER_HISTORY_LIMIT = 100;

/**
 * 店長権限を確認し、発注管理画面の情報を取得する。
 * authToken - string
 * 戻り値 - {today: string, todayInput: string, nowText: string, storeName: string, testRecipient: string, maxQuantity: number, vendors: ({id: string, name: string, address: string, phone: string, email: string, businessHours: string, closedDays: string, deadlineText: string, leadTimeText: string, planText: string})[], items: ({id: string, name: string, unit: string, vendorId: string, vendorName: unknown})[], history: ({orderNo: string, orderedAt: string, orderedAtLabel: string, vendorId: string, vendorName: string, orderedBy: string, mailTo: string, channelText: string, expectedDate: string, note: string, statusTexts: string[], lines: ({name: string, quantityText: string, status: string})[]})[], upcoming: ({id: string, name: string, orderedAt: string, quantityText: string, expectedDate: string, vendorId: string, vendorName: unknown, status: string})[]}
 */
function getOrderBoard(authToken) {
  requireManagerRole_(authToken);
  ensureOrderMasterData_();
  return buildOrderBoard_();
}

/**
 * 発注管理画面の業者、食材、履歴、入荷予定を組み立てる。
 * 戻り値 - {today: string, todayInput: string, nowText: string, storeName: string, testRecipient: string, maxQuantity: number, vendors: ({id: string, name: string, address: string, phone: string, email: string, businessHours: string, closedDays: string, deadlineText: string, leadTimeText: string, planText: string})[], items: ({id: string, name: string, unit: string, vendorId: string, vendorName: unknown})[], history: ({orderNo: string, orderedAt: string, orderedAtLabel: string, vendorId: string, vendorName: string, orderedBy: string, mailTo: string, channelText: string, expectedDate: string, note: string, statusTexts: string[], lines: ({name: string, quantityText: string, status: string})[]})[], upcoming: ({id: string, name: string, orderedAt: string, quantityText: string, expectedDate: string, vendorId: string, vendorName: unknown, status: string})[]}
 */
function buildOrderBoard_() {
  const now = new Date();
  const suppliers = getIngredientSuppliers();
  const masters = getIngredientMasters();
  const orders = getIngredientOrders();

  const supplierById = {};
  const vendors = [];
  for (let i = 0; i < suppliers.length; i++) {
    supplierById[suppliers[i].id] = suppliers[i];
    vendors.push(buildOrderVendorView_(suppliers[i], now));
  }

  const items = [];
  for (let i = 0; i < masters.length; i++) {
    const master = masters[i];
    if (!master.active) {
      continue;
    }
    if (master.category !== ORDER_DEFAULT_ITEM_CATEGORY) {
      continue;
    }
    const supplier = supplierById[master.supplierId];
    items.push({
      'id': master.id,
      'name': master.name,
      'unit': master.orderUnit,
      'vendorId': master.supplierId,
      'vendorName': supplier ? supplier.name : ''
    });
  }
  items.sort(function (a, b) {
    if (a.vendorName !== b.vendorName) {
      return a.vendorName < b.vendorName ? -1 : 1;
    }
    return a.name < b.name ? -1 : (a.name > b.name ? 1 : 0);
  });

  return {
    'today': formatIngredientDate(now),
    'todayInput': Utilities.formatDate(now, TIME_ZONE, 'yyyy-MM-dd'),
    'nowText': formatIngredientDateTime(now),
    'storeName': ORDER_STORE_NAME,
    'testRecipient': ORDER_TEST_RECIPIENT_EMAIL,
    'maxQuantity': ORDER_MAX_QUANTITY,
    'vendors': vendors,
    'items': items,
    'history': buildOrderHistory_(orders, supplierById),
    'upcoming': buildOrderUpcoming_(orders, supplierById, now)
  };
}

/**
 * 業者の連絡先と発注条件を表示用に整える。
 * supplier - {id: string, name: string, deadlineTime: string, orderDays: number[], shipDays: number[], deliveryDays: number[], leadTimeDays: null または number, leadTimeCount: string, closedDates: string[], pickupDays: number[], note: string, address: string, phone: string, email: string, businessHours: string, closedDays: string, mainItems: string, orderNote: string, rowNumber: number}
 * now - Date
 * 戻り値 - {id: string, name: string, address: string, phone: string, email: string, businessHours: string, closedDays: string, deadlineText: string, leadTimeText: string, planText: string}
 */
function buildOrderVendorView_(supplier, now) {
  return {
    'id': supplier.id,
    'name': supplier.name,
    'address': supplier.address,
    'phone': supplier.phone,
    'email': supplier.email,
    'businessHours': supplier.businessHours,
    'closedDays': supplier.closedDays,

    'deadlineText': buildOrderDeadlineText_(supplier),
    'leadTimeText': buildOrderLeadTimeText_(supplier),
    'planText': buildIngredientSupplierPlan(supplier, now).text
  };
}

/**
 * 発注締切の表示文を作る（未設定のときは推測しない）。
 * supplier - {id: string, name: string, deadlineTime: string, orderDays: number[], shipDays: number[], deliveryDays: number[], leadTimeDays: null または number, leadTimeCount: string, closedDates: string[], pickupDays: number[], note: string, address: string, phone: string, email: string, businessHours: string, closedDays: string, mainItems: string, orderNote: string, rowNumber: number}
 * 戻り値 - string
 */
function buildOrderDeadlineText_(supplier) {
  if (supplier.deadlineTime === '') {
    return '';
  }
  if (supplier.orderDays.length === 0) {
    return supplier.deadlineTime + 'まで';
  }
  return formatIngredientWeekdayList(supplier.orderDays) + 'の' + supplier.deadlineTime + 'まで';
}

/**
 * 通常納品日数の表示文を作る。
 * supplier - {id: string, name: string, deadlineTime: string, orderDays: number[], shipDays: number[], deliveryDays: number[], leadTimeDays: null または number, leadTimeCount: string, closedDates: string[], pickupDays: number[], note: string, address: string, phone: string, email: string, businessHours: string, closedDays: string, mainItems: string, orderNote: string, rowNumber: number}
 * 戻り値 - string
 */
function buildOrderLeadTimeText_(supplier) {
  if (supplier.leadTimeDays === null) {
    return '';
  }
  const countText = supplier.leadTimeCount === '' ? '営業日' : supplier.leadTimeCount;
  return supplier.leadTimeDays + countText;
}

/**
 * 実際にメールを送る宛先を求める 動作確認用の宛先が設定されている間は、業者マスタのアドレスではなくそちらへ送る。
 * supplier - unknown
 * 戻り値 - unknown
 */
function resolveOrderRecipient_(supplier) {
  if (normalizeIngredientText(ORDER_TEST_RECIPIENT_EMAIL) !== '') {
    return ORDER_TEST_RECIPIENT_EMAIL;
  }
  return supplier.email;
}

/**
 * 発注記録を業者別のメール単位にまとめ、新しい順に並べる。
 * orders - ({id: string, ingredientId: string, name: string, supplierId: string, quantity: number, unit: string, orderedAt: string, expectedDate: string, status: string, note: string, orderNo: string, orderedBy: string, mailTo: string, rowNumber: number})[]
 * supplierById - unknown
 * 戻り値 - ({orderNo: string, orderedAt: string, orderedAtLabel: string, vendorId: string, vendorName: string, orderedBy: string, mailTo: string, channelText: string, expectedDate: string, note: string, statusTexts: string[], lines: ({name: string, quantityText: string, status: string})[]})[]
 */
function buildOrderHistory_(orders, supplierById) {
  const groups = [];
  const groupByKey = {};

  for (let i = 0; i < orders.length; i++) {
    const order = orders[i];
    const key = order.orderNo !== '' ? order.orderNo : ('単票:' + order.id);
    if (!groupByKey[key]) {
      const supplier = supplierById[order.supplierId];
      groupByKey[key] = {
        'orderNo': order.orderNo,
        'orderedAt': order.orderedAt,
        'orderedAtLabel': formatOrderHistoryLabel_(order.orderedAt),
        'vendorId': order.supplierId,
        'vendorName': supplier ? supplier.name : (order.supplierId === '' ? '業者未設定' : order.supplierId),
        'orderedBy': order.orderedBy,
        'mailTo': order.mailTo,
        'channelText': order.mailTo !== '' ? 'メール送信' : '手元で記録',
        'expectedDate': order.expectedDate,
        'note': order.note,
        'statusTexts': [],
        'lines': []
      };
      groups.push(groupByKey[key]);
    }
    const group = groupByKey[key];
    group.lines.push({
      'name': order.name,
      'quantityText': formatIngredientQuantity(order.quantity) + order.unit,
      'status': order.status
    });
    if (group.statusTexts.indexOf(order.status) === -1) {
      group.statusTexts.push(order.status);
    }
  }

  groups.sort(function (a, b) {
    const timeA = parseIngredientDateTime(a.orderedAt);
    const timeB = parseIngredientDateTime(b.orderedAt);
    const valueA = timeA ? timeA.getTime() : 0;
    const valueB = timeB ? timeB.getTime() : 0;
    return valueB - valueA;
  });

  return groups.slice(0, ORDER_HISTORY_LIMIT);
}

/**
 * 履歴を削除せず、当日以降の入荷予定を食材ごとに取り出す。
 * orders - ({id: string, ingredientId: string, name: string, supplierId: string, quantity: number, unit: string, orderedAt: string, expectedDate: string, status: string, note: string, orderNo: string, orderedBy: string, mailTo: string, rowNumber: number})[]
 * supplierById - unknown
 * now - Date
 * 戻り値 - ({id: string, name: string, orderedAt: string, quantityText: string, expectedDate: string, vendorId: string, vendorName: unknown, status: string})[]
 */
function buildOrderUpcoming_(orders, supplierById, now) {
  const today = formatIngredientDate(now);
  return orders.map(function (order) {
    const supplier = supplierById[order.supplierId];
    const expectedDate = order.expectedDate || resolveOrderExpectedDate_(supplier, parseIngredientDateTime(order.orderedAt));
    const parsed = parseIngredientDate(expectedDate.replace(/-/g, '/'));
    return {
      id: order.id, name: order.name, orderedAt: order.orderedAt,
      quantityText: formatIngredientQuantity(order.quantity) + order.unit,
      expectedDate: parsed ? formatIngredientDate(parsed) : '',
      vendorId: order.supplierId,
      vendorName: supplier ? supplier.name : (order.supplierId || '業者未設定'),
      status: order.status
    };
  }).filter(function (order) {
    return (!order.expectedDate || order.expectedDate >= today) &&
      ['入荷済み', '入荷済', 'キャンセル', '取消', '取消済み'].indexOf(order.status) === -1;
  }).sort(function (a, b) {
    return (a.expectedDate || '9999').localeCompare(b.expectedDate || '9999') || b.orderedAt.localeCompare(a.orderedAt);
  });
}

/**
 * 発注日時を一覧表示用の短い文字列に変換する。
 * orderedAt - string
 * 戻り値 - string
 */
function formatOrderHistoryLabel_(orderedAt) {
  const date = parseIngredientDateTime(orderedAt);
  if (!date) {
    return orderedAt;
  }
  return Utilities.formatDate(date, TIME_ZONE, 'M/d HH:mm');
}

/**
 * 画面から届いた発注内容を確かめ、食材マスタの内容で組み立て直す 食材名・業者・単位はマスタの値を使い、画面から届いた値は信用しない。
 * items - ({ingredientId: string, quantity: number})[]
 * 戻り値 - ({master: {id: string, name: string, category: string, defaultUnit: string, supplierId: string, active: boolean, orderUnit: string, rowNumber: number}, supplier: unknown, quantity: number, unit: string})[]
 */
function normalizeOrderItems_(items) {
  const list = Array.isArray(items) ? items : [];
  if (list.length === 0) {
    throw new Error('発注する食材が選ばれていません');
  }

  const masters = getIngredientMasters();
  const suppliers = getIngredientSuppliers();
  const supplierById = {};
  for (let i = 0; i < suppliers.length; i++) {
    supplierById[suppliers[i].id] = suppliers[i];
  }

  const normalized = [];
  const usedIds = {};
  for (let i = 0; i < list.length; i++) {
    const ingredientId = normalizeIngredientText(list[i] ? list[i].ingredientId : '');
    if (usedIds[ingredientId]) {
      throw new Error('同じ食材が重複しています。画面を読み込み直してください');
    }
    usedIds[ingredientId] = true;

    const master = findIngredientMasterById(masters, ingredientId);
    if (!master.active) {
      throw new Error('いまは使わない設定の食材です: ' + master.name);
    }
    const supplier = supplierById[master.supplierId];
    if (!supplier) {
      throw new Error(master.name + ' に業者が設定されていません。食材マスタの業者IDを確認してください');
    }
    const quantity = parseIngredientQuantity(list[i].quantity, master.name + 'の数量', false);
    if (quantity > ORDER_MAX_QUANTITY) {
      throw new Error(master.name + 'の数量が大きすぎます（' + ORDER_MAX_QUANTITY + 'まで）');
    }
    if (master.orderUnit === '') {
      throw new Error(master.name + ' の発注単位が設定されていません。食材マスタの発注単位を確認してください');
    }
    normalized.push({
      'master': master,
      'supplier': supplier,
      'quantity': quantity,
      'unit': master.orderUnit
    });
  }
  return normalized;
}

/**
 * 発注内容から業者別のメール下書きを生成する。
 * items - ({ingredientId: string, quantity: number})[]
 * now - Date
 * 戻り値 - ({vendorId: string, vendorName: string, to: string, vendorEmail: string, lines: ({ingredientId: string, name: string, quantity: number, unit: string, quantityText: string})[], subject: string, body: string, canSend: boolean, sendBlockReason: string})[]
 */
function buildOrderDrafts_(items, now) {
  const normalized = normalizeOrderItems_(items);

  const drafts = [];
  const draftByVendorId = {};
  for (let i = 0; i < normalized.length; i++) {
    const line = normalized[i];
    const vendorId = line.supplier.id;
    if (!draftByVendorId[vendorId]) {
      draftByVendorId[vendorId] = {
        'vendorId': vendorId,
        'vendorName': line.supplier.name,
        'to': resolveOrderRecipient_(line.supplier),
        'vendorEmail': line.supplier.email,
        'lines': []
      };
      drafts.push(draftByVendorId[vendorId]);
    }
    draftByVendorId[vendorId].lines.push({
      'ingredientId': line.master.id,
      'name': line.master.name,
      'quantity': line.quantity,
      'unit': line.unit,
      'quantityText': formatIngredientQuantity(line.quantity) + line.unit
    });
  }

  for (let i = 0; i < drafts.length; i++) {
    const draft = drafts[i];
    draft.subject = buildOrderMailSubject_(now);
    draft.body = buildOrderMailBody_(draft);
    draft.canSend = (draft.to !== '');
    draft.sendBlockReason = draft.canSend ? '' : (draft.vendorName + ' のメールアドレスが未登録です');
  }
  return drafts;
}

/**
 * 件名を作る（例：9月16日 発注依頼）。
 * now - Date
 * 戻り値 - string
 */
function buildOrderMailSubject_(now) {
  return (now.getMonth() + 1) + '月' + now.getDate() + '日 発注依頼';
}

/**
 * 本文を作る 業者名・商品・数量は登録された内容から組み立て、利用者に文章を入力させない。
 * draft - {vendorId: string, vendorName: string, to: string, vendorEmail: string, lines: {ingredientId: string, name: string, quantity: number, unit: string, quantityText: string}[], subject: string, body: string, canSend: boolean, sendBlockReason: string}
 * 戻り値 - string
 */
function buildOrderMailBody_(draft) {
  const parts = [];
  parts.push(draft.vendorName + ' 御中');
  parts.push('');
  parts.push('いつもお世話になっております。' + ORDER_STORE_NAME + 'です。');
  parts.push('以下の商品をお願いいたします。');
  parts.push('');

  for (let i = 0; i < draft.lines.length; i++) {
    const line = draft.lines[i];
    parts.push(line.name + ' ' + line.quantityText);
  }
  parts.push('');

  parts.push('よろしくお願いいたします。');
  parts.push('');
  parts.push(ORDER_STORE_NAME);
  if (normalizeIngredientText(ORDER_STORE_PHONE) !== '') {
    parts.push('TEL：' + ORDER_STORE_PHONE);
  }
  return parts.join('\n');
}

/**
 * 発注内容確認の画面に出す内容を作る（メールはまだ送らない）。
 * input - {authToken: string, items: {ingredientId: string, quantity: number}[]}
 * 戻り値 - {drafts: ({vendorId: string, vendorName: string, to: string, vendorEmail: string, lines: ({ingredientId: string, name: string, quantity: number, unit: string, quantityText: string})[], subject: string, body: string, canSend: boolean, sendBlockReason: string})[], vendorCount: number}
 */
function getOrderDrafts(input) {
  const payload = input || {};
  requireManagerRole_(payload.authToken);
  const drafts = buildOrderDrafts_(payload.items, new Date());
  return { 'drafts': drafts, 'vendorCount': drafts.length };
}

/**
 * 業者別に発注メールを送り、送信できた食材の発注履歴を保存する。
 * input - {authToken: string, operationKey: string, items: {ingredientId: string, quantity: number}[]}
 * 戻り値 - {message: string, detail: null または {message: string, actor: string, sentVendorIds: string[], sentCount: number, failures: string[]}, board: {today: string, todayInput: string, nowText: string, storeName: string, testRecipient: string, maxQuantity: number, vendors: ({id: string, name: string, address: string, phone: string, email: string, businessHours: string, closedDays: string, deadlineText: string, leadTimeText: string, planText: string})[], items: ({id: string, name: string, unit: string, vendorId: string, vendorName: unknown})[], history: ({orderNo: string, orderedAt: string, orderedAtLabel: string, vendorId: string, vendorName: string, orderedBy: string, mailTo: string, channelText: string, expectedDate: string, note: string, statusTexts: string[], lines: ({name: string, quantityText: string, status: string})[]})[], upcoming: ({id: string, name: string, orderedAt: string, quantityText: string, expectedDate: string, vendorId: string, vendorName: unknown, status: string})[]}}
 */
function sendOrderMails(input) {
  const payload = input || {};
  requireManagerRole_(payload.authToken);
  return runOrderOperation_(payload.operationKey, '発注メール送信', function () {
    const now = new Date();
    const drafts = buildOrderDrafts_(payload.items, now);
    const actor = resolveOrderActor_();
    const orderedAt = formatIngredientDateTime(now);
    const suppliers = getIngredientSuppliers();
    const supplierById = {};
    for (let i = 0; i < suppliers.length; i++) {
      supplierById[suppliers[i].id] = suppliers[i];
    }

    if (MailApp.getRemainingDailyQuota() < drafts.length) {
      throw new Error('本日のメール送信可能数が足りません。明日以降にもう一度お試しください');
    }

    const sheet = ensureIngredientOrderSheet();
    const sentVendorIds = [];
    const sentVendorNames = [];
    const failures = [];

    for (let i = 0; i < drafts.length; i++) {
      const draft = drafts[i];
      if (!draft.canSend) {
        failures.push(draft.sendBlockReason);
        continue;
      }
      try {
        sendOrderMail_(draft);
      } catch (error) {
        failures.push(draft.vendorName + '：' + error.message);
        continue;
      }

      appendOrderHistoryRows_(sheet, draft, supplierById[draft.vendorId], orderedAt, actor, now);
      sentVendorIds.push(draft.vendorId);
      sentVendorNames.push(draft.vendorName);
    }

    let message = '';
    if (sentVendorNames.length > 0) {
      message = sentVendorNames.join('・') + ' へ発注メールを送信し、発注履歴に保存しました';
    }
    if (failures.length > 0) {
      message = (message === '' ? '' : message + '。') + '送信できませんでした：' + failures.join(' / ');
    }
    if (message === '') {
      message = '送信できる業者がありませんでした';
    }

    return {
      'message': message,
      'actor': actor,
      'sentVendorIds': sentVendorIds,
      'sentCount': sentVendorNames.length,
      'failures': failures
    };
  });
}

/**
 * MailApp.sendEmailで業者別に組み立てたメールを送る。
 * draft - {vendorId: string, vendorName: string, to: string, vendorEmail: string, lines: {ingredientId: string, name: string, quantity: number, unit: string, quantityText: string}[], subject: string, body: string, canSend: boolean, sendBlockReason: string}
 * 戻り値 - なし
 */
function sendOrderMail_(draft) {
  MailApp.sendEmail({
    to: draft.to,
    subject: draft.subject,
    body: draft.body,
    name: ORDER_STORE_NAME
  });
}

/**
 * 送信した1業者分を食材ごとの行として食材発注へ追加する。
 * sheet - GoogleAppsScript.Spreadsheet.Sheet
 * draft - {vendorId: string, vendorName: string, to: string, vendorEmail: string, lines: {ingredientId: string, name: string, quantity: number, unit: string, quantityText: string}[], subject: string, body: string, canSend: boolean, sendBlockReason: string}
 * supplier - unknown
 * orderedAt - string
 * actor - string
 * now - Date
 * 戻り値 - なし
 */
function appendOrderHistoryRows_(sheet, draft, supplier, orderedAt, actor, now) {
  const orderNo = newIngredientId('ODR');
  const expectedDate = resolveOrderExpectedDate_(supplier, now);
  const updatedAt = formatIngredientDateTime(now);

  const rows = [];
  for (let i = 0; i < draft.lines.length; i++) {
    const line = draft.lines[i];
    rows.push([
      newIngredientId('ORD'), line.ingredientId, line.name, draft.vendorId,
      formatIngredientQuantity(line.quantity), line.unit, orderedAt, expectedDate,
      '発注済み', '', '', updatedAt,
      orderNo, actor, draft.to
    ]);
  }
  appendIngredientRows(sheet, rows, INGREDIENT_ORDER_SHEET_HEADER.length);
}

/**
 * 入荷予定日を求める（分からないときは空欄のままにして、推測で埋めない）。
 * supplier - unknown
 * now - null または Date
 * 戻り値 - undefined または string
 */
function resolveOrderExpectedDate_(supplier, now) {
  if (!supplier || !now) { return ''; }
  const plan = buildIngredientSupplierPlan(supplier, now);
  return plan.available ? plan.deliveryDateKey : '';
}

/**
 * 発注した人を求める 発注管理は店長画面の機能なので役割名を残し、実行中のGoogleアカウントが分かれば併記する。
 * 戻り値 - string
 */
function resolveOrderActor_() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (email) {
      return ROLE_MANAGER + '（' + email + '）';
    }
  } catch (error) {

  }
  return ROLE_MANAGER;
}

/**
 * ロック下で操作キーの履歴を照合し、完了済みの同じキーなら送信を省く。
 * operationKey - string
 * type - string
 * executor - (() => {message: string, actor: string, sentVendorIds: string[], sentCount: number, failures: string[]})
 * 戻り値 - {message: string, detail: null または {message: string, actor: string, sentVendorIds: string[], sentCount: number, failures: string[]}, board: {today: string, todayInput: string, nowText: string, storeName: string, testRecipient: string, maxQuantity: number, vendors: ({id: string, name: string, address: string, phone: string, email: string, businessHours: string, closedDays: string, deadlineText: string, leadTimeText: string, planText: string})[], items: ({id: string, name: string, unit: string, vendorId: string, vendorName: unknown})[], history: ({orderNo: string, orderedAt: string, orderedAtLabel: string, vendorId: string, vendorName: string, orderedBy: string, mailTo: string, channelText: string, expectedDate: string, note: string, statusTexts: string[], lines: ({name: string, quantityText: string, status: string})[]})[], upcoming: ({id: string, name: string, orderedAt: string, quantityText: string, expectedDate: string, vendorId: string, vendorName: unknown, status: string})[]}}
 */
function runOrderOperation_(operationKey, type, executor) {
  const key = normalizeIngredientText(operationKey);
  if (key === '') {
    throw new Error('操作キーがありません。画面を読み込み直してからもう一度お試しください');
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    throw new Error('ほかの保存処理が実行中です。少し待ってからもう一度お試しください');
  }

  let message = '';
  let detail = null;
  try {
    const eventSheet = ensureIngredientEventSheet();
    const existing = findIngredientEventByKey(eventSheet, key);
    let isDuplicated = false;
    if (existing) {
      const status = existing.values['処理状態'];
      if (status === '完了') {

        message = existing.values['内容'] + '（この発注はすでに送信済みです）';
        isDuplicated = true;
      } else if (status === '処理中') {
        throw new Error('前回の送信が完了していない可能性があります。発注履歴を確認してください');
      }

    }

    if (!isDuplicated) {
      const eventRowNumber = startIngredientEvent(eventSheet, key, type);
      let result;
      try {
        result = executor();
      } catch (error) {
        failIngredientEvent(eventSheet, eventRowNumber, error);
        throw error;
      }
      SpreadsheetApp.flush();
      completeIngredientEvent(eventSheet, eventRowNumber, result);
      message = result.message;
      detail = result;
    }
  } finally {
    lock.releaseLock();
  }

  return { 'message': message, 'detail': detail, 'board': buildOrderBoard_() };
}

/**
 * 発注に必要なマスタがそろっていない場合だけ、初期データを入れる すでに入っている値は書き換えず、空欄のところだけを埋める。
 * 戻り値 - なし
 */
function ensureOrderMasterData_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return;
  }
  try {
    const supplierSheet = ensureIngredientSupplierSheet();
    let supplierRows = readIngredientRows(supplierSheet, INGREDIENT_SUPPLIER_SHEET_HEADER);
    if (supplierRows.length === 0) {
      appendIngredientRows(supplierSheet, buildOrderDefaultVendorRows_(), INGREDIENT_SUPPLIER_SHEET_HEADER.length);
      supplierRows = readIngredientRows(supplierSheet, INGREDIENT_SUPPLIER_SHEET_HEADER);
    }

    const supplierIdByName = {};
    for (let i = 0; i < supplierRows.length; i++) {
      const row = supplierRows[i];
      const oldName = row.values['業者名'];
      const name = normalizeOrderSupplierName_(oldName);
      if (name !== oldName) {
        supplierSheet.getRange(row.rowNumber, INGREDIENT_SUPPLIER_SHEET_HEADER.indexOf('業者名') + 1).setValue(name);
        row.values['業者名'] = name;
      }
      supplierIdByName[name] = row.values['業者ID'];
    }

    const masterSheet = ensureIngredientMasterSheet();
    const masterRows = readIngredientRows(masterSheet, INGREDIENT_MASTER_SHEET_HEADER);
    const masterRowByName = {};
    for (let i = 0; i < masterRows.length; i++) {
      masterRowByName[masterRows[i].values['食材名']] = masterRows[i];
    }

    const newMasterRows = [];
    const updatedAt = formatIngredientDateTime(new Date());
    for (let i = 0; i < ORDER_DEFAULT_ITEMS.length; i++) {
      const item = ORDER_DEFAULT_ITEMS[i];
      const vendorId = supplierIdByName[item.vendorName] || '';
      const row = masterRowByName[item.name];
      if (!row) {
        newMasterRows.push([
          newIngredientId('ING'), item.name, ORDER_DEFAULT_ITEM_CATEGORY, item.unit,
          '', '', '', vendorId, '○', updatedAt, item.unit
        ]);
        continue;
      }

      const values = row.values;
      let changed = false;
      if (values['業者ID'] === '' && vendorId !== '') {
        values['業者ID'] = vendorId;
        changed = true;
      }
      if (values['発注単位'] === '') {
        values['発注単位'] = item.unit;
        changed = true;
      }
      if (changed) {
        values['更新日時'] = updatedAt;
        writeOrderSheetRow_(masterSheet, row.rowNumber, INGREDIENT_MASTER_SHEET_HEADER, values);
      }
    }
    if (newMasterRows.length > 0) {
      appendIngredientRows(masterSheet, newMasterRows, INGREDIENT_MASTER_SHEET_HEADER.length);
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * 業者マスタへ入れる初期データの行を作る。
 * 戻り値 - (string[])[]
 */
function buildOrderDefaultVendorRows_() {
  const rows = [];
  const updatedAt = formatIngredientDateTime(new Date());
  for (let i = 0; i < ORDER_DEFAULT_VENDORS.length; i++) {
    const vendor = ORDER_DEFAULT_VENDORS[i];
    rows.push([
      newIngredientId('SUP'), vendor.name, vendor.deadlineTime,
      vendor.orderDays.join(','), vendor.shipDays.join(','), vendor.deliveryDays.join(','),
      vendor.leadTimeDays, vendor.leadTimeCount, '', vendor.pickupDays.join(','),
      vendor.note, updatedAt,
      vendor.address, vendor.phone, vendor.email, vendor.businessHours,
      vendor.closedDays, vendor.mainItems, vendor.orderNote
    ]);
  }
  return rows;
}

/**
 * 見出し名をキーにした値で、シートの1行を書き直す。
 * sheet - GoogleAppsScript.Spreadsheet.Sheet
 * rowNumber - unknown
 * header - string[]
 * values - unknown
 * 戻り値 - なし
 */
function writeOrderSheetRow_(sheet, rowNumber, header, values) {
  const row = [];
  for (let i = 0; i < header.length; i++) {
    const value = values[header[i]];
    row.push((value === undefined || value === null) ? '' : value);
  }
  const range = sheet.getRange(rowNumber, 1, 1, header.length);
  range.setNumberFormat('@');
  range.setValues([row]);
}
