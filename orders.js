// 発注管理（電話・FAXでやっていた食材の発注を、アプリの中でまとめて行う）を扱う
//
// 扱い方の前提
// ・業者と食材のマスタは「食材業者・食材マスタ」シートで管理する
// ・「どの食材をどの業者へ発注するか」は食材マスタの業者IDで決まるので、利用者が業者を選ぶ必要はない
// ・1通のメールにまとめた食材は同じ発注番号を持つ。発注履歴はこの発注番号で業者ごとにまとめて表示する
// ・メール本文は業者マスタの内容から組み立てる。画面から受け取った文章をそのまま送ることはしない

// 発注履歴の一覧に出す最大件数（古い記録はシートに残る）
const ORDER_HISTORY_LIMIT = 100;

/**
 * 発注管理画面に必要な情報をまとめて返す（画面から呼ぶ入口）
 * 発注は店長画面の機能なので、ここで店長かどうかを確認する
 * @param {string} authToken ログイントークン
 * @return {Object} 業者・食材・発注履歴・設定値
 */
function getOrderBoard(authToken) {
  requireManagerRole_(authToken);
  ensureOrderMasterData_();
  return buildOrderBoard_();
}

/**
 * 発注管理画面に渡す情報を組み立てる
 * 送信処理からも呼ぶため、ここでは権限を確認しない
 * @return {Object} 業者・食材・発注履歴・設定値
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
      continue; // 使わなくなった食材は発注できないようにする
    }
    if (master.category !== ORDER_DEFAULT_ITEM_CATEGORY) {
      continue; // 仕込み済みの品は店内で作るものなので発注の対象にしない
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
    'history': buildOrderHistory_(orders, supplierById)
  };
}

/**
 * 業者1件を画面に出す形に整える
 * 未登録の項目は推測せず、画面側で「未登録」と出せるよう空文字のまま渡す
 */
function buildOrderVendorView_(supplier, now) {
  const rules = buildOrderMailRules_(supplier.orderNote);
  return {
    'id': supplier.id,
    'name': supplier.name,
    'address': supplier.address,
    'phone': supplier.phone,
    'email': supplier.email,
    'businessHours': supplier.businessHours,
    'closedDays': supplier.closedDays,
    // 発注締切は業者マスタに登録した「受注締切時刻」と「受注営業日」をそのまま使う
    'deadlineText': buildOrderDeadlineText_(supplier),
    'leadTimeText': buildOrderLeadTimeText_(supplier),
    'deliveryDaysText': formatIngredientWeekdayList(supplier.deliveryDays),
    'mainItems': supplier.mainItems,
    'orderNote': supplier.orderNote,
    'note': supplier.note,
    'mailToText': resolveOrderRecipient_(supplier),
    'appliedRuleTexts': rules.appliedRuleTexts,
    'planText': buildIngredientSupplierPlan(supplier, now).text
  };
}

/**
 * 発注締切の表示文を作る（未設定のときは推測しない）
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
 * 通常納品日数の表示文を作る
 */
function buildOrderLeadTimeText_(supplier) {
  if (supplier.leadTimeDays === null) {
    return '';
  }
  const countText = supplier.leadTimeCount === '' ? '営業日' : supplier.leadTimeCount;
  return supplier.leadTimeDays + countText;
}

/**
 * 実際にメールを送る宛先を求める
 * 動作確認用の宛先が設定されている間は、業者マスタのアドレスではなくそちらへ送る
 * @param {Object} supplier 業者
 * @return {string} 宛先アドレス（どちらも未設定なら空文字）
 */
function resolveOrderRecipient_(supplier) {
  if (normalizeIngredientText(ORDER_TEST_RECIPIENT_EMAIL) !== '') {
    return ORDER_TEST_RECIPIENT_EMAIL;
  }
  return supplier.email;
}

// ここから下は発注履歴（食材発注シート）の読み取り

/**
 * 発注の記録を、1通の発注（業者ごと）にまとめ直して新しい順に並べる
 * 発注番号が空の行はメール送信を通さずに記録された発注なので、1件ずつのまとまりとして扱う
 * @param {Array<Object>} orders 食材発注シートの内容
 * @param {Object} supplierById 業者IDをキーにした業者の対応表
 * @return {Array<Object>} 発注履歴の配列（新しい順）
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
    return valueB - valueA; // 新しい発注を上にする
  });

  return groups.slice(0, ORDER_HISTORY_LIMIT);
}

/**
 * 発注日時（yyyy/MM/dd HH:mm）を一覧用の短い表示（9/16 21:05）にする
 */
function formatOrderHistoryLabel_(orderedAt) {
  const date = parseIngredientDateTime(orderedAt);
  if (!date) {
    return orderedAt;
  }
  return Utilities.formatDate(date, TIME_ZONE, 'M/d HH:mm');
}

// ここから下はメールの組み立て

/**
 * 業者の「発注時の注意事項」から、メールの作り方の決まりを読み取る
 * 決まった言葉が書かれていればその形に合わせ、書かれていなければ何もしない（勝手な書式は足さない）
 * @param {string} orderNote 発注時の注意事項
 * @return {Object} { includeDeliveryDate, storeNameInSubject, extraTexts, appliedRuleTexts }
 */
function buildOrderMailRules_(orderNote) {
  const rules = {
    'includeDeliveryDate': false,
    'storeNameInSubject': false,
    'extraTexts': [],
    'appliedRuleTexts': []
  };
  const text = normalizeIngredientText(orderNote);
  if (text === '') {
    return rules;
  }

  const lines = text.split(/\r\n|\r|\n/);
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (line === '') {
      continue;
    }
    // 「本文追記：〜」と書かれた行は、その文章をそのまま本文に入れる
    const extraPrefix = ORDER_MAIL_EXTRA_TEXT_PREFIX.replace(/[：:]\s*$/, '');
    if (line.indexOf(extraPrefix) === 0) {
      const extra = line.substring(extraPrefix.length).replace(/^[：:]\s*/, '').trim();
      if (extra !== '') {
        rules.extraTexts.push(extra);
        rules.appliedRuleTexts.push('本文に「' + extra + '」を入れます');
      }
      continue;
    }
    if (line.indexOf(ORDER_MAIL_RULE_KEYWORDS.DELIVERY_DATE) !== -1) {
      rules.includeDeliveryDate = true;
      rules.appliedRuleTexts.push('本文に納品希望日を入れます');
    }
    if (line.indexOf(ORDER_MAIL_RULE_KEYWORDS.SUBJECT_STORE_NAME) !== -1 &&
        line.indexOf(ORDER_MAIL_RULE_KEYWORDS.STORE_NAME) !== -1) {
      rules.storeNameInSubject = true;
      rules.appliedRuleTexts.push('件名に店舗名を入れます');
    }
  }
  return rules;
}

/**
 * 画面から届いた発注内容を確かめ、食材マスタの内容で組み立て直す
 * 食材名・業者・単位はマスタの値を使い、画面から届いた値は信用しない
 * @param {Array<Object>} items [{ ingredientId, quantity }] の配列
 * @return {Array<Object>} { master, supplier, quantity, unit } の配列
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
 * 発注内容を業者ごとにまとめ、業者ごとのメール（件名・本文・宛先）を組み立てる
 * 確認画面の表示にも送信にも同じこの関数を使うので、確認した内容とそのまま同じメールが送られる
 * @param {Array<Object>} items [{ ingredientId, quantity }] の配列
 * @param {string} deliveryDate 納品希望日（未入力なら空文字）
 * @param {string} note 連絡事項（未入力なら空文字）
 * @param {Date} now 現在日時
 * @return {Array<Object>} 業者ごとの発注メールの配列
 */
function buildOrderDrafts_(items, deliveryDate, note, now) {
  const normalized = normalizeOrderItems_(items);
  const requestedDate = normalizeIngredientDateInput(deliveryDate, '納品希望日');
  const noteText = normalizeIngredientText(note);

  const drafts = [];
  const draftByVendorId = {};
  for (let i = 0; i < normalized.length; i++) {
    const line = normalized[i];
    const vendorId = line.supplier.id;
    if (!draftByVendorId[vendorId]) {
      draftByVendorId[vendorId] = {
        'vendorId': vendorId,
        'vendorName': line.supplier.name,
        'supplier': line.supplier,
        'rules': buildOrderMailRules_(line.supplier.orderNote),
        'to': resolveOrderRecipient_(line.supplier),
        'vendorEmail': line.supplier.email,
        'orderNote': line.supplier.orderNote,
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
    draft.subject = buildOrderMailSubject_(draft.rules, now);
    draft.body = buildOrderMailBody_(draft, requestedDate, noteText);
    draft.appliedRuleTexts = draft.rules.appliedRuleTexts;
    draft.canSend = (draft.to !== '');
    draft.sendBlockReason = draft.canSend ? '' : (draft.vendorName + ' のメールアドレスが未登録です');
    draft.deliveryDate = requestedDate;
    delete draft.supplier; // 画面へはメールの内容だけを渡す
    delete draft.rules;
  }
  return drafts;
}

/**
 * 件名を作る（例：9月16日 発注依頼）
 */
function buildOrderMailSubject_(rules, now) {
  const base = (now.getMonth() + 1) + '月' + now.getDate() + '日 発注依頼';
  return rules.storeNameInSubject ? '【' + ORDER_STORE_NAME + '】' + base : base;
}

/**
 * 本文を作る
 * 業者名・商品・数量は登録された内容から組み立て、利用者に文章を入力させない
 */
function buildOrderMailBody_(draft, deliveryDate, noteText) {
  const parts = [];
  parts.push(draft.vendorName + ' 御中');
  parts.push('');
  parts.push('いつもお世話になっております。' + ORDER_STORE_NAME + 'です。');
  parts.push('以下の商品をお願いいたします。');
  parts.push('');

  if (draft.rules.includeDeliveryDate) {
    parts.push('納品希望日：' + (deliveryDate === '' ? '指定なし（通常の納品日でお願いいたします）' : deliveryDate));
    parts.push('');
  }

  for (let i = 0; i < draft.lines.length; i++) {
    const line = draft.lines[i];
    parts.push(line.name + ' ' + line.quantityText);
  }
  parts.push('');

  if (noteText !== '') {
    parts.push('【連絡事項】');
    parts.push(noteText);
    parts.push('');
  }
  for (let i = 0; i < draft.rules.extraTexts.length; i++) {
    parts.push(draft.rules.extraTexts[i]);
  }
  if (draft.rules.extraTexts.length > 0) {
    parts.push('');
  }

  parts.push('よろしくお願いいたします。');
  parts.push('');
  parts.push(ORDER_STORE_NAME);
  if (normalizeIngredientText(ORDER_STORE_PHONE) !== '') {
    parts.push('TEL：' + ORDER_STORE_PHONE);
  }
  return parts.join('\n');
}

/**
 * 発注内容確認の画面に出す内容を作る（メールはまだ送らない）
 * @param {Object} input { authToken, items, deliveryDate, note }
 * @return {Object} { drafts: 業者ごとの発注メール, vendorCount: 業者数 }
 */
function getOrderDrafts(input) {
  const payload = input || {};
  requireManagerRole_(payload.authToken);
  const drafts = buildOrderDrafts_(payload.items, payload.deliveryDate, payload.note, new Date());
  return { 'drafts': drafts, 'vendorCount': drafts.length };
}

// ここから下は送信と発注履歴の保存

/**
 * 発注メールを業者ごとに送り、送れた分を発注履歴として保存する
 * 途中の業者で失敗しても、送信できた業者の記録は残す（同じ内容を二重に送らないため）
 * @param {Object} input { authToken, operationKey, items, deliveryDate, note }
 * @return {Object} { message, detail, board }
 */
function sendOrderMails(input) {
  const payload = input || {};
  requireManagerRole_(payload.authToken);
  return runOrderOperation_(payload.operationKey, '発注メール送信', function () {
    const now = new Date();
    const drafts = buildOrderDrafts_(payload.items, payload.deliveryDate, payload.note, now);
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
      // 送れた業者の分だけを記録する。ここで失敗しても、送ったメールは取り消せないので例外にしない
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
 * 発注メールを1通送る
 * 差出人に使いたいアドレスがGmailの別名に登録されていればその名前で送る
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
 * 送信できた1業者分を食材発注シートへ追記する（食材1つにつき1行）
 * 入荷予定日は、納品希望日があればその日、無ければ業者の発注条件から求めた納品日を入れる
 */
function appendOrderHistoryRows_(sheet, draft, supplier, orderedAt, actor, now) {
  const orderNo = newIngredientId('ODR');
  const expectedDate = resolveOrderExpectedDate_(draft.deliveryDate, supplier, now);
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
 * 入荷予定日を求める（分からないときは空欄のままにして、推測で埋めない）
 */
function resolveOrderExpectedDate_(deliveryDate, supplier, now) {
  if (deliveryDate !== '') {
    return deliveryDate;
  }
  if (!supplier) {
    return '';
  }
  const plan = buildIngredientSupplierPlan(supplier, now);
  if (!plan.available) {
    return '';
  }
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const expected = findIngredientDeliveryDate(supplier, today);
  return expected ? formatIngredientDate(expected) : '';
}

/**
 * 発注した人を求める
 * 発注管理は店長画面の機能なので役割名を残し、実行中のGoogleアカウントが分かれば併記する
 * @return {string} 発注者
 */
function resolveOrderActor_() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (email) {
      return ROLE_MANAGER + '（' + email + '）';
    }
  } catch (error) {
    // 実行者を取得できない設定のときは役割名だけを残す
  }
  return ROLE_MANAGER;
}

/**
 * 送信処理の共通の入れ物
 * 同時実行を締め出し、同じ操作キーの再送では二重に送らない
 * @param {string} operationKey 画面が1回の送信ごとに発行する操作キー
 * @param {string} type 履歴に残す操作の種別
 * @param {Function} executor 実際の送信処理
 * @return {Object} { message: 画面に出す文言, detail: 送信結果, board: 最新の発注管理情報 }
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
        // 通信のやり直しで同じ送信が届いた場合。メールは送らず、前回の結果をそのまま返す
        message = existing.values['内容'] + '（この発注はすでに送信済みです）';
        isDuplicated = true;
      } else if (status === '処理中') {
        throw new Error('前回の送信が完了していない可能性があります。発注履歴を確認してください');
      }
      // '失敗' のときはメールを送っていないので、同じ操作キーのままやり直せる
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
      SpreadsheetApp.flush(); // 発注履歴の書き込みを確定させてから履歴を完了にする
      completeIngredientEvent(eventSheet, eventRowNumber, result);
      message = result.message;
      detail = result;
    }
  } finally {
    lock.releaseLock();
  }

  return { 'message': message, 'detail': detail, 'board': buildOrderBoard_() };
}

// ここから下は初期データの用意

/**
 * 発注に必要なマスタがそろっていない場合だけ、初期データを入れる
 * すでに入っている値は書き換えず、空欄のところだけを埋める
 * @return {void} 戻り値なし
 */
function ensureOrderMasterData_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return; // ほかの処理が動いているときは初期データを入れず、いまある内容で画面を出す
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
      supplierIdByName[supplierRows[i].values['業者名']] = supplierRows[i].values['業者ID'];
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
      // すでにある食材は、空欄の項目だけを埋める（店長が設定した内容は書き換えない）
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
 * 業者マスタへ入れる初期データの行を作る
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
 * 見出し名をキーにした値で、シートの1行を書き直す
 * @param {Sheet} sheet 対象シート
 * @param {number} rowNumber 書き直す行番号
 * @param {Array<string>} header 見出し行
 * @param {Object} values 見出し名をキーにした値
 * @return {void} 戻り値なし
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
