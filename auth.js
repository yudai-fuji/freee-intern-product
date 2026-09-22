// 共用PCで使うログイン（パスコード認証）と、店長専用処理の権限確認をまとめる
// パスコードの照合はすべてこのサーバー側で行い、画面から渡された役割名は信用しない

/**
 * パスコードを照合してログインする
 * @param {string} passcode 画面で入力されたパスコード（先頭の 0 を保持するため文字列で受け取る）
 * @return {Object} 成功なら { success: true, role: 役割名, token: ログイントークン }、失敗なら { success: false, message: 文言 }
 */
function login(passcode) {
  ensureAppDataMigration_();
  if (typeof passcode !== 'string' || passcode !== PropertiesService.getScriptProperties().getProperty(MANAGER_PASSWORD_PROPERTY)) {
    return { 'success': false, 'message': 'パスワードが違います' };
  }
  return { 'success': true, 'role': ROLE_MANAGER, 'token': issueAuthToken_(ROLE_MANAGER) };
}

// 従業員用のトークンを発行
function createEmployeeSession() {
  ensureAppDataMigration_();
  return { 'success': true, 'role': ROLE_EMPLOYEE, 'token': issueAuthToken_(ROLE_EMPLOYEE)};
}

// 更新版の初回起動でのみ移行。途中で失敗した場合は次回起動で再試行する。
function ensureAppDataMigration_() {
  const properties = PropertiesService.getScriptProperties();
  if (properties.getProperty(APP_DATA_MIGRATION_PROPERTY) === 'done') return;
  withAppLock_(function () {
    if (properties.getProperty(APP_DATA_MIGRATION_PROPERTY) === 'done') return;
    migrateEmployeePins_();
    if (properties.getProperty(MANAGER_PASSWORD_PROPERTY) === null) {
      properties.setProperty(MANAGER_PASSWORD_PROPERTY, MANAGER_PASSWORD_DEFAULT);
    }
    migrateSkillRatingsToText();
    SpreadsheetApp.flush();
    properties.setProperty(APP_DATA_MIGRATION_PROPERTY, 'done');
  });
}

function requireValidEmployeePin_(pin) {
  if (typeof pin !== 'string' || pin.length !== 6 || /[^0-9]/.test(pin)) {
    throw new Error('PINは半角数字6桁で入力してください');
  }
}

function requireValidManagerPassword_(password) {
  if (typeof password !== 'string' || password.length < MANAGER_PASSWORD_MIN_LENGTH ||
      password.length > MANAGER_PASSWORD_MAX_LENGTH || /[^A-Za-z0-9]/.test(password)) {
    throw new Error('新しいパスワードは半角英数字8〜64文字で入力してください（記号・空白は使用できません）');
  }
}

function changeManagerPassword(currentPassword, newPassword, confirmPassword, authToken) {
  requireManagerRole_(authToken);
  requireValidManagerPassword_(newPassword);
  requireValidManagerPassword_(confirmPassword);
  if (newPassword !== confirmPassword) throw new Error('新しいパスワードと確認欄が一致しません');
  return withAppLock_(function () {
    const properties = PropertiesService.getScriptProperties();
    if (typeof currentPassword !== 'string' || currentPassword !== properties.getProperty(MANAGER_PASSWORD_PROPERTY)) {
      throw new Error('現在のパスワードが違います');
    }
    properties.setProperty(MANAGER_PASSWORD_PROPERTY, newPassword);
    return { message: 'パスワードを変更しました' };
  });
}

// PINそのものは画面へ返さず、今回の打刻だけに使える認証結果を返す。
function authenticateClockPin(empId, type, pin, authToken) {
  requireLoggedIn_(authToken);
  requireValidEmployeePin_(pin);
  if (CLOCK_TYPES.indexOf(type) === -1) throw new Error('不正な打刻種別です');
  return withAppLock_(function () {
    // readEmployeePin_でも有効な従業員かを確認する。
    if (readEmployeePin_(empId) !== pin) throw new Error('PINが違います');
    const now = new Date();
    const state = readTodayClockState_(empId, getSheet(SHEET_NAMES.TIME_CLOCKS), now);
    if (state.state === 'invalid') throw new Error(state.error);
    if (nextClockState_(state.state, type) === null) throw new Error('現在の打刻状態では「' + type + '」を記録できません');
    const ticket = Utilities.getUuid();
    const expiresAt = now.getTime() + CLOCK_PIN_TICKET_TTL_SECONDS * 1000;
    CacheService.getScriptCache().put(CLOCK_PIN_TICKET_PREFIX + ticket, JSON.stringify({
      employeeId: String(empId), type: type, date: state.date, authToken: authToken, expiresAt: expiresAt
    }), CLOCK_PIN_TICKET_TTL_SECONDS);
    return { ticket: ticket, expiresAt: expiresAt };
  });
}

// 打刻のロック内で呼ぶ。認証結果を他の従業員・操作・日付・セッションに使い回せない。
function consumeClockPinTicket_(empId, type, ticket, authToken, now) {
  requireLoggedIn_(authToken);
  const cache = CacheService.getScriptCache();
  const raw = typeof ticket === 'string' && ticket !== '' ? cache.get(CLOCK_PIN_TICKET_PREFIX + ticket) : null;
  const grant = raw ? JSON.parse(raw) : null;
  if (!grant || grant.employeeId !== String(empId) || grant.type !== type || grant.authToken !== authToken ||
      grant.date !== Utilities.formatDate(now, TIME_ZONE, 'yyyy/MM/dd') || grant.expiresAt <= now.getTime()) {
    throw new Error('PIN認証が無効または期限切れです。PINを再入力してください');
  }
  cache.remove(CLOCK_PIN_TICKET_PREFIX + ticket);
}

/**
 * ログイントークンを無効にする（店長画面から従業員画面へ戻るときに、使い終えた店長用トークンを破棄する）
 * @param {string} token 破棄するログイントークン
 * @return {boolean} 破棄を実行したら true、トークンが空なら false
 */
function logout(token) {
  const cacheKey = buildAuthCacheKey_(token);
  if (cacheKey === '') {
    return false;
  }
  CacheService.getScriptCache().remove(cacheKey);
  return true;
}

/**
 * 役割に対するログイントークンを新しく発行し、キャッシュに保存する
 * @param {string} role 役割名（ROLE_EMPLOYEE か ROLE_MANAGER）
 * @return {string} 発行したログイントークン
 */
function issueAuthToken_(role) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put(AUTH_TOKEN_CACHE_PREFIX + token, role, AUTH_TOKEN_TTL_SECONDS);
  return token;
}

/**
 * ログイントークンから、サーバー側に記録してある役割を取り出す
 * @param {string} token 画面から渡されたログイントークン
 * @return {string} 役割名。トークンが無効・期限切れなら空文字
 */
function getRoleByToken_(token) {
  const cacheKey = buildAuthCacheKey_(token);
  if (cacheKey === '') {
    return '';
  }
  const role = CacheService.getScriptCache().get(cacheKey);
  return role ? role : '';
}

/**
 * ログインしているかを確認し、そうでなければ処理を中断する
 * 役割を問わず使ってよいデータ（従業員一覧など）の取得の先頭で呼ぶ
 * @param {string} token 画面から渡されたログイントークン
 * @return {void} 確認に通れば何も返さない（通らない場合は例外を投げる）
 */
function requireLoggedIn_(token) {
  if (getRoleByToken_(token) === '') {
    throw new Error('ログインし直してください');
  }
}

/**
 * 店長としてログインしているかを確認し、そうでなければ処理を中断する
 * 店長専用のデータ取得・更新の先頭で呼ぶ
 * @param {string} token 画面から渡されたログイントークン
 * @return {void} 確認に通れば何も返さない（通らない場合は例外を投げる）
 */
function requireManagerRole_(token) {
  if (getRoleByToken_(token) !== ROLE_MANAGER) {
    throw new Error('店長としてログインし直してください');
  }
}

/**
 * ログイントークンをキャッシュのキーに変換する
 * @param {string} token ログイントークン
 * @return {string} キャッシュのキー。トークンが空なら空文字
 */
function buildAuthCacheKey_(token) {
  const text = (token === null || token === undefined) ? '' : String(token).trim();
  if (text === '') {
    return '';
  }
  return AUTH_TOKEN_CACHE_PREFIX + text;
}
