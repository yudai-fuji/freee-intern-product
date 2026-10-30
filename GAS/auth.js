/**
 * 店長パスワードを照合し、認証結果を返す。
 * passcode - string
 * 戻り値 - {success: boolean, message: string, role?: undefined, token?: undefined} または {success: boolean, role: string, token: string, message?: undefined}
 */
function login(passcode) {
  ensureAppDataMigration_();
  if (typeof passcode !== 'string' || passcode !== PropertiesService.getScriptProperties().getProperty(MANAGER_PASSWORD_PROPERTY)) {
    return { 'success': false, 'message': 'パスワードが違います' };
  }
  return { 'success': true, 'role': ROLE_MANAGER, 'token': issueAuthToken_(ROLE_MANAGER) };
}

/**
 * 従業員用のトークンを発行。
 * 戻り値 - {success: boolean, role: string, token: string}
 */
function createEmployeeSession() {
  ensureAppDataMigration_();
  return { 'success': true, 'role': ROLE_EMPLOYEE, 'token': issueAuthToken_(ROLE_EMPLOYEE)};
}

/**
 * 更新版の初回起動でのみ移行。
 * 戻り値 - なし
 */
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

/**
 * PINが文字列で半角数字6桁かを確認する。
 * pin - string
 * 戻り値 - なし
 */
function requireValidEmployeePin_(pin) {
  if (typeof pin !== 'string' || pin.length !== 6 || /[^0-9]/.test(pin)) {
    throw new Error('PINは半角数字6桁で入力してください');
  }
}

/**
 * 店長パスワードの長さ・文字種を検証する。
 * password - string
 * 戻り値 - なし
 */
function requireValidManagerPassword_(password) {
  if (typeof password !== 'string' || password.length < MANAGER_PASSWORD_MIN_LENGTH ||
      password.length > MANAGER_PASSWORD_MAX_LENGTH || /[^A-Za-z0-9]/.test(password)) {
    throw new Error('新しいパスワードは半角英数字8〜64文字で入力してください（記号・空白は使用できません）');
  }
}

/**
 * 店長権限と現在のパスワードを確かめ、新しいパスワードをScript Propertiesへ保存する。
 * currentPassword - string
 * newPassword - string
 * confirmPassword - string
 * authToken - string
 * 戻り値 - {message: string}
 */
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

/**
 * PINそのものは画面へ返さず、今回の打刻だけに使える認証結果を返す。
 * empId - unknown
 * type - unknown
 * pin - string
 * authToken - string
 * 戻り値 - {ticket: string, expiresAt: number}
 */
function authenticateClockPin(empId, type, pin, authToken) {
  requireLoggedIn_(authToken);
  requireValidEmployeePin_(pin);
  if (CLOCK_TYPES.indexOf(type) === -1) throw new Error('不正な打刻種別です');
  return withAppLock_(function () {

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

/**
 * 打刻のロック内で呼ぶ。
 * empId - unknown
 * type - unknown
 * ticket - unknown
 * authToken - unknown
 * now - Date
 * 戻り値 - なし
 */
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
 * ログイントークンを無効にする（店長画面から従業員画面へ戻るときに、使い終えた店長用トークンを破棄する）。
 * token - string
 * 戻り値 - boolean
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
 * 役割に対するログイントークンを新しく発行し、キャッシュに保存する。
 * role - string
 * 戻り値 - string
 */
function issueAuthToken_(role) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put(AUTH_TOKEN_CACHE_PREFIX + token, role, AUTH_TOKEN_TTL_SECONDS);
  return token;
}

/**
 * ログイントークンから、サーバー側に記録してある役割を取り出す。
 * token - unknown
 * 戻り値 - string
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
 * ログインしているかを確認し、そうでなければ処理を中断する 役割を問わず使ってよいデータ（従業員一覧など）の取得の先頭で呼ぶ。
 * token - unknown
 * 戻り値 - なし
 */
function requireLoggedIn_(token) {
  if (getRoleByToken_(token) === '') {
    throw new Error('ログインし直してください');
  }
}

/**
 * 店長としてログインしているかを確認し、そうでなければ処理を中断する 店長専用のデータ取得・更新の先頭で呼ぶ。
 * token - unknown
 * 戻り値 - なし
 */
function requireManagerRole_(token) {
  if (getRoleByToken_(token) !== ROLE_MANAGER) {
    throw new Error('店長としてログインし直してください');
  }
}

/**
 * ログイントークンをキャッシュのキーに変換する。
 * token - unknown
 * 戻り値 - string
 */
function buildAuthCacheKey_(token) {
  const text = (token === null || token === undefined) ? '' : String(token).trim();
  if (text === '') {
    return '';
  }
  return AUTH_TOKEN_CACHE_PREFIX + text;
}
