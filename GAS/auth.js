 
 

 




function login(passcode) {
  ensureAppDataMigration_();
  if (typeof passcode !== 'string' || passcode !== PropertiesService.getScriptProperties().getProperty(MANAGER_PASSWORD_PROPERTY)) {
    return { 'success': false, 'message': 'パスワードが違います' };
  }
  return { 'success': true, 'role': ROLE_MANAGER, 'token': issueAuthToken_(ROLE_MANAGER) };
}

 
function createEmployeeSession() {
  ensureAppDataMigration_();
  return { 'success': true, 'role': ROLE_EMPLOYEE, 'token': issueAuthToken_(ROLE_EMPLOYEE)};
}

 
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

 




function logout(token) {
  const cacheKey = buildAuthCacheKey_(token);
  if (cacheKey === '') {
    return false;
  }
  CacheService.getScriptCache().remove(cacheKey);
  return true;
}

 




function issueAuthToken_(role) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put(AUTH_TOKEN_CACHE_PREFIX + token, role, AUTH_TOKEN_TTL_SECONDS);
  return token;
}

 




function getRoleByToken_(token) {
  const cacheKey = buildAuthCacheKey_(token);
  if (cacheKey === '') {
    return '';
  }
  const role = CacheService.getScriptCache().get(cacheKey);
  return role ? role : '';
}

 





function requireLoggedIn_(token) {
  if (getRoleByToken_(token) === '') {
    throw new Error('ログインし直してください');
  }
}

 





function requireManagerRole_(token) {
  if (getRoleByToken_(token) !== ROLE_MANAGER) {
    throw new Error('店長としてログインし直してください');
  }
}

 




function buildAuthCacheKey_(token) {
  const text = (token === null || token === undefined) ? '' : String(token).trim();
  if (text === '') {
    return '';
  }
  return AUTH_TOKEN_CACHE_PREFIX + text;
}
