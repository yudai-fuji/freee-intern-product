// 共用PCで使うログイン（パスコード認証）と、店長専用処理の権限確認をまとめる
// パスコードの照合はすべてこのサーバー側で行い、画面から渡された役割名は信用しない

/**
 * パスコードを照合してログインする
 * @param {string} passcode 画面で入力されたパスコード（先頭の 0 を保持するため文字列で受け取る）
 * @return {Object} 成功なら { success: true, role: 役割名, token: ログイントークン }、失敗なら { success: false, message: 文言 }
 */
function login(passcode) {
  const inputText = (passcode === null || passcode === undefined) ? '' : String(passcode).trim();
  const role = LOGIN_PASSCODES[inputText];
  if (!role) {
    return { 'success': false, 'message': 'パスコードが違います' };
  }
  return { 'success': true, 'role': role, 'token': issueAuthToken_(role) };
}

/**
 * ログイントークンを無効にする（手動ログアウト・自動ログアウトの両方から呼ぶ）
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
