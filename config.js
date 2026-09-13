//　アプリ全体で共通して使う定数・設定値をまとめる

const SHEET_NAMES = {
  CHECK_RESULT: 'チェック結果',
  EMPLOYEES: '従業員名簿',
  TIME_CLOCKS: '打刻履歴',
  HOLIDAYS: '祝日',
  SKILLS: 'スキル管理',
  TODAY_WORK: '本日の業務',
  SHIFTS: 'シフト',
  SHIFT_REQUIREMENTS: 'シフト必要人数'
};

const SKILL_ITEMS = ['接客', 'レジ', 'バー', '厨房', '肉の仕込み', '締め作業'];

const SKILL_SHEET_HEADER = [
  '従業員ID', '氏名',
  '接客評価', '更新日',
  'レジ評価', '更新日',
  'バー評価', '更新日',
  '厨房評価', '厨房更新日',
  '肉の仕込み評価', '更新日',
  '締め作業評価', '更新日'
];

const SKILL_LEVELS = ['◎', '○', '△', '×'];

const TIME_ZONE = 'Asia/Tokyo';

const UNRECORDED_LABEL = '未打刻';

const NO_VALUE_LABEL = '-';

const CLOCK_TYPES = ['出勤', '退勤', '休憩開始', '休憩終了'];

const POSITION_ITEMS = ['ホール', '厨房', '肉の仕込み', '締め作業'];

const TODAY_WORK_SHEET_HEADER = ['日付', '従業員ID', 'ポジション', '確認事項'];

const SHIFT_SHEET_HEADER = ['日付', '従業員ID', '出勤時刻', '退勤時刻', 'ポジション', '備考', '更新日時'];

const SHIFT_REQUIREMENT_SHEET_HEADER = ['日付', '開始時刻', '終了時刻', 'ポジション', '必要人数', '更新日時'];

// ここから下は食材管理で使う設定値

// 食材管理で使うシート名（SHEET_NAMES に後から足すことで既存の定義を書き換えない）
SHEET_NAMES.INGREDIENT_MASTERS = '食材マスタ';
SHEET_NAMES.INGREDIENT_LOTS = '食材在庫';
SHEET_NAMES.INGREDIENT_EVENTS = '食材在庫履歴';
SHEET_NAMES.INGREDIENT_SUPPLIERS = '食材業者';
SHEET_NAMES.INGREDIENT_ORDERS = '食材発注';

// 食材の区分（原材料そのものか、店内で仕込んだものか）
const INGREDIENT_CATEGORIES = ['原材料', '仕込み済み'];

// 保存状態。区分とは別の属性なので、仕込み済みの食品も冷凍・解凍中になり得る
const INGREDIENT_STORAGE_STATES = ['冷蔵', '冷凍', '解凍中'];

// 目視で数えられる単位の候補。ここに足せば画面の選択肢が増える
const INGREDIENT_UNITS = ['人前', '杯', '袋', '個', 'g', 'kg', '本', '枚', 'パック', '営業日分'];

// ラベルに書かれている期限の種類。「不明」は要確認として扱う
const INGREDIENT_EXPIRY_KINDS = ['表示期限', '使用期限', '不明'];

// 冷凍の区別。納品時から冷凍の品と、店内で冷凍した品を混同しないために持つ
const INGREDIENT_FREEZE_KINDS = ['納品時冷凍', '店内冷凍'];

// 期限が何日前から「期限接近」になるか。運用に合わせてここだけ変えれば画面と通知の両方に反映される
const INGREDIENT_EXPIRY_SOON_DAYS = 3;

// 納品リードタイムの数え方
const INGREDIENT_LEAD_TIME_COUNTS = ['営業日', '暦日'];

// 発注の状態
const INGREDIENT_ORDER_STATUSES = ['発注済み', '入荷済み', '取消'];

// 在庫明細の行状態。'保留' は保存が途中で終わった行で、残量には数えず要確認として表示する
const INGREDIENT_LOT_STATUSES = ['有効', '保留', '終了'];

// 曜日の表示文字列（0=日曜）
const INGREDIENT_WEEKDAY_LABELS = ['日', '月', '火', '水', '木', '金', '土'];

const INGREDIENT_MASTER_SHEET_HEADER = [
  '食材ID', '食材名', '区分', '既定単位',
  '最低残量', '最低残量単位', '期限接近通知', '業者ID', '有効', '更新日時'
];

const INGREDIENT_LOT_SHEET_HEADER = [
  '在庫ID', '食材ID', '食材名', '区分', '数量', '単位', '保存状態',
  'ラベル期限種別', 'ラベル期限', '使用期限',
  '入荷日', '仕込み日時', '仕込み日時不明', '冷凍日時', '冷凍区分',
  '解凍開始日時', '解凍完了日時', '由来在庫ID',
  '確認日時', '確認者', '備考', '行状態', '版数', '更新日時'
];

const INGREDIENT_EVENT_SHEET_HEADER = [
  'イベントID', '操作キー', '記録日時', '種別', '対象在庫ID', '相手在庫ID',
  '食材ID', '数量', '単位', '内容', '実施者', '処理状態'
];

const INGREDIENT_SUPPLIER_SHEET_HEADER = [
  '業者ID', '業者名', '受注締切時刻', '受注営業日', '出荷営業日', '納品営業日',
  '納品リードタイム日数', 'リードタイムの数え方', '臨時休業日', '店舗受取可能日', '備考', '更新日時'
];

const INGREDIENT_ORDER_SHEET_HEADER = [
  '発注ID', '食材ID', '食材名', '業者ID', '数量', '単位',
  '発注日時', '入荷予定日', '状態', '入荷在庫ID', '備考', '更新日時'
];

// 食材マスタを新しく作ったときだけ入れる初期候補（在庫数量・期限は入れない）
const INGREDIENT_DEFAULT_MASTERS = [
  { 'name': '鶏ガラ', 'category': '原材料', 'unit': '袋' },
  { 'name': '卵', 'category': '原材料', 'unit': '個' },
  { 'name': '豚（ミンチ用）', 'category': '原材料', 'unit': 'kg' },
  { 'name': '小松菜', 'category': '原材料', 'unit': '袋' },
  { 'name': 'ニンニク', 'category': '原材料', 'unit': '個' },
  { 'name': 'カルビ', 'category': '原材料', 'unit': 'kg' },
  { 'name': 'ロース', 'category': '原材料', 'unit': 'kg' },
  { 'name': '牛タン', 'category': '原材料', 'unit': 'kg' },
  { 'name': 'シマチョウ', 'category': '原材料', 'unit': 'kg' },
  { 'name': 'スープ', 'category': '仕込み済み', 'unit': '杯' },
  { 'name': '豚ミンチ', 'category': '仕込み済み', 'unit': 'kg' },
  { 'name': 'カット済みの肉', 'category': '仕込み済み', 'unit': '人前' }
];

// ここから下はログイン（パスコード認証）で使う設定値

// 役割の名称。画面の出し分けと、店長専用処理の権限確認の両方でこの値を使う
const ROLE_EMPLOYEE = '従業員';
const ROLE_MANAGER = '店長';

// パスコードと役割の対応。先頭の 0 を保持するため、キーは文字列として扱う
const LOGIN_PASSCODES = {
  '0000': ROLE_EMPLOYEE,
  '9999': ROLE_MANAGER
};

// ログイントークンをキャッシュに保存するときのキーの接頭辞
const AUTH_TOKEN_CACHE_PREFIX = 'auth_token_';

// ログイントークンの有効時間（秒）。CacheService の上限である 6 時間にしている
const AUTH_TOKEN_TTL_SECONDS = 21600;
