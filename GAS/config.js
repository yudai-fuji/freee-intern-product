const SHEET_NAMES = {
  CHECK_RESULT: 'チェック結果',
  EMPLOYEES: '従業員名簿',
  TIME_CLOCKS: '打刻履歴',
  HOLIDAYS: '祝日',
  SKILLS: 'スキル管理',
  TODAY_WORK: '本日の業務'
};

const EMPLOYEE_COL_ID = 1;
const EMPLOYEE_COL_NAME = 2;
const EMPLOYEE_COL_PART_TIME_MARK = 3;

const EMPLOYEE_STATUS_HEADER = '状態';
const EMPLOYEE_DELETED_AT_HEADER = '削除日時';

const EMPLOYEE_STATUS_ACTIVE = '有効';
const EMPLOYEE_STATUS_DELETED = '削除済み';

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

const SKILL_LEVEL_NOT_YET = 'F';
const SKILL_LEVEL_LEARNING = 'B';
const SKILL_LEVEL_MASTERED = 'A';

const SKILL_LEVELS = [SKILL_LEVEL_NOT_YET, SKILL_LEVEL_LEARNING, SKILL_LEVEL_MASTERED];
const SKILL_LEVEL_LABELS = { F: '未経験', B: '習得中', A: '習得済み' };

const SKILL_LEVEL_DEFAULT = SKILL_LEVEL_NOT_YET;

const SKILL_LEGACY_LEVEL_MAP = {
  '未習得': SKILL_LEVEL_NOT_YET,
  '習得中': SKILL_LEVEL_LEARNING,
  '習得済': SKILL_LEVEL_MASTERED,
  '×': SKILL_LEVEL_NOT_YET,
  '✕': SKILL_LEVEL_NOT_YET,
  '△': SKILL_LEVEL_LEARNING,
  '○': SKILL_LEVEL_MASTERED,
  '〇': SKILL_LEVEL_MASTERED,
  '◯': SKILL_LEVEL_MASTERED,
  '◎': SKILL_LEVEL_MASTERED
};

const EMPLOYEE_SKILL_ALLOWANCE_HEADER = 'スキル手当';

const BASE_HOURLY_WAGE = 1100;
const HOLIDAY_HOURLY_BONUS = 50;

const SKILL_ALLOWANCE_TIERS = [
  { 'masteredCount': 4, 'amount': 50 },
  { 'masteredCount': 6, 'amount': 100 }
];

const SKILL_ALLOWANCE_DEFAULT = 0;

const TIME_ZONE = 'Asia/Tokyo';

const UNRECORDED_LABEL = '未打刻';

const NO_VALUE_LABEL = '-';

const CLOCK_TYPES = ['出勤', '退勤', '休憩開始', '休憩終了'];

const ATTENDANCE_BREAK_ADJUSTMENT_HEADER = '休憩補正（分）';

const POSITION_ITEMS = ['ホール', '厨房', '肉の仕込み', '締め作業'];

const TODAY_WORK_SHEET_HEADER = ['日付', '従業員ID', 'ポジション', '確認事項'];

SHEET_NAMES.SHARED_NOTES = '共有メモ';

const SHARED_NOTE_SHEET_HEADER = [
  '共有メモID', '送信者従業員番号', '送信者名', '受信者従業員番号', '受信者名', '共有メモ本文', '作成日時', '初回表示日'
];

SHEET_NAMES.INGREDIENT_MASTERS = '食材マスタ';
SHEET_NAMES.INGREDIENT_EVENTS = '食材在庫履歴';
SHEET_NAMES.INGREDIENT_SUPPLIERS = '食材業者';
SHEET_NAMES.INGREDIENT_ORDERS = '食材発注';

const INGREDIENT_CATEGORIES = ['原材料', '仕込み済み'];

const INGREDIENT_LEAD_TIME_COUNTS = ['営業日', '暦日'];

const INGREDIENT_WEEKDAY_LABELS = ['日', '月', '火', '水', '木', '金', '土'];

const INGREDIENT_MASTER_SHEET_HEADER = [
  '食材ID', '食材名', '区分', '既定単位',
  '最低残量', '最低残量単位', '期限接近通知', '業者ID', '有効', '更新日時',

  '発注単位'
];

const INGREDIENT_EVENT_SHEET_HEADER = [
  'イベントID', '操作キー', '記録日時', '種別', '対象在庫ID', '相手在庫ID',
  '食材ID', '数量', '単位', '内容', '実施者', '処理状態'
];

const INGREDIENT_SUPPLIER_SHEET_HEADER = [
  '業者ID', '業者名', '受注締切時刻', '受注営業日', '出荷営業日', '納品営業日',
  '納品リードタイム日数', 'リードタイムの数え方', '臨時休業日', '店舗受取可能日', '備考', '更新日時',

  '住所', '電話番号', 'メールアドレス', '営業時間', '定休日', '主な取扱商品', '発注時の注意事項'
];

const INGREDIENT_ORDER_SHEET_HEADER = [
  '発注ID', '食材ID', '食材名', '業者ID', '数量', '単位',
  '発注日時', '入荷予定日', '状態', '入荷在庫ID', '備考', '更新日時',

  '発注番号', '発注者', '送信先メール'
];

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

const ROLE_EMPLOYEE = '従業員';
const ROLE_MANAGER = '店長';

const MANAGER_PASSWORD_PROPERTY = 'manager_password';
const MANAGER_PASSWORD_DEFAULT = '00000000';
const MANAGER_PASSWORD_MIN_LENGTH = 8;
const MANAGER_PASSWORD_MAX_LENGTH = 64;
const EMPLOYEE_PIN_HEADER = '打刻PIN';
const EMPLOYEE_PIN_DEFAULT = '000000';
const APP_DATA_MIGRATION_PROPERTY = 'pin_and_skill_ranks_v1';
const CLOCK_PIN_TICKET_PREFIX = 'clock_pin_';
const CLOCK_PIN_TICKET_TTL_SECONDS = 600;

const AUTH_TOKEN_CACHE_PREFIX = 'auth_token_';

const AUTH_TOKEN_TTL_SECONDS = 21600;

SHEET_NAMES.ORDER_HISTORY = SHEET_NAMES.INGREDIENT_ORDERS;

const ORDER_SENDER_EMAIL = 'uu.kawamura@gmail.com';

const ORDER_TEST_RECIPIENT_EMAIL = 'yu150516@icloud.com';

const ORDER_STORE_NAME = 'ラーメン焼肉freee';
const ORDER_STORE_PHONE = '000-0000-0000';

const ORDER_MAX_QUANTITY = 999;

const ORDER_DEFAULT_VENDORS = [
  {
    'name': 'freee精肉店',
    'address': '東京都港区芝公園1-2-3',
    'phone': '03-1234-5678',
    'email': 'yu150516@icloud.com',
    'businessHours': '8:00～17:00',
    'closedDays': '日曜・祝日',
    'deadlineTime': '15:00',
    'leadTimeDays': '1',
    'leadTimeCount': '営業日',
    'orderDays': ['月', '火', '水', '木', '金', '土'],
    'shipDays': ['月', '火', '水', '木', '金', '土'],
    'deliveryDays': ['月', '火', '水', '木', '金', '土'],
    'pickupDays': ['月', '火', '水', '木', '金', '土'],
    'mainItems': '牛肉・豚肉・鶏肉',
    'orderNote': '',
    'note': '急ぎの場合は電話'
  },
  {
    'name': 'freee青果',
    'address': '東京都大田区東海3-2-1',
    'phone': '03-2345-6789',
    'email': 'yu150516@icloud.com',
    'businessHours': '5:00～14:00',
    'closedDays': '日曜',
    'deadlineTime': '16:00',
    'leadTimeDays': '1',
    'leadTimeCount': '営業日',
    'orderDays': ['月', '火', '水', '木', '金', '土'],
    'shipDays': ['月', '火', '水', '木', '金', '土'],
    'deliveryDays': ['月', '火', '水', '木', '金', '土'],
    'pickupDays': ['月', '火', '水', '木', '金', '土'],
    'mainItems': '野菜・果物',
    'orderNote': '',
    'note': '箱のまま納品されます'
  },
  {
    'name': 'freee食品',
    'address': '東京都江東区豊洲5-6-7',
    'phone': '03-3456-7890',
    'email': 'yu150516@icloud.com',
    'businessHours': '9:00～18:00',
    'closedDays': '土曜・日曜・祝日',
    'deadlineTime': '12:00',
    'leadTimeDays': '2',
    'leadTimeCount': '営業日',
    'orderDays': ['月', '火', '水', '木', '金'],
    'shipDays': ['月', '水', '金'],
    'deliveryDays': ['月', '火', '水', '木', '金'],
    'pickupDays': ['月', '火', '水', '木', '金'],
    'mainItems': '鶏ガラ・卵・調味料',
    'orderNote': '',
    'note': '土日は受注できません'
  }
];

const ORDER_DEFAULT_ITEMS = [
  { 'name': '牛タン', 'vendorName': 'freee精肉店', 'unit': 'kg' },
  { 'name': 'カルビ', 'vendorName': 'freee精肉店', 'unit': 'kg' },
  { 'name': 'ロース', 'vendorName': 'freee精肉店', 'unit': 'kg' },
  { 'name': 'シマチョウ', 'vendorName': 'freee精肉店', 'unit': 'kg' },
  { 'name': '豚（ミンチ用）', 'vendorName': 'freee精肉店', 'unit': 'kg' },
  { 'name': '豚ロース', 'vendorName': 'freee精肉店', 'unit': 'kg' },
  { 'name': '小松菜', 'vendorName': 'freee青果', 'unit': '袋' },
  { 'name': '玉ねぎ', 'vendorName': 'freee青果', 'unit': 'ケース' },
  { 'name': 'ニンニク', 'vendorName': 'freee青果', 'unit': '個' },
  { 'name': 'キャベツ', 'vendorName': 'freee青果', 'unit': 'ケース' },
  { 'name': '鶏ガラ', 'vendorName': 'freee食品', 'unit': '袋' },
  { 'name': '卵', 'vendorName': 'freee食品', 'unit': 'ケース' }
];

const ORDER_DEFAULT_ITEM_CATEGORY = '原材料';
