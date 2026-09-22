//　アプリ全体で共通して使う定数・設定値をまとめる

const SHEET_NAMES = {
  CHECK_RESULT: 'チェック結果',
  EMPLOYEES: '従業員名簿',
  TIME_CLOCKS: '打刻履歴',
  HOLIDAYS: '祝日',
  SKILLS: 'スキル管理',
  TODAY_WORK: '本日の業務'
};

// 従業員名簿の列番号（1始まり）。A列から順に 従業員番号・氏名・雇用区分の印
const EMPLOYEE_COL_ID = 1;
const EMPLOYEE_COL_NAME = 2;
const EMPLOYEE_COL_PART_TIME_MARK = 3;

// 削除済みかどうかを管理する列の見出し。既存の列は書き換えず、見つからないときだけ名簿の右端に足す
const EMPLOYEE_STATUS_HEADER = '状態';
const EMPLOYEE_DELETED_AT_HEADER = '削除日時';

// 状態列に入れる値。空欄の既存従業員は「有効」として扱うので、手作業でのデータ移行はいらない
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

// スキル評価は記号（◎○△×）をやめ、そのまま読める3段階の文字列でシートに保存する
const SKILL_LEVEL_NOT_YET = '未習得';
const SKILL_LEVEL_LEARNING = '習得中';
const SKILL_LEVEL_MASTERED = '習得済';

// 画面のドロップダウンにもこの順番で並べる（上から順に習得が進む）
const SKILL_LEVELS = [SKILL_LEVEL_NOT_YET, SKILL_LEVEL_LEARNING, SKILL_LEVEL_MASTERED];

// まだ一度も評価していない項目の既定値
const SKILL_LEVEL_DEFAULT = SKILL_LEVEL_NOT_YET;

// 記号で保存されていた既存データの読み替え表。◎と○はどちらも「習得済」にまとめる
// 手入力で入りやすい別字体の丸（〇・◯）も同じ扱いにして、変換のとりこぼしを防ぐ
const SKILL_LEGACY_LEVEL_MAP = {
  '×': SKILL_LEVEL_NOT_YET,
  '✕': SKILL_LEVEL_NOT_YET,
  '△': SKILL_LEVEL_LEARNING,
  '○': SKILL_LEVEL_MASTERED,
  '〇': SKILL_LEVEL_MASTERED,
  '◯': SKILL_LEVEL_MASTERED,
  '◎': SKILL_LEVEL_MASTERED
};

// 従業員名簿でスキル手当（円/時）を保存する列の見出し
// 状態・削除日時と同じく、見出しが見つからないときだけ名簿の右端に足す
const EMPLOYEE_SKILL_ALLOWANCE_HEADER = 'スキル手当';

// スキル手当の段階。「習得済」の数がこの数に達したら、その金額まで手当を上げる
// 100円は「4個到達の +50円」と「6個到達のさらに +50円」を合計した金額
const SKILL_ALLOWANCE_TIERS = [
  { 'masteredCount': 4, 'amount': 50 },
  { 'masteredCount': 6, 'amount': 100 }
];

// まだ一度も条件を満たしていない従業員のスキル手当
const SKILL_ALLOWANCE_DEFAULT = 0;

const TIME_ZONE = 'Asia/Tokyo';

const UNRECORDED_LABEL = '未打刻';

const NO_VALUE_LABEL = '-';

const CLOCK_TYPES = ['出勤', '退勤', '休憩開始', '休憩終了'];

const POSITION_ITEMS = ['ホール', '厨房', '肉の仕込み', '締め作業'];

const TODAY_WORK_SHEET_HEADER = ['日付', '従業員ID', 'ポジション', '確認事項'];

// ここから下は共有メモで使う設定値

// 共有メモ用のシート名（SHEET_NAMES に後から足すことで既存の定義を書き換えない）
SHEET_NAMES.SHARED_NOTES = '共有メモ';

// 「1つの共有メモ × 受信者1人」を1行として保存する。初回表示日は、受信者が出勤して初めて表示するまで空欄
const SHARED_NOTE_SHEET_HEADER = [
  '共有メモID', '送信者従業員番号', '送信者名', '受信者従業員番号', '受信者名', '共有メモ本文', '作成日時', '初回表示日'
];

// ここから下は発注管理で使う食材・業者のマスタ設定値
// 食材の名前・発注単位・紐づく業者は、これらのシートで管理する

// 発注管理で使うシート名（SHEET_NAMES に後から足すことで既存の定義を書き換えない）
SHEET_NAMES.INGREDIENT_MASTERS = '食材マスタ';
SHEET_NAMES.INGREDIENT_EVENTS = '食材在庫履歴';
SHEET_NAMES.INGREDIENT_SUPPLIERS = '食材業者';
SHEET_NAMES.INGREDIENT_ORDERS = '食材発注';

// 食材の区分。発注できるのは '原材料' だけで、'仕込み済み' は店内で作るものとして発注の対象にしない
const INGREDIENT_CATEGORIES = ['原材料', '仕込み済み'];

// 納品リードタイムの数え方
const INGREDIENT_LEAD_TIME_COUNTS = ['営業日', '暦日'];

// 曜日の表示文字列（0=日曜）
const INGREDIENT_WEEKDAY_LABELS = ['日', '月', '火', '水', '木', '金', '土'];

const INGREDIENT_MASTER_SHEET_HEADER = [
  '食材ID', '食材名', '区分', '既定単位',
  '最低残量', '最低残量単位', '期限接近通知', '業者ID', '有効', '更新日時',
  // ここから右は発注管理で足した列。既存の列は並びを変えていないので、これまでの記録はそのまま読める
  '発注単位'
];

const INGREDIENT_EVENT_SHEET_HEADER = [
  'イベントID', '操作キー', '記録日時', '種別', '対象在庫ID', '相手在庫ID',
  '食材ID', '数量', '単位', '内容', '実施者', '処理状態'
];

const INGREDIENT_SUPPLIER_SHEET_HEADER = [
  '業者ID', '業者名', '受注締切時刻', '受注営業日', '出荷営業日', '納品営業日',
  '納品リードタイム日数', 'リードタイムの数え方', '臨時休業日', '店舗受取可能日', '備考', '更新日時',
  // ここから右は発注管理で足した列（連絡先とメール作成のルール）。既存の列は並びを変えていない
  '住所', '電話番号', 'メールアドレス', '営業時間', '定休日', '主な取扱商品', '発注時の注意事項'
];

const INGREDIENT_ORDER_SHEET_HEADER = [
  '発注ID', '食材ID', '食材名', '業者ID', '数量', '単位',
  '発注日時', '入荷予定日', '状態', '入荷在庫ID', '備考', '更新日時',
  // ここから右は発注管理で足した列。1通の発注メールにまとめた食材は同じ発注番号になる
  '発注番号', '発注者', '送信先メール'
];

// 食材マスタを新しく作ったときだけ入れる初期候補
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
  '0': ROLE_MANAGER
};

// ログイントークンをキャッシュに保存するときのキーの接頭辞
const AUTH_TOKEN_CACHE_PREFIX = 'auth_token_';

// ログイントークンの有効時間（秒）。CacheService の上限である 6 時間にしている
const AUTH_TOKEN_TTL_SECONDS = 21600;

// ここから下は発注管理（電話・FAXでの発注をアプリにまとめる）で使う設定値
// 業者・食材のマスタは上で定義した「食材マスタ・食材業者」シートをそのまま使い、
// 連絡先やメール作成のルールなど、発注に必要な列だけを右に足している

// 発注の記録は「食材発注」シートに残す（発注履歴として画面に出すのもこのシート）
SHEET_NAMES.ORDER_HISTORY = SHEET_NAMES.INGREDIENT_ORDERS;

// 発注メールの差出人にしたいGoogleアカウント
// Gmailの別名（エイリアス）に登録されていればこの名前で送り、無ければスクリプトを実行している本人のアドレスで送る
const ORDER_SENDER_EMAIL = 'uu.kawamura@gmail.com';

// 動作確認中の送信先。値が入っている間は、業者マスタのメールアドレスではなくここへ送る
// 本番の業者へ送るようにするときは空文字にする
const ORDER_TEST_RECIPIENT_EMAIL = 'yu150516@icloud.com';

// メールの署名と件名に使う店舗の情報
const ORDER_STORE_NAME = 'freee食堂';
const ORDER_STORE_PHONE = '03-0000-0000';

// 1回の発注で受け付ける数量の上限（入力ミスで極端な値が入るのを防ぐ）
const ORDER_MAX_QUANTITY = 999;

// 業者の「発注時の注意事項」に書かれていたら、メールの作り方を変える合図にする言葉
const ORDER_MAIL_RULE_KEYWORDS = {
  'DELIVERY_DATE': '納品希望日',
  'SUBJECT_STORE_NAME': '件名',
  'STORE_NAME': '店舗名'
};

// 注意事項の行がこの言葉で始まっていたら、その行の残りをメール本文にそのまま入れる
// （「特定の文章を入れる」というルールを、業者ごとの設定だけで足せるようにするための決まり）
const ORDER_MAIL_EXTRA_TEXT_PREFIX = '本文追記：';

// 業者マスタが空のときだけ入れる初期データ
// メールアドレスは動作確認用の宛先にしてあるので、実際の業者に送る前に業者マスタを書き換える
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
    'name': '○○青果',
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
    'orderNote': '納品希望日を記載',
    'note': '箱のまま納品されます'
  },
  {
    'name': '△△食品',
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
    'orderNote': '件名に店舗名を入れる\n本文追記：伝票は商品に同梱してください',
    'note': '土日は受注できません'
  }
];

// 食材と業者の結び付け（食材マスタの業者IDが空欄のときだけ入れる初期値）
// 発注単位は食材マスタの既定単位と別に持つので、ここで指定した単位を発注画面で使う
const ORDER_DEFAULT_ITEMS = [
  { 'name': '牛タン', 'vendorName': 'freee精肉店', 'unit': 'kg' },
  { 'name': 'カルビ', 'vendorName': 'freee精肉店', 'unit': 'kg' },
  { 'name': 'ロース', 'vendorName': 'freee精肉店', 'unit': 'kg' },
  { 'name': 'シマチョウ', 'vendorName': 'freee精肉店', 'unit': 'kg' },
  { 'name': '豚（ミンチ用）', 'vendorName': 'freee精肉店', 'unit': 'kg' },
  { 'name': '豚ロース', 'vendorName': 'freee精肉店', 'unit': 'kg' },
  { 'name': '小松菜', 'vendorName': '○○青果', 'unit': '袋' },
  { 'name': '玉ねぎ', 'vendorName': '○○青果', 'unit': 'ケース' },
  { 'name': 'ニンニク', 'vendorName': '○○青果', 'unit': '個' },
  { 'name': 'キャベツ', 'vendorName': '○○青果', 'unit': 'ケース' },
  { 'name': '鶏ガラ', 'vendorName': '△△食品', 'unit': '袋' },
  { 'name': '卵', 'vendorName': '△△食品', 'unit': 'ケース' }
];

// 初期データで足りない食材は食材マスタに追加する（区分は原材料・既定単位は発注単位と同じにする）
const ORDER_DEFAULT_ITEM_CATEGORY = '原材料';
