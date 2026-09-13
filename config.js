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