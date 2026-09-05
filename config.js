var SHEET_NAMES = {
  CHECK_RESULT: 'チェック結果',
  EMPLOYEES: '従業員名簿',
  TIME_CLOCKS: '打刻履歴',
  HOLIDAYS: '祝日',
  SKILLS: 'スキル管理',
  TODAY_WORK: '本日の業務'
};

var SKILL_ITEMS = ['接客', 'レジ', 'バー', '厨房', '肉の仕込み', '締め作業'];

var SKILL_SHEET_HEADER = [
  '従業員ID', '氏名',
  '接客評価', '更新日',
  'レジ評価', '更新日',
  'バー評価', '更新日',
  '厨房評価', '厨房更新日',
  '肉の仕込み評価', '更新日',
  '締め作業評価', '更新日'
];

var SKILL_LEVELS = ['◎', '○', '△', '×'];

var TIME_ZONE = 'Asia/Tokyo';

var UNRECORDED_LABEL = '未打刻';

var NO_VALUE_LABEL = '-';

var CLOCK_TYPES = ['出勤', '退勤', '休憩開始', '休憩終了'];

var POSITION_ITEMS = ['ホール', '厨房', '肉の仕込み', '締め作業'];

var TODAY_WORK_SHEET_HEADER = ['日付', '従業員ID', 'ポジション', '確認事項'];