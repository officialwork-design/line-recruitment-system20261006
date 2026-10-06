/**
 * 1_Config.gs
 * 定数・設定値
 */

/**
 * シート名
 */
const SHEET_CONFIG = '設定';
const SHEET_QUESTIONS = '質問設定';
const SHEET_CHOICES = '選択肢設定';
const SHEET_APPLICATIONS = '応募管理';
const SHEET_SUPPORT = '対応管理';
const SHEET_SUPPORT_HISTORY = '対応履歴';
const SHEET_CONTACTS = '連絡先';
const SHEET_USER_MANAGEMENT = 'ユーザー管理';
const SHEET_PROCESS_LOG = '処理ログ';
const SHEET_ERROR_LOG = 'エラーログ';
const SHEET_RICH_MENU = 'リッチメニュー設定';

/**
 * 設定キー
 */
const CONFIG_ADMIN_GROUP_ID = 'ADMIN_GROUP_ID';

/**
 * 応募開始ワード
 */
const DEFAULT_START_WORD = '応募';
const RESTART_WORD = '再応募';

/**
 * 応募フロー状態
 */
const STATUS_WAIT_MEDIA = '募集媒体待ち';
const STATUS_WAIT_MEDIA_OTHER = 'その他媒体待ち';
const STATUS_WAITING = '回答待ち';
const STATUS_WAIT_FACE_PHOTO = '顔写真待ち';
const STATUS_WAIT_FULL_BODY_PHOTO = '全体写真待ち';
const STATUS_WAIT_REMARKS = '備考欄待ち';
const STATUS_WAIT_INQUIRY = '問い合わせ待ち';
const STATUS_DONE = '完了';
const STATUS_CANCEL = 'キャンセル';

/**
 * 対応管理ステータス
 */
const SUPPORT_STATUS_NOT_STARTED = '未対応';
const SUPPORT_STATUS_IN_PROGRESS = '対応中';
const SUPPORT_STATUS_DONE = '対応完了';

/**
 * 募集媒体
 */
const MEDIA_CHOICES = [
  'カフェるん',
  'Instagram広告',
  'X',
  'その他'
];

/**
 * その他媒体
 */
const MEDIA_OTHER_CHOICES = [
  'YouTube',
  'TikTok',
  'ポケパラ',
  '紹介'
];

/**
 * 写真関連
 */
const FACE_PHOTO_MESSAGE = '📸 顔が写ったお写真を１枚お送りください。';
const FULL_BODY_PHOTO_MESSAGE = '📸 全身が写ったお写真を1枚お送りください。';
const PHOTO_FILE_NOTE = '対応形式：jpg / jpeg / png';
const PHOTO_RECEIVED_TEXT = '受信済み';

/**
 * キャッシュ
 */
const CACHE_SECONDS = 300;

/**
 * 処理ログON/OFF
 */
const ENABLE_PROCESS_LOG = true;

/**
 * 処理ログに残す重要ログ
 */
const PROCESS_LOG_ALLOWED_TYPES = [
  'FOLLOW',
  'QUESTION_1_SENT',
  'QUESTION_1_ANSWER',
  'APPLICATION_SAVED',
  'APPLICATION_SAVE_ERROR'
];