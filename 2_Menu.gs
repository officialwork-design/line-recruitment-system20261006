/**
 * 2_Menu.gs
 * スプレッドシート上部メニュー・管理画面・シート移動
 */

/**
 * スプレッドシートを開いた時のメニュー
 */
function onOpen() {
  SpreadsheetApp
    .getUi()
    .createMenu('求人管理')
    .addItem('初期セットアップを実行', 'setupRecruitSpreadsheet')
    .addSeparator()

    .addItem('対応管理を開く', 'openSupportSheet')
    .addItem('応募管理を見る', 'openApplicationSheet')
    .addItem('連絡先を見る', 'openContactSheet')
    .addItem('ユーザー管理を見る', 'openUserManagementSheet')
    .addSeparator()

    .addItem('対応管理同期トリガーを設定', 'installSupportApplicationSyncTrigger')
    .addItem('応募管理から対応管理へ手動同期', 'syncApplicationsToSupportManual')
    .addItem('過去の面接/体入/本入履歴を反映', 'backfillSupportActionHistory')
    .addSeparator()

    .addItem('応募Noを指定して再メンション', 'resendApplicationMentionByPrompt')
    .addItem('連絡先にLINE送信', 'sendLineMessageToContactsByPrompt')
    .addSeparator()

    .addItem('未応募者追客トリガーを設定', 'installUnappliedFollowupTrigger')
    .addItem('未応募者追客を手動実行', 'sendUnappliedFollowupMessages')
    .addSeparator()

    .addItem('設定USERIDへ応募ボタン送信', 'sendApplyButtonToConfiguredUser')
    .addItem('USERID入力で応募ボタン送信', 'sendApplyButtonToUser')
    .addItem('選択ユーザーへ応募ボタン送信', 'sendApplyButtonToSelectedUser')
    .addSeparator()

    .addItem('連絡先シートだけ作成', 'setupContactSheetOnly')
    .addItem('対応完了の色付けを設定', 'setupSupportStatusConditionalFormat')
    .addItem('キャッシュをクリア', 'clearSystemCache')
    .addSeparator()

    .addItem('LINEグループID確認ログON', 'enableGroupIdLog')
    .addItem('LINEグループID確認ログOFF', 'disableGroupIdLog')
    .addSeparator()

    .addItem('リッチメニュー設定作成', 'setupRichMenuConfigSheet')
    .addItem('リッチメニュー作成', 'createApplyRichMenuFromSheet')
    .addItem('リッチメニュー一覧', 'listRichMenus')
    .addToUi();
}

/**
 * HTML include
 */
function include(filename) {
  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();
}

/**
 * 管理画面を開く
 */
function openAdminUi() {
  const html = HtmlService
    .createTemplateFromFile('ui')
    .evaluate()
    .setWidth(1000)
    .setHeight(760);

  SpreadsheetApp
    .getUi()
    .showModalDialog(html, '求人管理');
}

/**
 * 各シートを開く
 */
function openQuestionSheet() {
  openSheetByName_(SHEET_QUESTIONS);
}

function openChoiceSheet() {
  openSheetByName_(SHEET_CHOICES);
}

function openSupportSheet() {
  openSheetByName_(SHEET_SUPPORT);
}

function openApplicationSheet() {
  openSheetByName_(SHEET_APPLICATIONS);
}

function openContactSheet() {
  openSheetByName_(SHEET_CONTACTS);
}

function openUserManagementSheet() {
  openSheetByName_(SHEET_USER_MANAGEMENT);
}

function openConfigSheet() {
  openSheetByName_(SHEET_CONFIG);
}

function openSpecSheet() {
  openSheetByName_('仕様書');
}

function openOperationMemoSheet() {
  openSheetByName_('運用メモ');
}

function openErrorLogSheet() {
  openSheetByName_(SHEET_ERROR_LOG);
}

function openProcessLogSheet() {
  openSheetByName_(SHEET_PROCESS_LOG);
}

/**
 * 指定シートを表示してアクティブ化
 */
function openSheetByName_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    SpreadsheetApp
      .getUi()
      .alert(
        `シートが見つかりません: ${sheetName}\n\n先に「初期セットアップを実行」してください。`
      );
    return;
  }

  try {
    sheet.showSheet();
  } catch (error) {
    // 非表示解除できない場合も処理継続
  }

  ss.setActiveSheet(sheet);
}

/**
 * 互換用
 * 旧コード側で activateSheetByName を呼んでいても動くように残す
 */
function activateSheetByName(sheetName) {
  openSheetByName_(sheetName);
}

/**
 * 管理画面用データ取得
 */
function getAdminDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  return {
    spreadsheetName: ss.getName(),
    activeSheetName: ss.getActiveSheet().getName(),
    startWord: getConfigValue_('START_WORD') || DEFAULT_START_WORD,
    applyButtonUserId: getConfigValue_('APPLY_BUTTON_USER_ID') || '',
    adminGroupId: getConfigValue_(CONFIG_ADMIN_GROUP_ID) || '',
    questionCount: typeof getBaseApplicationSteps_ === 'function'
      ? getBaseApplicationSteps_()
      : 0,
    maxQuestionCount: typeof getMaxApplicationSteps_ === 'function'
      ? getMaxApplicationSteps_()
      : 0,
    applicationCount: getApplicationCount_(),
    supportCount: getSupportCount_(),
    lastError: getLastLogText_(SHEET_ERROR_LOG),
    lastProcess: ENABLE_PROCESS_LOG
      ? getLastLogText_(SHEET_PROCESS_LOG)
      : '通常ログは停止中'
  };
}

/**
 * 応募数
 */
function getApplicationCount_() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_APPLICATIONS);

  if (!sheet) return 0;

  return Math.max(sheet.getLastRow() - 2, 0);
}

/**
 * 対応管理件数
 */
function getSupportCount_() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_SUPPORT);

  if (!sheet) return 0;

  return Math.max(sheet.getLastRow() - 2, 0);
}

/**
 * 最新ログ取得
 */
function getLastLogText_(sheetName) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(sheetName);

  if (!sheet || sheet.getLastRow() < 2) return '';

  const row = sheet
    .getRange(
      sheet.getLastRow(),
      1,
      1,
      Math.min(sheet.getLastColumn(), 5)
    )
    .getValues()[0];

  return row
    .map(value => {
      if (value instanceof Date) {
        return Utilities.formatDate(
          value,
          'Asia/Tokyo',
          'yyyy/MM/dd HH:mm:ss'
        );
      }

      return String(value || '');
    })
    .join(' / ');
}

/**
 * 連絡先シートだけ作成
 */
function setupContactSheetOnly() {
  setupContactSheetForRecruit_(
    SpreadsheetApp.getActiveSpreadsheet()
  );

  SpreadsheetApp
    .getUi()
    .alert('連絡先シートを作成しました。');
}

/**
 * キャッシュクリア
 */
function clearSystemCache() {
  CacheService
    .getScriptCache()
    .removeAll([
      'activeQuestions',
      'allChoices',
      'config:LINE_CHANNEL_ACCESS_TOKEN',
      'config:START_WORD',
      'config:COMPLETE_MESSAGE',
      'config:APPLY_BUTTON_USER_ID',
      'config:APPLY_BUTTON_TITLE',
      'config:APPLY_BUTTON_NOTE',
      'config:APPLY_BUTTON_LABEL',
      'config:APPLY_BUTTON_COLOR',
      `config:${CONFIG_ADMIN_GROUP_ID}`
    ]);

  SpreadsheetApp
    .getUi()
    .alert('キャッシュをクリアしました。');
}

/**
 * LINEグループID確認ログ ON
 */
function enableGroupIdLog() {
  PropertiesService
    .getScriptProperties()
    .setProperty('ENABLE_GROUP_ID_LOG', 'TRUE');

  SpreadsheetApp
    .getUi()
    .alert(
      [
        'LINEグループID確認ログをONにしました。',
        '',
        '通知したいグループで何かメッセージを送ってください。',
        'エラーログに LINE_SOURCE_ID が出ます。'
      ].join('\n')
    );
}

/**
 * LINEグループID確認ログ OFF
 */
function disableGroupIdLog() {
  PropertiesService
    .getScriptProperties()
    .setProperty('ENABLE_GROUP_ID_LOG', 'FALSE');

  SpreadsheetApp
    .getUi()
    .alert('LINEグループID確認ログをOFFにしました。');
}