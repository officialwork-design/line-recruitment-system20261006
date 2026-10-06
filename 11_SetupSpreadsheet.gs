/**
 * 11_SetupSpreadsheet.gs
 * 求人管理スプレッドシート初期セットアップ
 *
 * 方針：
 * - 新規ファイルは作らない
 * - 既存シートは削除しない
 * - 既存データは残す
 * - ヘッダー、メモ、入力規則、表示形式は更新する
 * - 本番誤操作防止のため、実行前に確認ダイアログを出す
 */

function setupRecruitSpreadsheet() {
  const ui = SpreadsheetApp.getUi();

  const result = ui.alert(
    '初期セットアップ確認',
    [
      '求人管理システムの初期セットアップを実行します。',
      '',
      '既存シートは削除しません。',
      'ただし、ヘッダー行・説明行・入力規則・表示形式は更新されます。',
      '',
      '本番運用中の場合は、念のため事前にスプレッドシートをコピーしてバックアップしてください。',
      '',
      '実行しますか？'
    ].join('\n'),
    ui.ButtonSet.OK_CANCEL
  );

  if (result !== ui.Button.OK) {
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  setupConfigSheetForRecruit_(ss);
  setupQuestionSheetForRecruit_(ss);
  setupChoiceSheetForRecruit_(ss);
  setupApplicationSheetForRecruit_(ss);
  setupSupportSheetForRecruit_(ss);
  setupContactSheetForRecruit_(ss);
  setupSupportHistorySheetForRecruit_(ss);
  setupUserManagementSheetForRecruit_(ss);
  setupErrorLogSheetForRecruit_(ss);
  setupProcessLogSheetForRecruit_(ss);
  setupSpecSheetForRecruit_(ss);
  setupOperationMemoSheetForRecruit_(ss);

  try {
    setupSupportStatusConditionalFormat();
  } catch (error) {
    errorLog_(
      'SUPPORT_STATUS_FORMAT_SETUP_ERROR',
      error.stack || error.message,
      '',
      ''
    );
  }

  try {
    setupRichMenuConfigSheet();
  } catch (error) {
    errorLog_(
      'RICH_MENU_CONFIG_SETUP_SKIPPED',
      error.stack || error.message,
      '',
      'リッチメニュー関数が未定義の場合は無視できます。'
    );
  }

  try {
    installSupportApplicationSyncTrigger();
  } catch (error) {
    errorLog_(
      'SYNC_TRIGGER_SETUP_ERROR',
      error.stack || error.message,
      '',
      ''
    );
  }

  ui.alert('初期セットアップが完了しました。');

  Logger.log('初期セットアップ完了');
}

function setupConfigSheetForRecruit_(ss) {
  const sheet = recreateSheet_(ss, SHEET_CONFIG);

  const rows = [
    ['キー', '値', '説明', '備考'],
    ['LINE_CHANNEL_ACCESS_TOKEN', 'ここにLINEチャネルアクセストークンを入力', 'LINE Messaging API チャネルアクセストークン', '必須'],
    ['START_WORD', DEFAULT_START_WORD, '応募開始ワード', '完全一致のみ。誤爆防止のため「応募開始」推奨 / 必須'],
    ['COMPLETE_MESSAGE', '✅ ご応募ありがとうございました！\n\n内容を確認後、担当者よりご連絡いたします。', '応募完了時の自動返信', ''],
    ['APPLY_BUTTON_USER_ID', 'ここに送信先のLINEユーザーIDを入力', '応募ボタン送信用USERID', '互換用'],
    ['APPLY_BUTTON_TITLE', '✨ ご応募はこちらからお願いします。', '応募ボタンタイトル', '互換用'],
    ['APPLY_BUTTON_NOTE', '下のボタンから応募を開始できます。', '応募ボタン補足', '互換用'],
    ['APPLY_BUTTON_LABEL', '応募する', '応募ボタンラベル', '互換用'],
    ['APPLY_BUTTON_COLOR', '#06C755', '応募ボタン色', '互換用'],
    [CONFIG_ADMIN_GROUP_ID, '', '管理者通知先LINEグループID', '必須']
  ];

  sheet
    .getRange(1, 1, rows.length, rows[0].length)
    .setValues(rows);

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, rows[0].length);
}

function setupQuestionSheetForRecruit_(ss) {
  const sheet = recreateSheet_(ss, SHEET_QUESTIONS);

  const headers = [
    '順番',
    '項目キー',
    '項目名',
    '質問文',
    '回答形式',
    '必須',
    '保存列名',
    '有効'
  ];

  const rows = [
    headers,
    [1, 'store', '希望店舗', '希望店舗を選択してください。', '選択ボタン', true, '希望店舗', true],
    [2, 'work_days', '勤務日数', '勤務できる日数を選択してください。', '選択ボタン', true, '勤務日数', true],
    [3, 'age', '年齢', '現時点で何歳ですか？', '文字入力', true, '年齢', true],
    [4, 'remarks', '備考欄', '最後に、補足や質問があれば入力してください。\n特になければ「なし」と入力してください。', '文字入力', false, '備考欄', true]
  ];

  sheet
    .getRange(1, 1, rows.length, headers.length)
    .setValues(rows);

  applyDropdown_(sheet, 'E2:E200', ['文字入力', '選択ボタン']);
  applyDropdown_(sheet, 'F2:F200', [true, false]);
  applyDropdown_(sheet, 'H2:H200', [true, false]);

  sheet.setFrozenRows(1);
  sheet.getRange('A:H').setWrap(true);
  sheet.autoResizeColumns(1, headers.length);
}

function setupChoiceSheetForRecruit_(ss) {
  const sheet = recreateSheet_(ss, SHEET_CHOICES);

  const headers = [
    '項目キー',
    '選択肢名',
    '送信テキスト',
    '表示順',
    '有効'
  ];

  const rows = [
    headers,
    ['store', '大阪店', '大阪店', 1, true],
    ['store', '新宿店', '新宿店', 2, true],
    ['store', '秋葉原店', '秋葉原店', 3, true],
    ['store', '複数店(新宿店/秋葉原店)', '複数店(新宿店/秋葉原店)', 4, true],
    ['work_days', '週1〜2日', '週1〜2日', 1, true],
    ['work_days', '週2〜3日', '週2〜3日', 2, true],
    ['work_days', '週4〜5日', '週4〜5日', 3, true],
    ['work_days', '相談したい', '相談したい', 4, true]
  ];

  sheet
    .getRange(1, 1, rows.length, headers.length)
    .setValues(rows);

  applyDropdown_(sheet, 'E2:E200', [true, false]);

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

function setupApplicationSheetForRecruit_(ss) {
  const sheet = recreateSheet_(ss, SHEET_APPLICATIONS);

  const headers = [
    'No',
    '受信日時',
    'ユーザーID',
    'LINE表示名',
    '募集媒体',
    '募集媒体大分類',
    'その他媒体',
    '希望店舗',
    '勤務日数',
    '年齢',
    '備考欄',
    '応募メッセージ',
    '写真受信数',
    '顔写真確認',
    '全体写真確認',
    '追加回答',
    '面接担当',
    '名前',
    '合否',
    '面接',
    '体入',
    '本入',
    'ステータス',
    '対応管理反映',
    '過去応募者',
    '応募回数',
    '更新日'
  ];

  const notes = [
    '自動採番',
    '自動',
    '自動',
    '自動',
    '自動',
    'その他選択時のみ',
    'その他選択時のみ',
    '自動',
    '自動',
    '自動',
    '自動',
    '自動まとめ',
    '自動',
    '受信済み/未受信',
    '受信済み/未受信',
    '自動JSON',
    '対応管理から連携',
    '対応管理から連携',
    '対応管理から連携',
    '対応管理から連携',
    '対応管理から連携',
    '対応管理から連携',
    '完了',
    '済/未',
    '初回/再応募',
    '自動',
    '自動'
  ];

  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  sheet
    .getRange(2, 1, 1, notes.length)
    .setValues([notes]);

  sheet.setFrozenRows(2);
  sheet.getRange('A:AA').setWrap(true);
  sheet.getRange('B:B').setNumberFormat('yyyy/mm/dd hh:mm:ss');
  sheet.getRange('AA:AA').setNumberFormat('yyyy/mm/dd hh:mm:ss');

  applyDropdown_(sheet, 'N3:N1000', ['受信済み', '未受信']);
  applyDropdown_(sheet, 'O3:O1000', ['受信済み', '未受信']);
  applyDropdown_(sheet, 'S3:S1000', ['', '未設定', '合格', '不合格', '保留']);
  applyDropdown_(sheet, 'T3:V1000', ['', '○', '×']);
  applyDropdown_(sheet, 'W3:W1000', [STATUS_DONE]);
  applyDropdown_(sheet, 'X3:X1000', ['済', '未']);
  applyDropdown_(sheet, 'Y3:Y1000', ['初回', '再応募']);

  sheet.autoResizeColumns(1, headers.length);
}

function setupSupportSheetForRecruit_(ss) {
  const sheet = recreateSheet_(ss, SHEET_SUPPORT);

  const headers = [
    'No',
    'ステータス',
    '面接担当',
    'LINE表示名',
    '応募メッセージ',
    '名前',
    '合否',
    '面接',
    '体入',
    '本入',
    '対応メモ',
    '過去応募者',
    '応募回数',
    '更新日'
  ];

  const notes = [
    '自動',
    '未対応/対応中/対応完了',
    '手動プルダウン',
    '応募管理から',
    '応募管理から',
    '手入力',
    'プルダウン',
    '○/×',
    '○/×',
    '○/×',
    '自由記入',
    '初回/再応募',
    '自動',
    '自動'
  ];

  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  sheet
    .getRange(2, 1, 1, notes.length)
    .setValues([notes]);

  sheet.setFrozenRows(2);
  sheet.getRange('A:N').setWrap(true);
  sheet.getRange('N:N').setNumberFormat('yyyy/mm/dd hh:mm:ss');

  applyDropdown_(sheet, 'B3:B1000', [
    SUPPORT_STATUS_NOT_STARTED,
    SUPPORT_STATUS_IN_PROGRESS,
    SUPPORT_STATUS_DONE
  ]);

  applyDropdown_(sheet, 'C3:C1000', [
    'OSAKA MG',
    '飯塚',
    'ひろ',
    'りゅうじ',
    'りゅうき'
  ]);

  applyDropdown_(sheet, 'G3:G1000', [
    '',
    '未設定',
    '合格',
    '不合格',
    '保留'
  ]);

  applyDropdown_(sheet, 'H3:J1000', [
    '',
    '○',
    '×'
  ]);

  sheet.autoResizeColumns(1, headers.length);
}

function setupContactSheetForRecruit_(ss) {
  const sheet = recreateSheet_(ss, SHEET_CONTACTS);

  const headers = [
    'ユーザーID',
    'LINE表示名',
    '友だち追加日時',
    '最終やりとり日時',
    '状態',
    '応募状況',
    '追客ステータス',
    '最新応募No',
    'メモ',
    '更新日'
  ];

  const notes = [
    'LINE userId',
    '取得できた場合のみ',
    'follow時に自動',
    'メッセージ受信時に自動',
    '友だち追加済み/やりとりあり/応募中/応募完了/ブロック不明',
    '未応募/応募中/応募完了',
    '未開始/2日送信済/5日送信済/7日送信済/停止/送信失敗',
    '応募完了時に自動',
    '自由記入',
    '自動'
  ];

  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  sheet
    .getRange(2, 1, 1, notes.length)
    .setValues([notes]);

  sheet.setFrozenRows(2);
  sheet.getRange('A:J').setWrap(true);
  sheet.getRange('C:D').setNumberFormat('yyyy/mm/dd hh:mm:ss');
  sheet.getRange('J:J').setNumberFormat('yyyy/mm/dd hh:mm:ss');

  applyDropdown_(sheet, 'E3:E1000', [
    '友だち追加済み',
    'やりとりあり',
    '応募中',
    '応募完了',
    'ブロック不明'
  ]);

  applyDropdown_(sheet, 'F3:F1000', [
    '未応募',
    '応募中',
    '応募完了'
  ]);

  applyDropdown_(sheet, 'G3:G1000', [
    '未開始',
    '2日送信済',
    '5日送信済',
    '7日送信済',
    '停止',
    '送信失敗'
  ]);

  sheet.autoResizeColumns(1, headers.length);
}

function setupSupportHistorySheetForRecruit_(ss) {
  const sheet = recreateSheet_(ss, SHEET_SUPPORT_HISTORY);

  const headers = [
    '受信日時',
    'ユーザーID',
    'LINE表示名',
    '種別',
    'メッセージ内容',
    '対応管理No',
    '関連応募No',
    '更新日'
  ];

  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  sheet.setFrozenRows(1);
  sheet.getRange('A:H').setWrap(true);
  sheet.getRange('A:A').setNumberFormat('yyyy/mm/dd hh:mm:ss');
  sheet.getRange('H:H').setNumberFormat('yyyy/mm/dd hh:mm:ss');

  try {
    sheet.hideSheet();
  } catch (error) {}
}

function setupUserManagementSheetForRecruit_(ss) {
  const sheet = recreateSheet_(ss, SHEET_USER_MANAGEMENT);

  const headers = [
    'ユーザーID',
    'LINE表示名',
    '現在の質問番号',
    'ステータス',
    'お名前',
    '希望店舗',
    '写真受信数',
    '顔写真状態',
    '全身写真状態',
    '追加回答JSON',
    '更新日',
    '応募メッセージ',
    '年齢',
    '募集媒体',
    'その他媒体フロー',
    '総質問数',
    '備考欄',
    '応募ボタン送信',
    '応募ボタン送信日時',
    '応募ボタン送信方法'
  ];

  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  sheet.setFrozenRows(1);
  sheet.getRange('A:T').setWrap(true);
  sheet.getRange('K:K').setNumberFormat('yyyy/mm/dd hh:mm:ss');
  sheet.getRange('S:S').setNumberFormat('yyyy/mm/dd hh:mm:ss');

  sheet.autoResizeColumns(1, headers.length);
}

function setupErrorLogSheetForRecruit_(ss) {
  const sheet = recreateSheet_(ss, SHEET_ERROR_LOG);

  const headers = [
    '日時',
    '種別',
    '内容',
    'ユーザーID',
    'メモ'
  ];

  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  sheet.setFrozenRows(1);
  sheet.getRange('A:E').setWrap(true);
  sheet.getRange('A:A').setNumberFormat('yyyy/mm/dd hh:mm:ss');

  sheet.autoResizeColumns(1, headers.length);
}

function setupProcessLogSheetForRecruit_(ss) {
  const sheet = recreateSheet_(ss, SHEET_PROCESS_LOG);

  const headers = [
    '日時',
    '種別',
    '内容',
    'ユーザーID',
    'メモ'
  ];

  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  sheet.setFrozenRows(1);
  sheet.getRange('A:E').setWrap(true);
  sheet.getRange('A:A').setNumberFormat('yyyy/mm/dd hh:mm:ss');

  sheet.autoResizeColumns(1, headers.length);
}

function setupSpecSheetForRecruit_(ss) {
  const sheet = recreateSheet_(ss, '仕様書');

  const rows = [
    ['項目', '内容'],
    ['応募開始', '友だち追加時に自動で質問1を送信します。以後は「応募開始」の完全一致のみで開始します。'],
    ['連絡先', '友だち追加時に連絡先シートへユーザーIDを保存します。'],
    ['追客', '連絡先シートの追客ステータスで、未応募者への2日/5日/7日追客を管理します。'],
    ['質問1', '募集媒体をFlexボタンで選択します。表示は 質問1/8 です。'],
    ['質問1選択肢', 'カフェるん / Instagram広告 / X / その他'],
    ['通常ルート', 'カフェるん / Instagram広告 / X の場合は 質問2/7 から通常フローへ進みます。'],
    ['その他ルート', 'その他を選ぶと 質問2/8 で YouTube / TikTok / ポケパラ / 紹介 を選択します。'],
    ['その他ルート後', 'その他媒体回答後は 質問3/8 から通常フローへ進みます。'],
    ['管理者通知', 'その他ルートの場合も 募集媒体：紹介 のように選択結果だけ表示します。'],
    ['再応募', '問い合わせフローには入らず、通常応募として保存します。'],
    ['再応募カウント', '同一ユーザーIDの応募済み件数から応募回数を自動計算します。'],
    ['管理者通知', 'ADMIN_GROUP_ID に @All メンション付きで通知します。'],
    ['対応管理同期', '対応管理の 面接担当/名前/合否/面接/体入/本入 を応募管理へ同期します。'],
    ['対応履歴', '面接/体入/本入を編集した日時を対応履歴へ保存します。'],
    ['不合格処理', '合否を不合格にすると、面接/体入/本入を × にして履歴へ残します。'],
    ['処理ログ', 'FOLLOW / QUESTION_1_SENT / QUESTION_1_ANSWER / APPLICATION_SAVED / APPLICATION_SAVE_ERROR のみ保存します。'],
    ['エラーログ', 'エラーはエラーログへ保存します。']
  ];

  sheet
    .getRange(1, 1, rows.length, 2)
    .setValues(rows);

  sheet.setFrozenRows(1);
  sheet.getRange('A:B').setWrap(true);
  sheet.autoResizeColumns(1, 2);
}

function setupOperationMemoSheetForRecruit_(ss) {
  const sheet = recreateSheet_(ss, '運用メモ');

  const rows = [
    ['項目', 'メモ'],
    ['初期設定', 'setupRecruitSpreadsheet を実行してください。'],
    ['注意', '既存データがある場合でも、現在の安全版ではシート削除は行いません。'],
    ['注意2', 'ただし、ヘッダー行・説明行・入力規則・表示形式は更新されます。'],
    ['連絡先', '友だち追加済み / やりとりあり / 応募中 / 応募完了 を管理します。'],
    ['追客', '未応募者へ2日/5日/7日で自動追客します。返信・応募開始・応募完了時は追客ステータスを停止します。'],
    ['START_WORD', '応募開始。完全一致のみです。'],
    ['LINEトークン', '設定シートの LINE_CHANNEL_ACCESS_TOKEN に入力してください。'],
    ['管理者通知', '設定シートの ADMIN_GROUP_ID に通知先グループIDを入力してください。'],
    ['グループID確認', '求人管理 → LINEグループID確認ログをONにして、対象グループでメッセージを送ってください。'],
    ['Webhook反映', 'コード変更後は Webアプリを新しいバージョンで再デプロイしてください。'],
    ['質問1', '質問1は 1/8 表示です。その他以外は 2/7 に進みます。'],
    ['その他分岐', '質問1でその他を選んだ場合、質問2/8 で YouTube / TikTok / ポケパラ / 紹介 を選択します。'],
    ['対応履歴', '面接 / 体入 / 本入 を編集すると対応履歴へ自動保存します。'],
    ['不合格処理', '合否を不合格にすると、面接 / 体入 / 本入 を × にして履歴へ残します。'],
    ['処理ログ', '重要ログのみ保存します。レスポンス速度を優先しています。'],
    ['重要ログ', 'FOLLOW / QUESTION_1_SENT / QUESTION_1_ANSWER / APPLICATION_SAVED / APPLICATION_SAVE_ERROR'],
    ['本番運用', '処理ログを完全停止したい場合は ENABLE_PROCESS_LOG を false にしてください。']
  ];

  sheet
    .getRange(1, 1, rows.length, 2)
    .setValues(rows);

  sheet.setFrozenRows(1);
  sheet.getRange('A:B').setWrap(true);
  sheet.autoResizeColumns(1, 2);
}

function setupSupportStatusConditionalFormat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SUPPORT);

  if (!sheet) {
    SpreadsheetApp
      .getUi()
      .alert(`シートが見つかりません: ${SHEET_SUPPORT}`);
    return;
  }

  const lastColumn = Math.max(sheet.getLastColumn(), 14);
  const targetRange = sheet.getRange(3, 1, 998, lastColumn);

  const existingRules = sheet.getConditionalFormatRules();

  const filteredRules = existingRules.filter(rule => {
    const ranges = rule.getRanges();

    return !ranges.some(range => {
      return (
        range.getSheet().getName() === SHEET_SUPPORT &&
        range.getRow() === 3 &&
        range.getColumn() === 1
      );
    });
  });

  const completedRule = SpreadsheetApp
    .newConditionalFormatRule()
    .whenFormulaSatisfied('=$B3="対応完了"')
    .setBackground('#4A4A4A')
    .setFontColor('#FFFFFF')
    .setRanges([targetRange])
    .build();

  filteredRules.push(completedRule);
  sheet.setConditionalFormatRules(filteredRules);
}

/**
 * 既存シートがあれば削除せず、そのまま返す。
 * 新規シートだけ作成する。
 */
function recreateSheet_(ss, sheetName) {
  const existing = ss.getSheetByName(sheetName);

  if (existing) {
    return existing;
  }

  return ss.insertSheet(sheetName);
}

function applyDropdown_(sheet, rangeA1, values) {
  const rule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();

  sheet
    .getRange(rangeA1)
    .setDataValidation(rule);
}