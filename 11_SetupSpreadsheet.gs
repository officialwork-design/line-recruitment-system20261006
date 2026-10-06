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
 *
 * 方針（V2移行・2026/10追記）：
 * - 応募管理／対応管理／連絡先／質問設定／選択肢設定の5シートは、
 *   旧バージョンのヘッダー・既存データ（本番からコピーされた可能性が
 *   あるものを含む）を一切削除・上書き・移動しない「V2ヘッダー追記
 *   方式」に変更した（10_CommonUtils.gs の ensureV2HeaderBlock_ 等を
 *   参照）。シートが空（新規）の場合は従来通り1行目から作成される。
 * - 既にV2ヘッダーが存在する場合（setup再実行時、または既にV2として
 *   一度セットアップ済みの場合）は、それを検出して再利用し、
 *   ヘッダー・初期データを重複生成しない。
 * - 上記5シート以外（設定／対応履歴／ユーザー管理／エラーログ／
 *   処理ログ／仕様書／運用メモ）は、もともと「現在の状態を都度
 *   全体書き換えするシート」（設定・ユーザー管理）か、「常に末尾へ
 *   追記するだけで途中に別ヘッダーを挟む必要がないシート」
 *   （対応履歴・エラーログ・処理ログ）、または単なる参考資料
 *   （仕様書・運用メモ）であり、V2ヘッダー追記方式の対象外とした
 *   （詳細はREADME.mdの「V2スプレッドシート移行方式」を参照）。
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

/**
 * 質問設定シート
 *
 * 方針（V2移行・2026/10）：
 * - 質問設定シートは「過去ログ」ではなく「現在有効な質問定義」を
 *   保持する設定シートである。旧バージョンの行がそのまま残っていると
 *   getActiveQuestions_() がそれも「有効な質問」として誤って拾って
 *   しまうため、V2のヘッダー・初期データは必ず区別できる位置
 *   （既存データの最終行の次。シートが空なら1行目）に追記し、
 *   getActiveQuestions_() 側もV2ヘッダー以降だけを読むようにしている
 *   （17_QuestionData.gs参照）。
 * - 旧バージョンの行は削除・上書きしない。
 * - 再実行してもV2ヘッダー・初期データが重複生成されないよう、
 *   V2ヘッダーが既に存在する場合は初期データの再投入を行わない。
 */
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

  const defaultRows = [
    [1, 'store', '希望店舗', '希望店舗を選択してください。', '選択ボタン', true, '希望店舗', true],
    [2, 'work_days', '勤務日数', '勤務できる日数を選択してください。', '選択ボタン', true, '勤務日数', true],
    [3, 'age', '年齢', '現時点で何歳ですか？', '文字入力', true, '年齢', true],
    [4, 'remarks', '備考欄', '最後に、補足や質問があれば入力してください。\n特になければ「なし」と入力してください。', '文字入力', false, '備考欄', true]
  ];

  const block = ensureV2HeaderBlock_(sheet, headers, null);

  if (block.created) {
    sheet
      .getRange(block.dataStartRow, 1, defaultRows.length, headers.length)
      .setValues(defaultRows);
  }

  const dropdownEnd = block.dataStartRow + 997;

  applyDropdown_(sheet, `E${block.dataStartRow}:E${dropdownEnd}`, ['文字入力', '選択ボタン']);
  applyDropdown_(sheet, `F${block.dataStartRow}:F${dropdownEnd}`, [true, false]);
  applyDropdown_(sheet, `H${block.dataStartRow}:H${dropdownEnd}`, [true, false]);

  if (block.headerRow === 1) {
    sheet.setFrozenRows(1);
  }

  const formatRows = Math.max(sheet.getMaxRows() - block.headerRow + 1, 1);
  sheet.getRange(block.headerRow, 1, formatRows, headers.length).setWrap(true);
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * 選択肢設定シート
 *
 * 方針：質問設定シートと同じ理由で、V2ヘッダー・初期データを
 * 既存データの下へ追記する（旧データは削除・上書きしない）。
 */
function setupChoiceSheetForRecruit_(ss) {
  const sheet = recreateSheet_(ss, SHEET_CHOICES);

  const headers = [
    '項目キー',
    '選択肢名',
    '送信テキスト',
    '表示順',
    '有効'
  ];

  const defaultRows = [
    ['store', '大阪店', '大阪店', 1, true],
    ['store', '新宿店', '新宿店', 2, true],
    ['store', '秋葉原店', '秋葉原店', 3, true],
    ['store', '複数店(新宿店/秋葉原店)', '複数店(新宿店/秋葉原店)', 4, true],
    ['work_days', '週1〜2日', '週1〜2日', 1, true],
    ['work_days', '週2〜3日', '週2〜3日', 2, true],
    ['work_days', '週4〜5日', '週4〜5日', 3, true],
    ['work_days', '相談したい', '相談したい', 4, true]
  ];

  const block = ensureV2HeaderBlock_(sheet, headers, null);

  if (block.created) {
    sheet
      .getRange(block.dataStartRow, 1, defaultRows.length, headers.length)
      .setValues(defaultRows);
  }

  const dropdownEnd = block.dataStartRow + 997;

  applyDropdown_(sheet, `E${block.dataStartRow}:E${dropdownEnd}`, [true, false]);

  if (block.headerRow === 1) {
    sheet.setFrozenRows(1);
  }

  sheet.autoResizeColumns(1, headers.length);
}

/**
 * 応募管理シート
 *
 * 方針（V2移行・2026/10）：
 * - 旧バージョンのヘッダー・既存データ（本番からコピーされた可能性が
 *   あるものを含む）は削除・上書き・移動しない。
 * - V2ヘッダー（＋説明行）は、既存データの最終行の次に追記する。
 *   シートが空（新規）の場合は従来通り1行目に作成する。
 * - 既にV2ヘッダーが存在する場合（setup再実行時）はそれを再利用し、
 *   重複作成しない。
 * - 入力規則・表示形式はV2ヘッダー以降の領域にのみ適用し、
 *   旧データ領域の書式は変更しない。
 * - 応募No採番・hasCompletedApplication_ 等の「過去データも含めて
 *   判定する」処理は、この関数の変更とは独立して
 *   12_ApplicationSheet.gs 側で従来通り3行目から最終行までを対象とする
 *   （途中に挟まるヘッダー・説明行は、No列が数値にならない／
 *   ユーザーIDが一致しないことで自然に除外される）。
 */
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

  const block = ensureV2HeaderBlock_(sheet, headers, notes);

  if (block.headerRow === 1) {
    sheet.setFrozenRows(2);
  }

  const formatRows = Math.max(sheet.getMaxRows() - block.headerRow + 1, 1);
  sheet.getRange(block.headerRow, 1, formatRows, headers.length).setWrap(true);

  const dataRows = Math.max(sheet.getMaxRows() - block.dataStartRow + 1, 1);
  sheet.getRange(block.dataStartRow, 2, dataRows, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss'); // 受信日時
  sheet.getRange(block.dataStartRow, 27, dataRows, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss'); // 更新日

  const dStart = block.dataStartRow;
  const dEnd = dStart + 997;

  applyDropdown_(sheet, `N${dStart}:N${dEnd}`, ['受信済み', '未受信']);
  applyDropdown_(sheet, `O${dStart}:O${dEnd}`, ['受信済み', '未受信']);
  applyDropdown_(sheet, `S${dStart}:S${dEnd}`, ['', '未設定', '合格', '不合格', '保留']);
  applyDropdown_(sheet, `T${dStart}:V${dEnd}`, ['', '○', '×']);
  applyDropdown_(sheet, `W${dStart}:W${dEnd}`, [STATUS_DONE]);
  applyDropdown_(sheet, `X${dStart}:X${dEnd}`, ['済', '未']);
  applyDropdown_(sheet, `Y${dStart}:Y${dEnd}`, ['初回', '再応募']);

  sheet.autoResizeColumns(1, headers.length);
}

/**
 * 対応管理シート（V2移行方針は応募管理シートと同様）
 */
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

  const block = ensureV2HeaderBlock_(sheet, headers, notes);

  if (block.headerRow === 1) {
    sheet.setFrozenRows(2);
  }

  const formatRows = Math.max(sheet.getMaxRows() - block.headerRow + 1, 1);
  sheet.getRange(block.headerRow, 1, formatRows, headers.length).setWrap(true);

  const dataRows = Math.max(sheet.getMaxRows() - block.dataStartRow + 1, 1);
  sheet.getRange(block.dataStartRow, 14, dataRows, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss'); // 更新日

  const dStart = block.dataStartRow;
  const dEnd = dStart + 997;

  applyDropdown_(sheet, `B${dStart}:B${dEnd}`, [
    SUPPORT_STATUS_NOT_STARTED,
    SUPPORT_STATUS_IN_PROGRESS,
    SUPPORT_STATUS_DONE
  ]);

  applyDropdown_(sheet, `C${dStart}:C${dEnd}`, [
    'OSAKA MG',
    '飯塚',
    'ひろ',
    'りゅうじ',
    'りゅうき'
  ]);

  applyDropdown_(sheet, `G${dStart}:G${dEnd}`, [
    '',
    '未設定',
    '合格',
    '不合格',
    '保留'
  ]);

  applyDropdown_(sheet, `H${dStart}:J${dEnd}`, [
    '',
    '○',
    '×'
  ]);

  sheet.autoResizeColumns(1, headers.length);
}

/**
 * 連絡先シート（V2移行方針は応募管理シートと同様）
 *
 * 重要：ensureContactHeader_()（15_Contacts.gs）は、友だち追加・
 * メッセージ受信のたびに毎回呼ばれる。旧実装は固定で1〜2行目へ
 * ヘッダー・説明行を上書きしていたため、旧データが存在する場合は
 * LINEイベントのたびに旧ヘッダーが破壊されてしまっていた。
 * ensureV2HeaderBlock_() による冪等化で、既存のV2ヘッダーがあれば
 * 何もしない（旧データへは一切書き込まない）。
 */
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

  const block = ensureV2HeaderBlock_(sheet, headers, notes);

  if (block.headerRow === 1) {
    sheet.setFrozenRows(2);
  }

  const formatRows = Math.max(sheet.getMaxRows() - block.headerRow + 1, 1);
  sheet.getRange(block.headerRow, 1, formatRows, headers.length).setWrap(true);

  const dataRows = Math.max(sheet.getMaxRows() - block.dataStartRow + 1, 1);
  sheet.getRange(block.dataStartRow, 3, dataRows, 2).setNumberFormat('yyyy/mm/dd hh:mm:ss'); // 友だち追加日時・最終やりとり日時
  sheet.getRange(block.dataStartRow, 10, dataRows, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss'); // 更新日

  const dStart = block.dataStartRow;
  const dEnd = dStart + 997;

  applyDropdown_(sheet, `E${dStart}:E${dEnd}`, [
    '友だち追加済み',
    'やりとりあり',
    '応募中',
    '応募完了',
    'ブロック不明'
  ]);

  applyDropdown_(sheet, `F${dStart}:F${dEnd}`, [
    '未応募',
    '応募中',
    '応募完了'
  ]);

  applyDropdown_(sheet, `G${dStart}:G${dEnd}`, [
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

/**
 * 対応管理の「対応完了」行の色付け
 *
 * V2移行（2026/10）：対象範囲の開始行を固定の3行目ではなく、
 * V2ヘッダー以降のデータ開始行から動的に計算する。これにより、
 * 旧データ領域には新しい条件付き書式を適用しない。
 */
function setupSupportStatusConditionalFormat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SUPPORT);

  if (!sheet) {
    SpreadsheetApp
      .getUi()
      .alert(`シートが見つかりません: ${SHEET_SUPPORT}`);
    return;
  }

  const supportHeaders = [
    'No', 'ステータス', '面接担当', 'LINE表示名', '応募メッセージ',
    '名前', '合否', '面接', '体入', '本入', '対応メモ',
    '過去応募者', '応募回数', '更新日'
  ];

  const headerRow = resolveV2HeaderRow_(sheet, supportHeaders);
  const dataStartRow = headerRow > 0 ? headerRow + 2 : 3;

  const lastColumn = Math.max(sheet.getLastColumn(), 14);
  const targetRange = sheet.getRange(dataStartRow, 1, 998, lastColumn);

  const existingRules = sheet.getConditionalFormatRules();

  // 旧データ領域に既に設定されている可能性のある条件付き書式ルールは
  // 変更しない。今回のV2データ開始行（dataStartRow）に同一のルールが
  // 既にある場合のみ、それを差し替える（setup再実行時の重複防止）。
  const filteredRules = existingRules.filter(rule => {
    const ranges = rule.getRanges();

    return !ranges.some(range => {
      return (
        range.getSheet().getName() === SHEET_SUPPORT &&
        range.getRow() === dataStartRow &&
        range.getColumn() === 1
      );
    });
  });

  const completedRule = SpreadsheetApp
    .newConditionalFormatRule()
    .whenFormulaSatisfied(`=$B${dataStartRow}="対応完了"`)
    .setBackground('#4A4A4A')
    .setFontColor('#FFFFFF')
    .setRanges([targetRange])
    .build();

  filteredRules.push(completedRule);
  sheet.setConditionalFormatRules(filteredRules);
}

/**
 * V2ヘッダー位置の検証用ユーティリティ（診断専用・非破壊）
 *
 * 位置づけ：
 * - 今回のV2ヘッダー追記方式は、「既にV2ヘッダーが存在するか」を
 *   ヘッダー配列の完全一致で自動判定する設計のため（
 *   ensureV2HeaderBlock_ 参照）、通常は setupRecruitSpreadsheet() を
 *   再実行するだけで、既存のV2テストスプレッドシート（既に一度
 *   セットアップ済みのもの）に対しても安全に追従する。
 *   そのため「データを移動する」という意味での移行処理は不要である。
 * - この関数は、その自動判定が各シートで実際にどう解決されたか
 *   （どの行をV2ヘッダーとして認識したか、新規作成したのか
 *   既存を再利用したのか）を確認するための診断専用関数であり、
 *   シートの内容は一切変更しない（ensureV2HeaderBlock_
 *   は「既に存在する場合は何もしない」ため、重複実行しても副作用はないが、
 *   本関数はあくまで確認目的でユーザーが手動実行するためのものであり、
 *   setupRecruitSpreadsheet() からは自動的に呼び出さない）。
 * - 実行結果はUIアラートとログ（Logger.log）の両方に出力する。
 */
function migrateExistingSheetsToV2_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const targets = [
    {
      name: SHEET_APPLICATIONS,
      headers: [
        'No', '受信日時', 'ユーザーID', 'LINE表示名', '募集媒体',
        '募集媒体大分類', 'その他媒体', '希望店舗', '勤務日数', '年齢',
        '備考欄', '応募メッセージ', '写真受信数', '顔写真確認',
        '全体写真確認', '追加回答', '面接担当', '名前', '合否', '面接',
        '体入', '本入', 'ステータス', '対応管理反映', '過去応募者',
        '応募回数', '更新日'
      ]
    },
    {
      name: SHEET_SUPPORT,
      headers: [
        'No', 'ステータス', '面接担当', 'LINE表示名', '応募メッセージ',
        '名前', '合否', '面接', '体入', '本入', '対応メモ', '過去応募者',
        '応募回数', '更新日'
      ]
    },
    {
      name: SHEET_CONTACTS,
      headers: [
        'ユーザーID', 'LINE表示名', '友だち追加日時', '最終やりとり日時',
        '状態', '応募状況', '追客ステータス', '最新応募No', 'メモ', '更新日'
      ]
    },
    {
      name: SHEET_QUESTIONS,
      headers: [
        '順番', '項目キー', '項目名', '質問文', '回答形式', '必須',
        '保存列名', '有効'
      ]
    },
    {
      name: SHEET_CHOICES,
      headers: ['項目キー', '選択肢名', '送信テキスト', '表示順', '有効']
    }
  ];

  const lines = targets.map(target => {
    const sheet = ss.getSheetByName(target.name);

    if (!sheet) {
      return `${target.name}: シートが存在しません（先に初期セットアップを実行してください）`;
    }

    const headerRow = resolveV2HeaderRow_(sheet, target.headers);
    const lastRow = sheet.getLastRow();

    if (headerRow === 0) {
      return `${target.name}: V2ヘッダー未作成（setupRecruitSpreadsheet() の実行が必要）/ 最終行=${lastRow}`;
    }

    const legacyRows = headerRow - 1;

    return (
      `${target.name}: V2ヘッダー=${headerRow}行目 / ` +
      `旧データ領域=${legacyRows > 0 ? `1〜${legacyRows}行目（${legacyRows}行）` : 'なし'} / ` +
      `最終行=${lastRow}`
    );
  });

  const report = lines.join('\n');

  Logger.log(report);

  SpreadsheetApp
    .getUi()
    .alert(
      'V2ヘッダー位置の確認結果（シート内容は変更していません）',
      report,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
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