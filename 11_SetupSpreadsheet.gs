/**
 * 11_SetupSpreadsheet.gs
 * 求人管理スプレッドシート初期セットアップ
 *
 * 方針：
 * - 新規ファイルは作らない
 * - 既存シートは削除しない
 * - 既存データ行は削除しない
 * - ヘッダー、メモ、入力規則、表示形式は更新する
 * - 本番誤操作防止のため、実行前に確認ダイアログを出す
 *
 * 方針（V2説明行統一・2026/10改訂）：
 * - 複数ヘッダーをシート内に追記して「旧データ領域」と「V2データ領域」を
 *   物理的に分離する方式（旧設計）は廃止した。
 * - 最終仕様：通常の表形式シートは常に「1行目＝最新のV2ヘッダー」
 *   「2行目＝各列の最新説明」「3行目以降＝実データ領域」とする
 *   （設定・仕様書・運用メモ等の特殊構造シートは対象外、下記）。
 *   setupを実行するたびに1行目のヘッダーは最新定義で無条件に上書きし、
 *   2行目の説明行は共通ヘルパー ensureDescriptionRow_()（10_CommonUtils.gs）
 *   で安全に保証する。3行目以降の既存データ行は一切変更しない。
 * - ensureDescriptionRow_() は、2行目が既に実データである場合のみ
 *   2行目の手前に新しい行を挿入してからデータを押し下げ、既存データを
 *   保護する。2行目が既に説明行（または空白）の場合は、行を増やさず
 *   その場で説明文を上書きするだけなので、setup再実行で説明行が
 *   増殖することはない。
 * - 既存シート：シート削除・データ行削除は行わない。1〜2行目のみ
 *   更新し、入力規則・表示形式を最新の列構成に合わせて再適用する。
 * - 新規シート：シートを作成し、1行目にヘッダー、2行目に説明、
 *   必要であれば3行目以降に初期値・入力規則・表示形式を書き込む。
 * - 列を削除した場合、旧データ行はその列の意味とズレて表示されることが
 *   あるが、これは許容する（旧データ行自体を削除することはしない）。
 * - 応募管理・対応管理の2シートは、本改訂以前から「1行目＝ヘッダー、
 *   2行目＝説明行、3行目以降＝実データ」という構成を採用しており、
 *   この構成はLockServiceによる応募No採番、重複保存防止
 *   （hasCompletedApplication_ 等）、対応管理同期など複数ファイルに
 *   またがる冪等性ロジックの前提になっている。今回の改訂後も構成自体は
 *   変わらないため、この2シートはそのまま3行目開始を維持する（詳細は
 *   README.mdの「Spreadsheetの構成方針」を参照）。
 * - 連絡先シートは、前回の改訂で一度「1行目＝ヘッダー／2行目以降＝
 *   データ」に簡素化したが、今回の改訂で再び「1行目＝ヘッダー／
 *   2行目＝説明行／3行目以降＝データ」へ統一した。
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

/**
 * 設定シート
 *
 * 方針（2026/10改訂）：
 * - setupを再実行しても、既に値が入力済みのキー（特に
 *   LINE_CHANNEL_ACCESS_TOKEN・ADMIN_GROUP_ID）の「値」列は
 *   絶対に上書き・リセットしない。
 * - 既存の値が空欄のキーのみ、デフォルト値／プレースホルダー文言で
 *   補完する。
 * - ヘッダー行・説明列・備考列は常に最新の定義で更新する。
 * - デフォルト定義にない既存キー（手動追加分など）は削除せず、
 *   末尾にそのまま残す。
 * - 秘密情報（実際のトークン値等）はこの関数のコードにもREADMEにも
 *   一切書き込まない。
 */
function setupConfigSheetForRecruit_(ss) {
  const sheet = recreateSheet_(ss, SHEET_CONFIG);

  const headerRow = ['キー', '値', '説明', '備考'];

  const defaultRows = [
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

  const lastRow = sheet.getLastRow();
  const existingMap = {};

  if (lastRow >= 2) {
    sheet
      .getRange(2, 1, lastRow - 1, 2)
      .getValues()
      .forEach(r => {
        const key = String(r[0] || '').trim();
        if (!key) return;
        existingMap[key] = r[1];
      });
  }

  const defaultKeys = defaultRows.map(r => r[0]);

  const mergedRows = defaultRows.map(defRow => {
    const key = defRow[0];
    const hasExisting = Object.prototype.hasOwnProperty.call(existingMap, key);
    const existingValue = hasExisting ? existingMap[key] : '';
    const keepExisting = hasExisting && notBlank_(existingValue);

    return [
      key,
      keepExisting ? existingValue : defRow[1],
      defRow[2],
      defRow[3]
    ];
  });

  // デフォルト定義にない既存キー（手動追加分）は削除せず末尾に残す
  const extraRows = Object.keys(existingMap)
    .filter(key => defaultKeys.indexOf(key) === -1)
    .map(key => [key, existingMap[key], '', '']);

  const allRows =
    [headerRow].concat(mergedRows, extraRows);

  sheet
    .getRange(1, 1, allRows.length, headerRow.length)
    .setValues(allRows);

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headerRow.length);
}

/**
 * 質問設定シート
 *
 * 方針（2026/10改訂・説明行統一）：1行目＝最新ヘッダー／2行目＝最新説明／
 * 3行目以降＝質問データ。初期データはシートが元々空だった場合のみ
 * 3行目以降へ投入する。2行目の説明行は ensureDescriptionRow_ により、
 * 既存の実データ行を保護しながら安全に挿入・更新する。
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

  // （17_QuestionData.gs の getQuestionHeaderDefinition_() と必ず一致させること）
  const descriptionRow = [
    '質問表示順',
    'システム内部キー',
    '管理用項目名',
    'LINEで表示する質問',
    '選択ボタン / 文字入力',
    'TRUE/FALSE',
    '応募管理で保存する列',
    'TRUE/FALSE'
  ];

  const defaultRows = [
    [1, 'store', '希望店舗', '希望店舗を選択してください。', '選択ボタン', true, '希望店舗', true],
    [2, 'work_days', '勤務日数', '勤務できる日数を選択してください。', '選択ボタン', true, '勤務日数', true],
    [3, 'age', '年齢', '現時点で何歳ですか？', '文字入力', true, '年齢', true],
    [4, 'remarks', '備考欄', '最後に、補足や質問があれば入力してください。\n特になければ「なし」と入力してください。', '文字入力', false, '備考欄', true]
  ];

  const hadNoDataRows = sheet.getLastRow() < 2;

  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  const insertedRow =
    ensureDescriptionRow_(sheet, descriptionRow);

  if (hadNoDataRows && !insertedRow) {
    sheet
      .getRange(3, 1, defaultRows.length, headers.length)
      .setValues(defaultRows);
  }

  const dataStartRow = 3;
  const dropdownEnd = dataStartRow + 997;

  applyDropdown_(sheet, `E${dataStartRow}:E${dropdownEnd}`, ['文字入力', '選択ボタン']);
  applyDropdown_(sheet, `F${dataStartRow}:F${dropdownEnd}`, [true, false]);
  applyDropdown_(sheet, `H${dataStartRow}:H${dropdownEnd}`, [true, false]);

  sheet.setFrozenRows(2);

  const formatRows = Math.max(sheet.getMaxRows() - 1 + 1, 1);
  sheet.getRange(1, 1, formatRows, headers.length).setWrap(true);
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * 選択肢設定シート
 *
 * 方針（2026/10改訂・説明行統一）：質問設定シートと同じ
 * （1行目＝最新ヘッダー、2行目＝最新説明、3行目以降＝データ、
 * 既存データは削除・上書きしない。初期データはシートが元々空だった
 * 場合のみ3行目以降へ投入する）。
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

  // （17_QuestionData.gs の getChoiceHeaderDefinition_() と必ず一致させること）
  const descriptionRow = [
    '質問設定と紐づくキー',
    'ユーザーへ表示する名称',
    'LINE送信時の値',
    '選択肢の並び順',
    'TRUE/FALSE'
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

  const hadNoDataRows = sheet.getLastRow() < 2;

  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  const insertedRow =
    ensureDescriptionRow_(sheet, descriptionRow);

  if (hadNoDataRows && !insertedRow) {
    sheet
      .getRange(3, 1, defaultRows.length, headers.length)
      .setValues(defaultRows);
  }

  const dataStartRow = 3;
  const dropdownEnd = dataStartRow + 997;

  applyDropdown_(sheet, `E${dataStartRow}:E${dropdownEnd}`, [true, false]);

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * 応募管理シート
 *
 * 方針（2026/10改訂）：
 * - このシートは、本改訂以前からの「1行目＝ヘッダー／2行目＝説明行／
 *   3行目以降＝実データ」という構成を維持する（他の一般シートも今回の
 *   改訂で同じ構成に統一されたが、こちらは元々この構成だった）。この構成は、応募No採番
 *   （generateApplicationNo_）・重複保存防止（hasCompletedApplication_／
 *   countCompletedApplicationsByUserId_）・LockServiceによる排他制御など、
 *   12_ApplicationSheet.gs・13_SupportSheetSync.gs 側の複数の処理が
 *   「3行目から最終行まで」を前提に実装されているため、変更すると
 *   それらの冪等性・重複防止ロジックが壊れるおそれがある。
 * - setupを再実行するたびに1行目・2行目（ヘッダー・説明行）のみを
 *   最新定義で上書きし、3行目以降の既存データ行は一切変更しない。
 * - 1枚写真フローへの簡略化に伴い、旧2枚写真フロー専用列
 *   「全体写真確認」は削除した。この列を物理的に持っていた旧データ行は、
 *   削除後の列構成とズレて表示される場合があるが、これは許容する
 *   （旧データ行自体は削除しない）。
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

  const headerRow = 1;
  const dataStartRow = 3;

  sheet
    .getRange(headerRow, 1, 1, headers.length)
    .setValues([headers]);

  // 2行目（説明行）にデータ用の入力規則が残っていないか必ず解除してから
  // 説明文を書き込む（旧バージョンで2行目がデータ行だった頃の入力規則が
  // 残っていると、書き込み時に「入力規則に違反しています」で失敗するため）。
  sheet
    .getRange(headerRow + 1, 1, 1, notes.length)
    .clearDataValidations();

  sheet
    .getRange(headerRow + 1, 1, 1, notes.length)
    .setValues([notes]);

  sheet.setFrozenRows(2);

  const formatRows = Math.max(sheet.getMaxRows() - headerRow + 1, 1);
  sheet.getRange(headerRow, 1, formatRows, headers.length).setWrap(true);

  const dataRows = Math.max(sheet.getMaxRows() - dataStartRow + 1, 1);
  sheet.getRange(dataStartRow, 2, dataRows, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss'); // 受信日時
  sheet.getRange(dataStartRow, 26, dataRows, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss'); // 更新日

  const dStart = dataStartRow;
  const dEnd = dStart + 997;

  applyDropdown_(sheet, `N${dStart}:N${dEnd}`, ['受信済み', '未受信']); // 顔写真確認
  applyDropdown_(sheet, `R${dStart}:R${dEnd}`, ['', '未設定', '合格', '不合格', '保留']); // 合否
  applyDropdown_(sheet, `S${dStart}:U${dEnd}`, ['', '○', '×']); // 面接/体入/本入
  applyDropdown_(sheet, `V${dStart}:V${dEnd}`, [STATUS_DONE]); // ステータス
  applyDropdown_(sheet, `W${dStart}:W${dEnd}`, ['済', '未']); // 対応管理反映
  applyDropdown_(sheet, `X${dStart}:X${dEnd}`, ['初回', '再応募']); // 過去応募者

  sheet.autoResizeColumns(1, headers.length);
}

/**
 * 対応管理シート
 *
 * 方針（2026/10改訂）：応募管理シートと同じ理由・同じ例外により、
 * 「1行目＝ヘッダー／2行目＝説明行／3行目以降＝実データ」の構成を
 * 維持する（13_SupportSheetSync.gs・14_SupportHistory.gs 側の複数の
 * 処理が3行目開始を前提にしているため）。setup再実行時は1〜2行目のみ
 * 最新定義で上書きし、3行目以降の既存データは変更しない。
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

  const headerRow = 1;
  const dataStartRow = 3;

  sheet
    .getRange(headerRow, 1, 1, headers.length)
    .setValues([headers]);

  // 2行目（説明行）にデータ用の入力規則が残っていないか必ず解除してから
  // 説明文を書き込む（旧バージョンで2行目がデータ行だった頃の入力規則が
  // 残っていると、書き込み時に「入力規則に違反しています」で失敗するため）。
  sheet
    .getRange(headerRow + 1, 1, 1, notes.length)
    .clearDataValidations();

  sheet
    .getRange(headerRow + 1, 1, 1, notes.length)
    .setValues([notes]);

  sheet.setFrozenRows(2);

  const formatRows = Math.max(sheet.getMaxRows() - headerRow + 1, 1);
  sheet.getRange(headerRow, 1, formatRows, headers.length).setWrap(true);

  const dataRows = Math.max(sheet.getMaxRows() - dataStartRow + 1, 1);
  sheet.getRange(dataStartRow, 14, dataRows, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss'); // 更新日

  const dStart = dataStartRow;
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
 * 連絡先シート
 *
 * 方針（2026/10改訂・説明行統一）：通常の表形式シートと同じ
 * 「1行目＝最新ヘッダー／2行目＝最新説明／3行目以降＝データ」に統一する。
 *
 * 重要：この関数が呼び出す ensureContactHeader_()（15_Contacts.gs）は、
 * 友だち追加・メッセージ受信のたびに毎回呼ばれる。そのたびに1行目を
 * 無条件で上書きするのではなく、1行目が既に最新ヘッダーと一致する
 * 場合は書き込みをスキップする（ensureContactHeader_ 側で判定）。
 * 2行目の説明行は ensureDescriptionRow_ により、既存の実データ行を
 * 上書きせず安全に挿入・更新する（3行目以降の既存データ行は一切
 * 変更しない）。
 */
function setupContactSheetForRecruit_(ss) {
  const sheet = recreateSheet_(ss, SHEET_CONTACTS);

  // （15_Contacts.gs の getContactHeaderDefinition_() と必ず一致させること）
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

  // （15_Contacts.gs の getContactDescriptionRow_() と必ず一致させること）
  const descriptionRow = [
    'LINEから取得',
    'LINEから取得',
    '自動（友だち追加時）',
    '自動（メッセージ受信時）',
    'システム内部管理',
    'システム内部管理',
    'システム内部管理',
    '応募管理と連携',
    '手入力',
    '自動'
  ];

  const headerRow = 1;
  const dataStartRow = 3;

  sheet
    .getRange(headerRow, 1, 1, headers.length)
    .setValues([headers]);

  ensureDescriptionRow_(sheet, descriptionRow);

  sheet.setFrozenRows(2);

  const formatRows = Math.max(sheet.getMaxRows() - headerRow + 1, 1);
  sheet.getRange(headerRow, 1, formatRows, headers.length).setWrap(true);

  const dataRows = Math.max(sheet.getMaxRows() - dataStartRow + 1, 1);
  sheet.getRange(dataStartRow, 3, dataRows, 2).setNumberFormat('yyyy/mm/dd hh:mm:ss'); // 友だち追加日時・最終やりとり日時
  sheet.getRange(dataStartRow, 10, dataRows, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss'); // 更新日

  const dStart = dataStartRow;
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

/**
 * 対応履歴シート
 *
 * 方針（2026/10改訂・説明行統一）：1行目＝最新ヘッダー／2行目＝最新説明／
 * 3行目以降＝履歴データ。1行目の無条件上書きは引き続き許容するが、
 * 2行目の説明行は ensureDescriptionRow_ により既存の履歴データを
 * 保護しながら安全に挿入・更新する（14_SupportHistory.gs の
 * ensureSupportHistoryHeader_ と必ず一致させること）。
 */
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

  ensureDescriptionRow_(sheet, getSupportHistoryDescriptionRow_());

  sheet.setFrozenRows(2);
  sheet.getRange('A:H').setWrap(true);
  // 日付書式はデータ領域（3行目以降）にのみ適用し、説明行（2行目）には
  // データ用の書式を残さない。
  sheet.getRange('A3:A').setNumberFormat('yyyy/mm/dd hh:mm:ss');
  sheet.getRange('H3:H').setNumberFormat('yyyy/mm/dd hh:mm:ss');

  try {
    sheet.hideSheet();
  } catch (error) {}
}

/**
 * ユーザー管理シート
 *
 * 方針：このシートは現在の userState（Script Properties）の
 * ライブミラーであり、syncUserStatesToSheet() が同期のたびに
 * 行全体を再生成する使い捨てシートである（過去ログではない）ため、
 * 毎回の全体書き換えは引き続き許容する。ここではヘッダー定義のみを
 * 最新仕様（1枚写真フローに伴い「全身写真状態」列を削除）に更新する。
 */
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

  // （18_UserManagement.gs の getUserManagementDescriptionRow_() と必ず一致させること）
  ensureDescriptionRow_(sheet, getUserManagementDescriptionRow_());

  sheet.setFrozenRows(2);
  sheet.getRange('A:S').setWrap(true);
  // 日付書式はデータ領域（3行目以降）にのみ適用する。
  sheet.getRange('J3:J').setNumberFormat('yyyy/mm/dd hh:mm:ss'); // 更新日
  sheet.getRange('R3:R').setNumberFormat('yyyy/mm/dd hh:mm:ss'); // 応募ボタン送信日時

  sheet.autoResizeColumns(1, headers.length);
}

/**
 * エラーログシート
 *
 * 方針（2026/10改訂・説明行統一）：1行目＝最新ヘッダー／2行目＝最新説明／
 * 3行目以降＝ログ（追記専用）。2行目は ensureDescriptionRow_
 * （説明は 10_CommonUtils.gs の getLogDescriptionRow_() と必ず
 * 一致させること）により、既存ログを保護しながら安全に挿入・更新する。
 */
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

  ensureDescriptionRow_(sheet, getLogDescriptionRow_());

  sheet.setFrozenRows(2);
  sheet.getRange('A:E').setWrap(true);
  // 日付書式はデータ領域（3行目以降）にのみ適用する。
  sheet.getRange('A3:A').setNumberFormat('yyyy/mm/dd hh:mm:ss');

  sheet.autoResizeColumns(1, headers.length);
}

/**
 * 処理ログシート
 *
 * 方針（2026/10改訂・説明行統一）：エラーログシートと同じ
 * （1行目＝最新ヘッダー／2行目＝最新説明／3行目以降＝ログ）。
 */
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

  ensureDescriptionRow_(sheet, getLogDescriptionRow_());

  sheet.setFrozenRows(2);
  sheet.getRange('A:E').setWrap(true);
  // 日付書式はデータ領域（3行目以降）にのみ適用する。
  sheet.getRange('A3:A').setNumberFormat('yyyy/mm/dd hh:mm:ss');

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
 * 方針（2026/10改訂）：対応管理シートは1行目＝ヘッダー／2行目＝説明行／
 * 3行目以降＝実データの固定構成（応募管理と同じ理由による例外）なので、
 * データ開始行は固定で3行目とする。
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

  const dataStartRow = 3;

  const lastColumn = Math.max(sheet.getLastColumn(), 14);
  const targetRange = sheet.getRange(dataStartRow, 1, 998, lastColumn);

  const existingRules = sheet.getConditionalFormatRules();

  // データ開始行（3行目）を対象とする既存の同種ルールがあれば差し替え、
  // それ以外の既存ルールはそのまま残す（setup再実行時の重複防止）。
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