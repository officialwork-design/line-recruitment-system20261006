/**
 * 18_UserManagement.gs
 * 応募中ユーザー状態管理
 *
 * 方針：
 * - 応募中の状態は ScriptProperties に保存
 * - ユーザー管理シートは確認・運用用の同期先
 * - syncUserStatesToSheet() は全削除せず、必要範囲だけ更新する
 *
 * 方針（V2改修・2026/10）：
 * - userState:${userId} は「応募途中セッション専用」の一時データとする。
 * - 完了履歴の正（Source of Truth）は応募管理シート（hasCompletedApplication_）。
 * - 応募完了・保存成功後は markUserDone_() が対象ユーザーの userState を
 *   削除するのではなく、個人情報を含まない最小限のDONEマーカーへ
 *   置き換える（二重保存防止の冪等性ガードとして使用。詳細は
 *   markUserDone_() のコメントを参照）。
 */

/**
 * ユーザー状態取得
 */
function getUserState_(userId) {
  if (!userId) return null;

  const json = PropertiesService
    .getScriptProperties()
    .getProperty(`userState:${userId}`);

  if (!json) return null;

  try {
    return JSON.parse(json);
  } catch (error) {
    return null;
  }
}

/**
 * ユーザー状態登録/更新
 */
function upsertUserManagement_(data) {
  if (!data || !data.userId) return;

  const stateData = {
    userId: data.userId,
    displayName: data.displayName || '',
    currentQuestionNo: data.currentQuestionNo || '',
    status: data.status || '',
    name: data.name || '',
    store: data.store || '',
    media: data.media || '',
    mediaOtherFlow: data.mediaOtherFlow || 'FALSE',
    totalSteps: data.totalSteps || '',
    photoCount: data.photoCount || 0,
    facePhotoStatus: data.facePhotoStatus || '',
    extraAnswers: data.extraAnswers || '{}',
    updatedAt: new Date().toISOString(),
    applicationMessage: data.applicationMessage || '',
    age: data.age || '',
    remarks: data.remarks || ''
  };

  PropertiesService
    .getScriptProperties()
    .setProperty(
      `userState:${data.userId}`,
      JSON.stringify(stateData)
    );
}

/**
 * 応募完了後のuserState更新（DONEマーカーへの置き換え）
 *
 * 方針（V2改修・2026/10）：
 * - 完了履歴の恒久的なSource of Truthは引き続き応募管理シート
 *   （hasCompletedApplication_ / countCompletedApplicationsByUserId_）。
 *   DONEマーカーはそれを置き換えるものではない。
 * - userState:${userId} はもう「応募途中セッション専用」の削除対象ではなく、
 *   保存成功後に以下の最小限のDONEマーカーへ置き換える。
 *     { status: STATUS_DONE, applicationNo, completedAt }
 *   氏名・年齢・電話番号・店舗回答・質問回答・備考・写真状態など、
 *   個人情報は一切含めない。
 * - DONEマーカーの目的は、Webhook再送等による同一完了処理の
 *   二重保存を防ぐための短期的な冪等性ガードのみ。
 *   saveApplicationFromState_ がLock取得直後に最新userStateを
 *   再取得し、status === STATUS_DONE であれば既に保存済みとして
 *   処理をスキップする（12_ApplicationSheet.gs参照）。
 * - 呼び出しは必ず「応募管理シートへの保存成功後」に限定すること。
 *   保存に失敗した場合はこの関数を呼ばない（＝応募途中state を残す）。
 * - deleteAllProperties() は絶対に使用しない。
 */
function markUserDone_(userId, applicationNo) {
  if (!userId) return;

  const doneMarker = {
    status: STATUS_DONE,
    applicationNo: applicationNo || '',
    completedAt: new Date().toISOString()
  };

  PropertiesService
    .getScriptProperties()
    .setProperty(
      `userState:${userId}`,
      JSON.stringify(doneMarker)
    );
}

/**
 * ScriptProperties のユーザー状態をユーザー管理シートへ同期
 *
 * 軽量化：
 * - ヘッダー・説明行は毎回保証（ensureUserManagementHeader_）
 * - データは3行目以降に必要行だけ setValues
 * - 余った古い行だけ clearContent
 */
function syncUserStatesToSheet() {
  const props = PropertiesService
    .getScriptProperties()
    .getProperties();

  const states = [];

  Object.keys(props).forEach(key => {
    if (key.indexOf('userState:') === 0) {
      try {
        const state =
          JSON.parse(props[key]);

        if (state && state.userId) {
          states.push(state);
        }
      } catch (error) {
        // 壊れたJSONは無視
      }
    }
  });

  states.sort((a, b) => {
    const timeA =
      new Date(a.updatedAt || 0).getTime();

    const timeB =
      new Date(b.updatedAt || 0).getTime();

    return timeB - timeA;
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    getOrCreateSheet_(
      ss,
      SHEET_USER_MANAGEMENT
    );

  ensureUserManagementHeader_();

  const headers =
    sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(h => String(h || '').trim());

  const rows =
    states.map(data => {
      const row =
        new Array(headers.length).fill('');

      setRowValueByHeader_(headers, row, 'ユーザーID', data.userId);
      setRowValueByHeader_(headers, row, 'LINE表示名', data.displayName || '');
      setRowValueByHeader_(headers, row, '現在の質問番号', data.currentQuestionNo || '');
      setRowValueByHeader_(headers, row, 'ステータス', data.status || '');
      setRowValueByHeader_(headers, row, 'お名前', data.name || '');
      setRowValueByHeader_(headers, row, '希望店舗', data.store || '');
      setRowValueByHeader_(headers, row, '写真受信数', data.photoCount || 0);
      setRowValueByHeader_(headers, row, '顔写真状態', data.facePhotoStatus || '');
      setRowValueByHeader_(headers, row, '追加回答JSON', data.extraAnswers || '{}');
      setRowValueByHeader_(headers, row, '更新日', data.updatedAt || '');
      setRowValueByHeader_(headers, row, '応募メッセージ', data.applicationMessage || '');
      setRowValueByHeader_(headers, row, '年齢', data.age || '');
      setRowValueByHeader_(headers, row, '募集媒体', data.media || '');
      setRowValueByHeader_(headers, row, 'その他媒体フロー', data.mediaOtherFlow || '');
      setRowValueByHeader_(headers, row, '総質問数', data.totalSteps || '');
      setRowValueByHeader_(headers, row, '備考欄', data.remarks || '');

      return row;
    });

  if (rows.length > 0) {
    sheet
      .getRange(
        3,
        1,
        rows.length,
        headers.length
      )
      .setValues(rows);
  }

  const lastRow =
    sheet.getLastRow();

  const clearStartRow =
    rows.length + 3;

  if (lastRow >= clearStartRow) {
    sheet
      .getRange(
        clearStartRow,
        1,
        lastRow - clearStartRow + 1,
        headers.length
      )
      .clearContent();
  }

  Logger.log(
    `ユーザー状態を同期しました。件数：${rows.length}`
  );
}

/**
 * ユーザー管理シートの説明行（2行目）配列
 * （11_SetupSpreadsheet.gs の setupUserManagementSheetForRecruit_ と
 * 必ず一致させること）
 */
function getUserManagementDescriptionRow_() {
  return [
    'LINEから取得',
    'LINEから取得',
    'システム内部管理',
    'システム内部管理',
    '応募回答から取得',
    '応募回答から取得',
    '自動',
    'システム内部管理（受信済み/未受信）',
    'JSON',
    '自動',
    '応募回答から取得',
    '応募回答から取得',
    '応募回答から取得',
    'システム内部管理（TRUE/FALSE）',
    'システム内部管理',
    '応募回答から取得',
    'システム内部管理',
    '日時',
    'システム内部管理'
  ];
}

/**
 * ユーザー管理シートのヘッダー保証
 *
 * 方針（2026/10改訂・説明行統一）：1行目＝最新ヘッダー／2行目＝最新説明／
 * 3行目以降＝ユーザー状態（ライブミラー）。2行目の説明行は
 * ensureDescriptionRow_ により、syncUserStatesToSheet() の全行再生成
 * を挟んでも維持されるよう、既存データを保護しながら安全に
 * 挿入・更新する。
 */
function ensureUserManagementHeader_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    getOrCreateSheet_(
      ss,
      SHEET_USER_MANAGEMENT
    );

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

  ensureDescriptionRow_(sheet, getUserManagementDescriptionRow_());

  sheet.setFrozenRows(2);
  sheet.getRange('A:S').setWrap(true);
  // 日付書式はデータ領域（3行目以降）にのみ適用する。
  sheet.getRange('J3:J').setNumberFormat('yyyy/mm/dd hh:mm:ss'); // 更新日
  sheet.getRange('R3:R').setNumberFormat('yyyy/mm/dd hh:mm:ss'); // 応募ボタン送信日時

  sheet.autoResizeColumns(1, headers.length);
}