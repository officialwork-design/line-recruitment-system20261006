/**
 * 18_UserManagement.gs
 * 応募中ユーザー状態管理
 *
 * 方針：
 * - 応募中の状態は ScriptProperties に保存
 * - ユーザー管理シートは確認・運用用の同期先
 * - syncUserStatesToSheet() は全削除せず、必要範囲だけ更新する
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
    fullBodyPhotoStatus: data.fullBodyPhotoStatus || '',
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
 * 応募完了状態に変更
 */
function markUserDone_(
  state,
  userId,
  displayName
) {
  upsertUserManagement_({
    userId,
    displayName: displayName || state.displayName || '',
    currentQuestionNo: '',
    status: STATUS_DONE,
    name: state.name || '',
    store: state.store || '',
    media: state.media || '',
    mediaOtherFlow: state.mediaOtherFlow || 'FALSE',
    totalSteps: state.totalSteps || '',
    photoCount: Number(state.photoCount || 0),
    facePhotoStatus: state.facePhotoStatus || '',
    fullBodyPhotoStatus: state.fullBodyPhotoStatus || '',
    extraAnswers: state.extraAnswers || '{}',
    applicationMessage: state.applicationMessage || '',
    age: state.age || '',
    remarks: state.remarks || '',
    updatedAt: new Date().toISOString()
  });
}

/**
 * ScriptProperties のユーザー状態をユーザー管理シートへ同期
 *
 * 軽量化：
 * - ヘッダーは毎回保証
 * - 必要行だけ setValues
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
      setRowValueByHeader_(headers, row, '全身写真状態', data.fullBodyPhotoStatus || '');
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
        2,
        1,
        rows.length,
        headers.length
      )
      .setValues(rows);
  }

  const lastRow =
    sheet.getLastRow();

  const clearStartRow =
    rows.length + 2;

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
 * ユーザー管理シートのヘッダー保証
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