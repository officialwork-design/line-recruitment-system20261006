/**
 * 12_ApplicationSheet.gs
 * 応募管理シート保存・応募No管理
 */

/**
 * 応募状態を応募管理へ保存
 *
 * 方針（V2改修・2026/10）：
 * - 応募No採番〜応募管理シート書き込み〜対応管理追加までを
 *   LockService.getScriptLock() で同一ロック内に収め、
 *   同時完了時の応募No重複・行競合を防ぐ。
 * - ロック中は外部API通信（UrlFetchApp）を一切行わない。
 *   管理者LINE通知はこの関数から呼ばず、ロック解放後に
 *   呼び出し元（saveCompletedApplication_）側で行う。
 * - ロック取得に失敗した場合は無言で続行せず、例外を投げる。
 */
function saveApplicationFromState_(
  state,
  userId,
  displayName
) {
  const lock =
    LockService.getScriptLock();

  const gotLock =
    lock.tryLock(10000);

  if (!gotLock) {
    throw new Error(
      '応募データ保存の排他ロックを取得できませんでした（LOCK_TIMEOUT）。'
    );
  }

  try {
    const sheet =
      SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName(SHEET_APPLICATIONS);

    if (!sheet) {
      throw new Error(`Sheet not found: ${SHEET_APPLICATIONS}`);
    }

    const headers =
      getApplicationHeaders_();

    const extraAnswers =
      parseJsonSafe_(state.extraAnswers || '{}');

    const applicationNo =
      generateApplicationNo_();

    const pastApplicationCount =
      countCompletedApplicationsByUserId_(userId);

    const applicationCount =
      pastApplicationCount + 1;

    const applicationHistoryStatus =
      pastApplicationCount > 0
        ? '再応募'
        : '初回';

    const row =
      buildApplicationRow_({
        headers,
        state,
        userId,
        displayName,
        extraAnswers,
        applicationNo,
        applicationHistoryStatus,
        applicationCount
      });

    const targetRow =
      Math.max(sheet.getLastRow() + 1, 3);

    sheet
      .getRange(targetRow, 1, 1, row.length)
      .setValues([row]);

    processLog_(
      'APPLICATION_SAVED',
      '応募管理へ保存しました。',
      userId,
      `応募No.${applicationNo}`
    );

    addSupportRowSafely_(
      applicationNo,
      state,
      userId,
      displayName,
      {
        applicationHistoryStatus,
        applicationCount
      }
    );

    return applicationNo;

  } finally {
    lock.releaseLock();
  }
}

/**
 * 応募管理保存用の1行データを作成
 */
function buildApplicationRow_(params) {
  const headers =
    params.headers;

  const state =
    params.state || {};

  const extraAnswers =
    params.extraAnswers || {};

  const row =
    new Array(headers.length).fill('');

  const media =
    state.media ||
    extraAnswers['募集媒体'] ||
    '';

  const mediaCategory =
    extraAnswers['募集媒体大分類'] || '';

  const mediaOther =
    extraAnswers['その他媒体'] || '';

  setRowValueByHeader_(headers, row, 'No', params.applicationNo);
  setRowValueByHeader_(headers, row, '受信日時', new Date());
  setRowValueByHeader_(headers, row, 'ユーザーID', params.userId);
  setRowValueByHeader_(headers, row, 'LINE表示名', params.displayName || state.displayName || '');
  setRowValueByHeader_(headers, row, '募集媒体', media);
  setRowValueByHeader_(headers, row, '募集媒体大分類', mediaCategory);
  setRowValueByHeader_(headers, row, 'その他媒体', mediaOther);
  setRowValueByHeader_(headers, row, '希望店舗', state.store || extraAnswers['希望店舗'] || '');
  setRowValueByHeader_(headers, row, '勤務日数', extraAnswers['勤務日数'] || '');
  setRowValueByHeader_(headers, row, '年齢', state.age || extraAnswers['年齢'] || '');
  setRowValueByHeader_(headers, row, '備考欄', state.remarks || extraAnswers['備考欄'] || '');
  setRowValueByHeader_(headers, row, '応募メッセージ', state.applicationMessage || '');
  setRowValueByHeader_(headers, row, '写真受信数', Number(state.photoCount || 0));
  setRowValueByHeader_(headers, row, '顔写真確認', state.facePhotoStatus || '未受信');
  setRowValueByHeader_(headers, row, '全体写真確認', state.fullBodyPhotoStatus || '未受信');
  setRowValueByHeader_(headers, row, '追加回答', JSON.stringify(extraAnswers));
  setRowValueByHeader_(headers, row, '面接担当', '');
  setRowValueByHeader_(headers, row, '名前', state.name || extraAnswers['名前'] || extraAnswers['お名前'] || '');
  setRowValueByHeader_(headers, row, '合否', '');
  setRowValueByHeader_(headers, row, '面接', '');
  setRowValueByHeader_(headers, row, '体入', '');
  setRowValueByHeader_(headers, row, '本入', '');
  setRowValueByHeader_(headers, row, 'ステータス', STATUS_DONE);
  setRowValueByHeader_(headers, row, '対応管理反映', '済');
  setRowValueByHeader_(headers, row, '過去応募者', params.applicationHistoryStatus);
  setRowValueByHeader_(headers, row, '応募回数', params.applicationCount);
  setRowValueByHeader_(headers, row, '更新日', new Date());

  Object.keys(extraAnswers).forEach(key => {
    setRowValueByHeader_(
      headers,
      row,
      key,
      extraAnswers[key]
    );
  });

  return row;
}

/**
 * 対応管理へ応募行を追加
 * 失敗しても応募保存自体は止めない
 */
function addSupportRowSafely_(
  applicationNo,
  state,
  userId,
  displayName,
  supportMeta
) {
  try {
    addSupportRowFromApplication_(
      applicationNo,
      state,
      userId,
      displayName,
      supportMeta
    );
  } catch (error) {
    errorLog_(
      'SUPPORT_ROW_ADD_ERROR',
      error.stack || error.message,
      userId,
      `応募No.${applicationNo}`
    );
  }
}

/**
 * 管理者通知
 * 失敗しても応募保存自体は止めない
 */
function notifyApplicationToAdminSafely_(
  applicationNo,
  state,
  userId,
  displayName
) {
  try {
    notifyApplicationToAdmin_(
      applicationNo,
      state,
      userId,
      displayName
    );

    processLog_(
      'ADMIN_NOTIFY_SENT',
      '管理者通知を送信しました。',
      userId,
      `応募No.${applicationNo}`
    );

  } catch (error) {
    errorLog_(
      'APPLICATION_ADMIN_NOTIFY_ERROR',
      error.stack || error.message,
      userId,
      `応募No.${applicationNo}`
    );
  }
}

/**
 * 応募管理ヘッダー取得
 */
function getApplicationHeaders_() {
  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_APPLICATIONS);

  if (!sheet) {
    throw new Error(`Sheet not found: ${SHEET_APPLICATIONS}`);
  }

  return sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(h => String(h || '').trim());
}

/**
 * 応募No採番
 */
function generateApplicationNo_() {
  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_APPLICATIONS);

  if (!sheet) {
    throw new Error(`Sheet not found: ${SHEET_APPLICATIONS}`);
  }

  const headers =
    getApplicationHeaders_();

  const noCol =
    headers.indexOf('No') + 1;

  if (noCol <= 0) {
    throw new Error('応募管理シートに No 列がありません。');
  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow <= 2) {
    return 1;
  }

  const values =
    sheet
      .getRange(3, noCol, lastRow - 2, 1)
      .getValues();

  let maxNo = 0;

  values.forEach(row => {
    const no =
      Number(row[0] || 0);

    if (!isNaN(no) && no > maxNo) {
      maxNo = no;
    }
  });

  return maxNo + 1;
}

/**
 * 応募済み判定
 */
function hasCompletedApplication_(userId) {
  return countCompletedApplicationsByUserId_(userId) > 0;
}

/**
 * 同一ユーザーIDの応募完了数を取得
 */
function countCompletedApplicationsByUserId_(userId) {
  if (!userId) return 0;

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_APPLICATIONS);

  if (!sheet) return 0;

  const headers =
    getApplicationHeaders_();

  const userIdCol =
    headers.indexOf('ユーザーID') + 1;

  const statusCol =
    headers.indexOf('ステータス') + 1;

  if (userIdCol <= 0) return 0;

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 3) return 0;

  const values =
    sheet
      .getRange(3, 1, lastRow - 2, sheet.getLastColumn())
      .getValues();

  let count = 0;

  values.forEach(row => {
    const rowUserId =
      String(row[userIdCol - 1] || '').trim();

    const rowStatus =
      statusCol > 0
        ? String(row[statusCol - 1] || '').trim()
        : '';

    if (
      rowUserId === userId &&
      (
        !rowStatus ||
        rowStatus === STATUS_DONE ||
        rowStatus === '完了'
      )
    ) {
      count++;
    }
  });

  return count;
}

/**
 * 指定ユーザーIDの最新応募Noを取得
 */
function getLatestApplicationNoByUserId_(userId) {
  if (!userId) return '';

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_APPLICATIONS);

  if (!sheet) return '';

  const headers =
    getApplicationHeaders_();

  const noCol =
    headers.indexOf('No') + 1;

  const userIdCol =
    headers.indexOf('ユーザーID') + 1;

  if (noCol <= 0 || userIdCol <= 0) {
    return '';
  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 3) return '';

  const values =
    sheet
      .getRange(3, 1, lastRow - 2, sheet.getLastColumn())
      .getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    const rowUserId =
      String(values[i][userIdCol - 1] || '').trim();

    if (rowUserId === userId) {
      return values[i][noCol - 1] || '';
    }
  }

  return '';
}

/**
 * 応募Noから応募管理の行データを取得
 *
 * 注意：
 * この関数は 12_ApplicationSheet.gs にだけ残す。
 * 9_AdminRemind.gs 側には同名関数を置かない。
 */
function getApplicationRowDataByNo_(applicationNo) {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(SHEET_APPLICATIONS);

  if (!sheet) {
    throw new Error(`シートが見つかりません: ${SHEET_APPLICATIONS}`);
  }

  const lastRow =
    sheet.getLastRow();

  const lastCol =
    sheet.getLastColumn();

  if (lastRow < 3) {
    return null;
  }

  const headers =
    sheet
      .getRange(1, 1, 1, lastCol)
      .getValues()[0]
      .map(h => String(h || '').trim());

  const noCol =
    headers.indexOf('No') + 1;

  if (noCol <= 0) {
    throw new Error('応募管理シートに No 列がありません。');
  }

  const values =
    sheet
      .getRange(3, 1, lastRow - 2, lastCol)
      .getValues();

  const targetNo =
    String(applicationNo || '').trim();

  for (let i = 0; i < values.length; i++) {
    const row =
      values[i];

    const rowNo =
      String(row[noCol - 1] || '').trim();

    if (rowNo === targetNo) {
      const data = {};

      headers.forEach((header, index) => {
        if (header) {
          data[header] = row[index];
        }
      });

      return data;
    }
  }

  return null;
}