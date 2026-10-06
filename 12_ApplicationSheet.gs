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
 *
 * 冪等性対策（同一応募の二重保存防止・2026/10改修：応募管理シートを
 * 最終Source of Truthとする方式）：
 * - ロック取得直後に、まず応募管理シート側の完了履歴
 *   （hasCompletedApplication_）を確認する。既に1件でも存在する場合は
 *   DONEマーカーの状態に関係なく、新しい応募No採番・行追加を一切
 *   行わず null を返して処理を打ち切る。DONEマーカーの保存に万一
 *   失敗していても、応募管理シート側の記録だけで二重保存を防げる
 *   ようにするための最終防衛線。
 *   ※現在のV2仕様では「完了済みユーザーの再応募」は許可しないため、
 *   この判定は常に有効（将来、正式な再応募機能を実装する際は
 *   別途設計し直す）。
 * - 応募管理シートに履歴が無い場合のみ、続けて userState:${userId} の
 *   最新状態を再取得する。
 *   1) stateが存在しない場合：応募途中state自体が無いため保存処理を
 *      続行せず、応募行を新規作成しない（ログのみ残す）。
 *   2) state.status === STATUS_DONE（DONEマーカー）の場合：同一ユーザーの
 *      同一完了処理が別の実行（Webhook再送等）で既に保存・マーカー更新
 *      まで完了済みとみなし、新しい応募No採番・行追加を行わず
 *      null を返して処理を打ち切る（対応管理・対応履歴・管理者通知も
 *      呼び出し元 saveCompletedApplication_ 側で行われない）。
 *      ※応募管理シートの履歴チェックを必ず先に行うため、この分岐に
 *      実際に到達するのは「シートには未反映だがDONEマーカーは
 *      存在する」という通常起こりえないケースのみだが、
 *      DONEマーカーとの役割分担を明確にするため残している。
 *   3) それ以外（応募途中state）の場合：通常通り採番・保存を行う。
 * - 応募管理シートへの行書き込み・対応管理行追加が成功した直後、
 *   同一ロック内で markUserDone_() によりDONEマーカーへ置き換える。
 *   これにより「保存成功 → DONEマーカー化」を実質アトミックにし、
 *   ロック待ちしていたもう一方の重複実行が、ロック取得時点で
 *   確実に重複を検知できる。
 * - これはLINEのWebhook再送（同一イベントの再送信）による
 *   二重処理を主な想定ケースとした冪等性ガードであり、
 *   ユーザーが手動で同じ回答を複数回送信した場合（LINE側で
 *   別イベントとして配信されるケース）までは防げない点に注意。
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
    if (hasCompletedApplication_(userId)) {
      // 応募管理シートを二重保存判定の最終Source of Truthとする。
      // DONEマーカーの状態に関係なく、シート側に既に完了履歴が
      // あれば新しい応募行は作らない。
      errorLog_(
        'APPLICATION_DUPLICATE_SAVE_SKIPPED',
        '応募管理シートに既に完了履歴が存在するため、重複保存をスキップしました（最終Source of Truthによる判定）。',
        userId,
        ''
      );

      return null;
    }

    const liveState =
      getUserState_(userId);

    if (!liveState) {
      // 応募途中stateが存在しない（Webhook重複等により既に処理済みで
      // 削除・変化した可能性を含む）。保存処理を続行しない。
      //
      // 【注】PROCESS_LOG_ALLOWED_TYPES（1_Config.gs）に含まれない
      // 種別はprocessLog_が無言でスキップするため、ここではerrorLog_
      // （常に記録される）を使用する。1_Config.gsは今回の変更対象外の
      // ため、許可リストへの追加は行わない。
      errorLog_(
        'APPLICATION_SAVE_SKIPPED_NO_STATE',
        '保存時点でuserStateが存在しないため、保存をスキップしました。',
        userId,
        ''
      );

      return null;
    }

    if (liveState.status === STATUS_DONE) {
      // 既にDONEマーカー化されている＝同一ユーザーの同一完了処理が
      // 別の実行で既に保存済みとみなし、重複した応募行を作らず
      // スキップする（Webhook再送等への冪等性対策）。
      //
      // 【注】同上の理由でerrorLog_を使用する（processLog_は許可リスト外
      // の種別を無言でスキップするため）。
      errorLog_(
        'APPLICATION_DUPLICATE_SAVE_SKIPPED',
        `既に応募No.${liveState.applicationNo || ''}で保存済み（DONEマーカー）のため、重複保存をスキップしました。`,
        userId,
        ''
      );

      return null;
    }

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

    try {
      markUserDone_(userId, applicationNo);
    } catch (cleanupError) {
      // DONEマーカーへの置き換えに失敗しても、応募管理シートへの
      // 保存自体は既に成功しているため、保存失敗
      // （APPLICATION_SAVE_ERROR）として扱わない。ログにのみ残す。
      errorLog_(
        'USER_STATE_CLEANUP_ERROR',
        cleanupError.stack || cleanupError.message,
        userId,
        `応募No.${applicationNo}`
      );
    }

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