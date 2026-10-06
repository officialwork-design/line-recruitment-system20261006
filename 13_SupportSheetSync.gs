/**
 * 13_SupportSheetSync.gs
 * 対応管理 ⇄ 応募管理 同期
 *
 * 方針：
 * - 対応管理で編集した内容を応募管理へ同期
 * - インストール型 onEdit トリガーのみ使用
 * - 複数セル編集は事故防止のためスキップ
 * - 不合格時は 面接 / 体入 / 本入 を自動で × にする
 * - 応募管理 → 対応管理 の手動同期も維持
 */

/**
 * 応募完了時に対応管理へ行追加
 */
function addSupportRowFromApplication_(
  applicationNo,
  state,
  userId,
  displayName,
  supportMeta
) {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(SHEET_SUPPORT);

  if (!sheet) {
    errorLog_(
      'SUPPORT_SHEET_NOT_FOUND',
      `対応管理シートが見つかりません。応募No=${applicationNo}`,
      userId,
      displayName
    );
    return;
  }

  const headers =
    getSheetHeaders_(sheet);

  const noIndex =
    headers.indexOf('No');

  if (noIndex === -1) {
    errorLog_(
      'SUPPORT_NO_HEADER_NOT_FOUND',
      '対応管理シートに No 列がありません。',
      userId,
      displayName
    );
    return;
  }

  if (supportNoExists_(sheet, headers, applicationNo)) {
    return;
  }

  const row =
    buildSupportRowFromApplication_(
      headers,
      applicationNo,
      state,
      displayName,
      supportMeta
    );

  const targetRow =
    Math.max(sheet.getLastRow() + 1, 3);

  sheet
    .getRange(targetRow, 1, 1, row.length)
    .setValues([row]);
}

/**
 * 対応管理に同じNoがあるか確認
 */
function supportNoExists_(
  sheet,
  headers,
  supportNo
) {
  const noCol =
    headers.indexOf('No') + 1;

  if (noCol <= 0) return false;

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 3) return false;

  const values =
    sheet
      .getRange(3, noCol, lastRow - 2, 1)
      .getValues();

  const target =
    String(supportNo || '').trim();

  return values.some(row => {
    return String(row[0] || '').trim() === target;
  });
}

/**
 * 応募データから対応管理行を作成
 */
function buildSupportRowFromApplication_(
  headers,
  applicationNo,
  state,
  displayName,
  supportMeta
) {
  const row =
    new Array(headers.length).fill('');

  const applicationHistoryStatus =
    supportMeta && supportMeta.applicationHistoryStatus
      ? supportMeta.applicationHistoryStatus
      : '';

  const applicationCount =
    supportMeta && supportMeta.applicationCount
      ? supportMeta.applicationCount
      : '';

  setRowValueByHeader_(headers, row, 'No', applicationNo);
  setRowValueByHeader_(headers, row, 'ステータス', SUPPORT_STATUS_NOT_STARTED);
  setRowValueByHeader_(headers, row, '面接担当', '');
  setRowValueByHeader_(headers, row, 'LINE表示名', displayName || state.displayName || '');
  setRowValueByHeader_(headers, row, '応募メッセージ', state.applicationMessage || '');
  setRowValueByHeader_(headers, row, '名前', '');
  setRowValueByHeader_(headers, row, '合否', '');
  setRowValueByHeader_(headers, row, '面接', '');
  setRowValueByHeader_(headers, row, '体入', '');
  setRowValueByHeader_(headers, row, '本入', '');
  setRowValueByHeader_(headers, row, '対応メモ', '');
  setRowValueByHeader_(headers, row, '過去応募者', applicationHistoryStatus);
  setRowValueByHeader_(headers, row, '応募回数', applicationCount);
  setRowValueByHeader_(headers, row, '更新日', new Date());

  return row;
}

/**
 * 対応ステータス更新
 */
function updateSupportStatusByNo_(
  supportNo,
  status
) {
  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_SUPPORT);

  if (!sheet) {
    errorLog_(
      'SUPPORT_SHEET_NOT_FOUND',
      '対応管理シートが見つかりません。',
      '',
      ''
    );
    return false;
  }

  const headers =
    getSheetHeaders_(sheet);

  const noCol =
    headers.indexOf('No') + 1;

  const statusCol =
    headers.indexOf('ステータス') + 1;

  const updatedAtCol =
    headers.indexOf('更新日') + 1;

  if (noCol <= 0 || statusCol <= 0) {
    errorLog_(
      'SUPPORT_STATUS_HEADER_NOT_FOUND',
      '対応管理シートに No または ステータス 列がありません。',
      '',
      ''
    );
    return false;
  }

  const targetRow =
    findRowByHeaderValue_(
      sheet,
      'No',
      supportNo,
      3
    );

  if (!targetRow) {
    return false;
  }

  sheet
    .getRange(targetRow, statusCol)
    .setValue(status);

  if (updatedAtCol > 0) {
    sheet
      .getRange(targetRow, updatedAtCol)
      .setValue(new Date());
  }

  return true;
}

/**
 * 対応管理 → 応募管理 同期トリガー設定
 *
 * 注意：
 * function onEdit(e) は作らない。
 * インストール型トリガーのみ使う。
 */
function installSupportApplicationSyncTrigger() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const triggers =
    ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (
      trigger.getHandlerFunction() ===
      'handleSupportApplicationSyncOnEdit_'
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp
    .newTrigger('handleSupportApplicationSyncOnEdit_')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  SpreadsheetApp
    .getUi()
    .alert('対応管理 → 応募管理の同期トリガーを設定しました。');
}

/**
 * インストール型 onEdit 入口
 */
function handleSupportApplicationSyncOnEdit_(e) {
  if (!e || !e.range) return;

  try {
    syncSupportEditToApplication_(e);
  } catch (error) {
    try {
      errorLog_(
        'ON_EDIT_SYNC_ERROR',
        error.stack || error.message,
        '',
        ''
      );
    } catch (logError) {}
  }
}

/**
 * 対応管理 → 応募管理 同期
 */
function syncSupportEditToApplication_(e) {
  if (!e || !e.range) return;

  const range =
    e.range;

  const supportSheet =
    range.getSheet();

  if (
    !supportSheet ||
    supportSheet.getName() !== SHEET_SUPPORT
  ) {
    return;
  }

  const editedRow =
    range.getRow();

  const editedCol =
    range.getColumn();

  if (editedRow < 3) return;

  if (
    range.getNumRows() !== 1 ||
    range.getNumColumns() !== 1
  ) {
    errorLog_(
      'SYNC_SKIPPED_MULTI_CELL_EDIT',
      `複数セル編集のためスキップ row=${editedRow}, col=${editedCol}`,
      '',
      ''
    );
    return;
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const applicationSheet =
    ss.getSheetByName(SHEET_APPLICATIONS);

  if (!applicationSheet) {
    errorLog_(
      'SYNC_APPLICATION_SHEET_NOT_FOUND',
      '応募管理シートが見つかりません。',
      '',
      ''
    );
    return;
  }

  const supportHeaders =
    getSheetHeaders_(supportSheet)
      .map(normalizeSyncHeader_);

  const applicationHeaders =
    getSheetHeaders_(applicationSheet)
      .map(normalizeSyncHeader_);

  const editedHeader =
    supportHeaders[editedCol - 1];

  const syncHeaders =
    getSupportSyncHeaders_();

  if (!syncHeaders.includes(editedHeader)) {
    return;
  }

  const supportNo =
    getSupportNoFromRow_(
      supportSheet,
      supportHeaders,
      editedRow
    );

  if (!supportNo) return;

  if (supportNo.indexOf('問い合わせ-') === 0) {
    return;
  }

  const applicationTarget =
    findApplicationTargetByNo_(
      applicationSheet,
      applicationHeaders,
      supportNo,
      editedHeader
    );

  if (!applicationTarget.row || !applicationTarget.col) {
    return;
  }

  const newValue =
    range.getValue();

  applicationSheet
    .getRange(applicationTarget.row, applicationTarget.col)
    .setValue(newValue);

  if (
    ['面接', '体入', '本入'].includes(editedHeader)
  ) {
    saveSupportActionLogFromEdit_(
      supportSheet,
      supportHeaders,
      editedRow,
      supportNo,
      editedHeader,
      newValue
    );
  }

  if (
    editedHeader === '合否' &&
    String(newValue || '').trim() === '不合格'
  ) {
    applyRejectedAutoMarks_(
      supportSheet,
      supportHeaders,
      editedRow,
      applicationSheet,
      applicationHeaders,
      applicationTarget.row,
      supportNo
    );
  }

  syncApplicationUpdatedAt_(
    applicationSheet,
    applicationHeaders,
    applicationTarget.row
  );

  syncSupportUpdatedAt_(
    supportSheet,
    supportHeaders,
    editedRow
  );
}

/**
 * 同期対象ヘッダー
 */
function getSupportSyncHeaders_() {
  return [
    '面接担当',
    '名前',
    '合否',
    '面接',
    '体入',
    '本入'
  ];
}

/**
 * 対応管理の行からNo取得
 */
function getSupportNoFromRow_(
  supportSheet,
  supportHeaders,
  row
) {
  const supportNoCol =
    supportHeaders.indexOf('No') + 1;

  if (supportNoCol <= 0) {
    errorLog_(
      'SYNC_SUPPORT_NO_HEADER_NOT_FOUND',
      '対応管理シートに No 列がありません。',
      '',
      ''
    );
    return '';
  }

  return String(
    supportSheet
      .getRange(row, supportNoCol)
      .getValue() || ''
  ).trim();
}

/**
 * 応募管理側の同期対象セルを探す
 */
function findApplicationTargetByNo_(
  applicationSheet,
  applicationHeaders,
  applicationNo,
  targetHeader
) {
  const applicationNoCol =
    applicationHeaders.indexOf('No') + 1;

  const applicationTargetCol =
    applicationHeaders.indexOf(targetHeader) + 1;

  if (
    applicationNoCol <= 0 ||
    applicationTargetCol <= 0
  ) {
    errorLog_(
      'SYNC_APPLICATION_HEADER_NOT_FOUND',
      `応募管理シートに No または ${targetHeader} 列がありません。`,
      '',
      ''
    );

    return {
      row: 0,
      col: 0
    };
  }

  const targetRow =
    findRowByHeaderValue_(
      applicationSheet,
      'No',
      applicationNo,
      3
    );

  if (!targetRow) {
    errorLog_(
      'SYNC_APPLICATION_ROW_NOT_FOUND',
      `応募管理シートに No=${applicationNo} が見つかりません。`,
      '',
      ''
    );

    return {
      row: 0,
      col: 0
    };
  }

  return {
    row: targetRow,
    col: applicationTargetCol
  };
}

/**
 * 合否が不合格になった時、
 * 面接 / 体入 / 本入 を自動で × にして、対応履歴にも残す。
 */
function applyRejectedAutoMarks_(
  supportSheet,
  supportHeaders,
  supportRow,
  applicationSheet,
  applicationHeaders,
  applicationRow,
  supportNo
) {
  const targets = [
    '面接',
    '体入',
    '本入'
  ];

  targets.forEach(header => {
    const supportCol =
      supportHeaders.indexOf(header) + 1;

    const applicationCol =
      applicationHeaders.indexOf(header) + 1;

    if (supportCol > 0) {
      supportSheet
        .getRange(supportRow, supportCol)
        .setValue('×');
    }

    if (applicationCol > 0) {
      applicationSheet
        .getRange(applicationRow, applicationCol)
        .setValue('×');
    }

    saveSupportActionLogFromEdit_(
      supportSheet,
      supportHeaders,
      supportRow,
      supportNo,
      header,
      '×（不合格により自動反映）'
    );
  });
}

/**
 * 応募管理 → 対応管理 手動同期
 */
function syncApplicationsToSupportManual() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const applicationSheet =
    ss.getSheetByName(SHEET_APPLICATIONS);

  const supportSheet =
    ss.getSheetByName(SHEET_SUPPORT);

  if (!applicationSheet || !supportSheet) {
    SpreadsheetApp
      .getUi()
      .alert(
        `シートが見つかりません。\n応募管理: ${!!applicationSheet}\n対応管理: ${!!supportSheet}`
      );
    return;
  }

  const appLastRow =
    applicationSheet.getLastRow();

  const supportLastRow =
    supportSheet.getLastRow();

  if (appLastRow < 3 || supportLastRow < 3) {
    SpreadsheetApp
      .getUi()
      .alert('同期対象のデータがありません。');
    return;
  }

  const appHeaders =
    getSheetHeaders_(applicationSheet)
      .map(normalizeSyncHeader_);

  const supportHeaders =
    getSheetHeaders_(supportSheet)
      .map(normalizeSyncHeader_);

  const syncHeaders =
    getSupportSyncHeaders_();

  const appNoCol =
    appHeaders.indexOf('No') + 1;

  const supportNoCol =
    supportHeaders.indexOf('No') + 1;

  if (appNoCol <= 0 || supportNoCol <= 0) {
    SpreadsheetApp
      .getUi()
      .alert('応募管理または対応管理に No 列がありません。');
    return;
  }

  const missingHeaders =
    syncHeaders.filter(header => {
      return (
        appHeaders.indexOf(header) === -1 ||
        supportHeaders.indexOf(header) === -1
      );
    });

  if (missingHeaders.length > 0) {
    SpreadsheetApp
      .getUi()
      .alert(
        '同期に必要な列が不足しています。\n' +
        missingHeaders.join(' / ')
      );
    return;
  }

  const appValues =
    applicationSheet
      .getRange(3, 1, appLastRow - 2, applicationSheet.getLastColumn())
      .getValues();

  const supportValues =
    supportSheet
      .getRange(3, 1, supportLastRow - 2, supportSheet.getLastColumn())
      .getValues();

  const supportRowMap =
    buildSupportRowMap_(
      supportValues,
      supportNoCol
    );

  let updatedCount = 0;
  let skippedCount = 0;
  let notFoundCount = 0;

  appValues.forEach(appRow => {
    const applicationNo =
      String(appRow[appNoCol - 1] || '').trim();

    if (!applicationNo) {
      skippedCount++;
      return;
    }

    const supportTargetRow =
      supportRowMap[applicationNo];

    if (!supportTargetRow) {
      notFoundCount++;
      return;
    }

    const changed =
      syncApplicationRowToSupportRow_(
        appRow,
        applicationSheet,
        appHeaders,
        supportSheet,
        supportHeaders,
        supportTargetRow,
        syncHeaders
      );

    if (changed) {
      syncSupportUpdatedAt_(
        supportSheet,
        supportHeaders,
        supportTargetRow
      );

      updatedCount++;
    } else {
      skippedCount++;
    }
  });

  SpreadsheetApp
    .getUi()
    .alert(
      [
        '応募管理 → 対応管理 の手動同期が完了しました。',
        '',
        `更新：${updatedCount}件`,
        `変更なし：${skippedCount}件`,
        `対応管理にNoなし：${notFoundCount}件`
      ].join('\n')
    );
}

/**
 * 対応管理No → 行番号Map作成
 */
function buildSupportRowMap_(
  supportValues,
  supportNoCol
) {
  const map = {};

  supportValues.forEach((row, index) => {
    const no =
      String(row[supportNoCol - 1] || '').trim();

    if (no && no.indexOf('問い合わせ-') !== 0) {
      map[no] = index + 3;
    }
  });

  return map;
}

/**
 * 応募管理1行を対応管理1行へ同期
 */
function syncApplicationRowToSupportRow_(
  appRow,
  applicationSheet,
  appHeaders,
  supportSheet,
  supportHeaders,
  supportTargetRow,
  syncHeaders
) {
  let rowChanged = false;

  syncHeaders.forEach(header => {
    const appCol =
      appHeaders.indexOf(header) + 1;

    const supportCol =
      supportHeaders.indexOf(header) + 1;

    if (appCol <= 0 || supportCol <= 0) {
      return;
    }

    const newValue =
      appRow[appCol - 1];

    const currentValue =
      supportSheet
        .getRange(supportTargetRow, supportCol)
        .getValue();

    if (
      String(currentValue || '') !==
      String(newValue || '')
    ) {
      supportSheet
        .getRange(supportTargetRow, supportCol)
        .setValue(newValue);

      rowChanged = true;
    }
  });

  return rowChanged;
}

/**
 * 応募管理 更新日同期
 */
function syncApplicationUpdatedAt_(
  applicationSheet,
  applicationHeaders,
  targetRow
) {
  const updatedAtCol =
    applicationHeaders.indexOf('更新日') + 1;

  if (updatedAtCol > 0) {
    applicationSheet
      .getRange(targetRow, updatedAtCol)
      .setValue(new Date());
  }
}

/**
 * 対応管理 更新日同期
 */
function syncSupportUpdatedAt_(
  supportSheet,
  supportHeaders,
  targetRow
) {
  const updatedAtCol =
    supportHeaders.indexOf('更新日') + 1;

  if (updatedAtCol > 0) {
    supportSheet
      .getRange(targetRow, updatedAtCol)
      .setValue(new Date());
  }
}