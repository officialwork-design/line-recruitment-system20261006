/**
 * 15_Contacts.gs
 * 連絡先管理
 *
 * 方針：
 * - 既存データは保持する
 * - ヘッダー名ベースで更新する
 * - 追客ステータスを自動更新する
 */

/**
 * 友だち追加時に連絡先を登録/更新
 */
function upsertContactOnFollow_(
  userId,
  displayName
) {
  if (!userId) return;

  upsertContact_({
    userId,
    displayName,
    status: '友だち追加済み',
    applyStatus: '未応募',
    followupStatusIfEmpty: '未開始',
    latestApplicationNo: '',
    updateFollowAt: true,
    updateLastMessageAt: false
  });
}

/**
 * メッセージ受信時に連絡先を登録/更新
 *
 * 未応募者から返信が来た場合は追客停止。
 */
function updateContactOnMessage_(
  userId,
  displayName
) {
  if (!userId) return;

  upsertContact_({
    userId,
    displayName,
    status: 'やりとりあり',
    applyStatus: '',
    followupStatus: '停止',
    latestApplicationNo: '',
    updateFollowAt: true,
    updateLastMessageAt: true,
    stopFollowupOnlyWhenUnapplied: true
  });
}

/**
 * 応募状況を更新
 *
 * 応募中・応募完了になったら追客停止。
 */
function updateContactApplicationStatus_(
  userId,
  status,
  applicationNo
) {
  if (!userId) return;

  upsertContact_({
    userId,
    displayName: '',
    status: status === '応募完了'
      ? '応募完了'
      : '応募中',
    applyStatus: status,
    followupStatus: '停止',
    latestApplicationNo: applicationNo || '',
    updateFollowAt: true,
    updateLastMessageAt: false,
    stopFollowupOnlyWhenUnapplied: false
  });
}

/**
 * 連絡先シートのヘッダー配列
 * （11_SetupSpreadsheet.gs の setupContactSheetForRecruit_ と
 * 必ず一致させること）
 */
function getContactHeaderDefinition_() {
  return [
    'ユーザーID', 'LINE表示名', '友だち追加日時', '最終やりとり日時',
    '状態', '応募状況', '追客ステータス', '最新応募No', 'メモ', '更新日'
  ];
}

/**
 * 連絡先の共通登録/更新処理
 *
 * V2移行（2026/10）：連絡先シートの旧データ領域（本番からコピーされた
 * 可能性のある既存の連絡先）には一切書き込まない。新規の登録・更新は
 * 必ずV2ヘッダー以降のデータ領域内でのみ行う（同一ユーザーIDの行が
 * 旧データ領域に存在していても、それは更新せず、V2データ領域に
 * 新しい行を作成する）。
 */
function upsertContact_(params) {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    getOrCreateSheet_(
      ss,
      SHEET_CONTACTS
    );

  ensureContactHeader_();

  const block =
    ensureV2HeaderBlock_(
      sheet,
      getContactHeaderDefinition_(),
      getContactNoteRowDefinition_()
    );

  const headerRow =
    block.headerRow;

  const dataStartRow =
    block.dataStartRow;

  const headers =
    sheet
      .getRange(headerRow, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(h => String(h || '').trim());

  const targetRow =
    findContactRowByUserId_(
      sheet,
      headers,
      params.userId,
      dataStartRow
    ) ||
    Math.max(sheet.getLastRow() + 1, dataStartRow);

  const existingRow =
    getContactRowValues_(
      sheet,
      targetRow,
      headers.length
    );

  const row =
    existingRow.length > 0
      ? existingRow
      : new Array(headers.length).fill('');

  const now =
    new Date();

  setContactRowValue_(
    headers,
    row,
    'ユーザーID',
    params.userId
  );

  if (params.displayName) {
    setContactRowValue_(
      headers,
      row,
      'LINE表示名',
      params.displayName
    );
  }

  const currentFollowAt =
    getContactRowValue_(
      headers,
      row,
      '友だち追加日時'
    );

  if (
    params.updateFollowAt &&
    !currentFollowAt
  ) {
    setContactRowValue_(
      headers,
      row,
      '友だち追加日時',
      now
    );
  }

  if (params.updateLastMessageAt) {
    setContactRowValue_(
      headers,
      row,
      '最終やりとり日時',
      now
    );
  }

  if (params.status) {
    setContactRowValue_(
      headers,
      row,
      '状態',
      params.status
    );
  }

  const currentApplyStatus =
    String(
      getContactRowValue_(
        headers,
        row,
        '応募状況'
      ) || ''
    ).trim();

  if (params.applyStatus) {
    setContactRowValue_(
      headers,
      row,
      '応募状況',
      params.applyStatus
    );
  } else if (!currentApplyStatus) {
    setContactRowValue_(
      headers,
      row,
      '応募状況',
      '未応募'
    );
  }

  if (params.followupStatus) {
    if (params.stopFollowupOnlyWhenUnapplied) {
      const applyStatusForJudge =
        String(
          params.applyStatus ||
          currentApplyStatus ||
          '未応募'
        ).trim();

      if (applyStatusForJudge === '未応募') {
        setContactRowValue_(
          headers,
          row,
          '追客ステータス',
          params.followupStatus
        );
      }
    } else {
      setContactRowValue_(
        headers,
        row,
        '追客ステータス',
        params.followupStatus
      );
    }
  }

  if (params.followupStatusIfEmpty) {
    const currentFollowupStatus =
      getContactRowValue_(
        headers,
        row,
        '追客ステータス'
      );

    if (!currentFollowupStatus) {
      setContactRowValue_(
        headers,
        row,
        '追客ステータス',
        params.followupStatusIfEmpty
      );
    }
  }

  if (params.latestApplicationNo) {
    setContactRowValue_(
      headers,
      row,
      '最新応募No',
      params.latestApplicationNo
    );
  }

  setContactRowValue_(
    headers,
    row,
    '更新日',
    now
  );

  sheet
    .getRange(targetRow, 1, 1, headers.length)
    .setValues([row]);
}

/**
 * 連絡先シートの説明行（備考）
 * （11_SetupSpreadsheet.gs の setupContactSheetForRecruit_ と
 * 必ず一致させること）
 */
function getContactNoteRowDefinition_() {
  return [
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
}

/**
 * 連絡先シートのヘッダーを保証
 *
 * V2移行（2026/10）：友だち追加・メッセージ受信のたびに毎回呼ばれる
 * 関数であるため、旧実装のように1〜2行目を無条件で上書きすることは
 * しない。ensureV2HeaderBlock_ により、既にV2ヘッダーが存在する場合は
 * 何もしない（旧データ領域には一切触れない）。
 */
function ensureContactHeader_() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    getOrCreateSheet_(
      ss,
      SHEET_CONTACTS
    );

  const headers =
    getContactHeaderDefinition_();

  const notes =
    getContactNoteRowDefinition_();

  const block =
    ensureV2HeaderBlock_(sheet, headers, notes);

  if (block.headerRow === 1) {
    sheet.setFrozenRows(2);
  }

  const formatRows =
    Math.max(sheet.getMaxRows() - block.headerRow + 1, 1);

  sheet
    .getRange(block.headerRow, 1, formatRows, headers.length)
    .setWrap(true);

  const dataRows =
    Math.max(sheet.getMaxRows() - block.dataStartRow + 1, 1);

  sheet
    .getRange(block.dataStartRow, 3, dataRows, 2)
    .setNumberFormat('yyyy/mm/dd hh:mm:ss');

  sheet
    .getRange(block.dataStartRow, 10, dataRows, 1)
    .setNumberFormat('yyyy/mm/dd hh:mm:ss');
}

/**
 * ユーザーIDから連絡先行を探す
 *
 * V2移行（2026/10）：startRow（通常はV2データ開始行）より前の行、
 * つまり旧データ領域は検索対象に含めない。旧データ領域に同一の
 * ユーザーIDが存在していても、それを更新対象にはしない（新しい行を
 * V2データ領域に作成する）。
 */
function findContactRowByUserId_(
  sheet,
  headers,
  userId,
  startRow
) {
  const userIdCol =
    headers.indexOf('ユーザーID') + 1;

  if (userIdCol <= 0) return 0;

  const searchStartRow =
    startRow || 3;

  const lastRow =
    sheet.getLastRow();

  if (lastRow < searchStartRow) return 0;

  const values =
    sheet
      .getRange(searchStartRow, userIdCol, lastRow - searchStartRow + 1, 1)
      .getValues();

  const targetUserId =
    String(userId || '').trim();

  for (let i = 0; i < values.length; i++) {
    const rowUserId =
      String(values[i][0] || '').trim();

    if (rowUserId === targetUserId) {
      return i + searchStartRow;
    }
  }

  return 0;
}

/**
 * 既存行を取得
 */
function getContactRowValues_(
  sheet,
  row,
  columnCount
) {
  if (!sheet || row < 1 || columnCount < 1) {
    return [];
  }

  const lastRow =
    sheet.getLastRow();

  if (row > lastRow) {
    return new Array(columnCount).fill('');
  }

  return sheet
    .getRange(row, 1, 1, columnCount)
    .getValues()[0];
}

/**
 * 行配列にヘッダー名で値をセット
 */
function setContactRowValue_(
  headers,
  row,
  headerName,
  value
) {
  const index =
    headers.indexOf(headerName);

  if (index === -1) return;

  row[index] = value;
}

/**
 * 行配列からヘッダー名で値を取得
 */
function getContactRowValue_(
  headers,
  row,
  headerName
) {
  const index =
    headers.indexOf(headerName);

  if (index === -1) return '';

  return row[index];
}

/**
 * 互換用：既存コードから呼ばれても壊れないように残す
 */
function setContactValue_(
  sheet,
  headers,
  row,
  headerName,
  value
) {
  const col =
    headers.indexOf(headerName) + 1;

  if (col <= 0) return;

  sheet
    .getRange(row, col)
    .setValue(value);
}

/**
 * 互換用：既存コードから呼ばれても壊れないように残す
 */
function getContactValue_(
  sheet,
  headers,
  row,
  headerName
) {
  const col =
    headers.indexOf(headerName) + 1;

  if (col <= 0) return '';

  return sheet
    .getRange(row, col)
    .getValue();
}