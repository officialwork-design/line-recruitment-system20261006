/**
 * 14_SupportHistory.gs
 * 対応履歴管理
 *
 * 方針：
 * - 対応履歴は追記専用
 * - 面接 / 体入 / 本入 の入力履歴を保存
 * - 過去データ反映も維持
 * - ヘッダー名ベースで保存し、列順変更に耐える
 */

/**
 * 対応履歴を保存
 */
function saveSupportHistory_(params) {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    getOrCreateSheet_(
      ss,
      SHEET_SUPPORT_HISTORY
    );

  ensureSupportHistoryHeader_(sheet);

  const headers =
    getSheetHeaders_(sheet);

  const row =
    buildSupportHistoryRow_(
      headers,
      params || {}
    );

  const targetRow =
    Math.max(sheet.getLastRow() + 1, 2);

  sheet
    .getRange(targetRow, 1, 1, row.length)
    .setValues([row]);

  try {
    sheet.hideSheet();
  } catch (error) {
    // 非表示にできなくても処理継続
  }
}

/**
 * 対応履歴ヘッダー保証
 */
function ensureSupportHistoryHeader_(sheet) {
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
}

/**
 * 対応履歴1行を作成
 */
function buildSupportHistoryRow_(
  headers,
  params
) {
  const row =
    new Array(headers.length).fill('');

  setRowValueByHeader_(headers, row, '受信日時', new Date());
  setRowValueByHeader_(headers, row, 'ユーザーID', params.userId || '');
  setRowValueByHeader_(headers, row, 'LINE表示名', params.displayName || '');
  setRowValueByHeader_(headers, row, '種別', params.type || '');
  setRowValueByHeader_(headers, row, 'メッセージ内容', params.message || '');
  setRowValueByHeader_(headers, row, '対応管理No', params.supportNo || '');
  setRowValueByHeader_(headers, row, '関連応募No', params.applicationNo || '');
  setRowValueByHeader_(headers, row, '更新日', new Date());

  return row;
}

/**
 * 面接 / 体入 / 本入 の入力履歴を対応履歴へ保存
 */
function saveSupportActionLogFromEdit_(
  supportSheet,
  supportHeaders,
  editedRow,
  supportNo,
  editedHeader,
  newValue
) {
  const value =
    String(newValue || '').trim();

  if (!value) return;

  const displayName =
    getSupportDisplayNameFromRow_(
      supportSheet,
      supportHeaders,
      editedRow
    );

  saveSupportHistory_({
    userId: '',
    displayName,
    type: `${editedHeader}入力`,
    message: `${editedHeader}：${value}`,
    supportNo,
    applicationNo: supportNo
  });
}

/**
 * 対応管理の行から表示名/名前を取得
 */
function getSupportDisplayNameFromRow_(
  supportSheet,
  supportHeaders,
  row
) {
  const lineNameCol =
    supportHeaders.indexOf('LINE表示名') + 1;

  const nameCol =
    supportHeaders.indexOf('名前') + 1;

  const displayName =
    lineNameCol > 0
      ? supportSheet
          .getRange(row, lineNameCol)
          .getValue()
      : '';

  const applicantName =
    nameCol > 0
      ? supportSheet
          .getRange(row, nameCol)
          .getValue()
      : '';

  return String(displayName || applicantName || '').trim();
}

/**
 * 既存の 面接 / 体入 / 本入 を対応履歴へ反映
 *
 * 注意：
 * - 反映日時は「実行した日時」
 * - 何度も実行すると重複する
 */
function backfillSupportActionHistory() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const supportSheet =
    ss.getSheetByName(SHEET_SUPPORT);

  if (!supportSheet) {
    SpreadsheetApp
      .getUi()
      .alert('対応管理シートが見つかりません。');
    return;
  }

  const headers =
    getSheetHeaders_(supportSheet);

  const requiredHeaders = [
    'No',
    '面接',
    '体入',
    '本入'
  ];

  const missingHeaders =
    requiredHeaders.filter(header => {
      return headers.indexOf(header) === -1;
    });

  if (missingHeaders.length > 0) {
    SpreadsheetApp
      .getUi()
      .alert(
        '必要な列が不足しています。\n' +
        missingHeaders.join(' / ')
      );
    return;
  }

  const lastRow =
    supportSheet.getLastRow();

  if (lastRow < 3) {
    SpreadsheetApp
      .getUi()
      .alert('反映対象データがありません。');
    return;
  }

  const confirm =
    SpreadsheetApp
      .getUi()
      .alert(
        '過去履歴反映の確認',
        [
          '対応管理の 面接 / 体入 / 本入 の既存入力を、対応履歴へ反映します。',
          '',
          '注意：この処理は追記方式です。',
          '同じデータに対して何度も実行すると履歴が重複します。',
          '',
          '実行しますか？'
        ].join('\n'),
        SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
      );

  if (confirm !== SpreadsheetApp.getUi().Button.OK) {
    return;
  }

  const values =
    supportSheet
      .getRange(
        3,
        1,
        lastRow - 2,
        supportSheet.getLastColumn()
      )
      .getValues();

  const createdCount =
    createBackfillSupportHistories_(
      supportSheet,
      headers,
      values
    );

  SpreadsheetApp
    .getUi()
    .alert(
      `過去データ反映が完了しました。\n作成件数：${createdCount}件`
    );
}

/**
 * 過去データから対応履歴を作成
 */
function createBackfillSupportHistories_(
  supportSheet,
  headers,
  values
) {
  const noCol =
    headers.indexOf('No');

  const interviewCol =
    headers.indexOf('面接');

  const trialCol =
    headers.indexOf('体入');

  const joinCol =
    headers.indexOf('本入');

  let createdCount = 0;

  values.forEach(row => {
    const supportNo =
      String(row[noCol] || '').trim();

    if (!supportNo) return;

    if (supportNo.indexOf('問い合わせ-') === 0) {
      return;
    }

    const name =
      pickSupportNameFromRow_(
        headers,
        row
      );

    const historyTargets = [
      {
        header: '面接',
        value: String(row[interviewCol] || '').trim()
      },
      {
        header: '体入',
        value: String(row[trialCol] || '').trim()
      },
      {
        header: '本入',
        value: String(row[joinCol] || '').trim()
      }
    ];

    historyTargets.forEach(target => {
      if (!target.value) return;

      saveSupportHistory_({
        userId: '',
        displayName: name,
        type: `${target.header}入力`,
        message: `${target.header}：${target.value}（過去データ反映）`,
        supportNo,
        applicationNo: supportNo
      });

      createdCount++;
    });
  });

  return createdCount;
}

/**
 * 対応管理の行配列から名前を取得
 */
function pickSupportNameFromRow_(
  headers,
  row
) {
  const lineNameIndex =
    headers.indexOf('LINE表示名');

  const nameIndex =
    headers.indexOf('名前');

  const displayName =
    lineNameIndex >= 0
      ? String(row[lineNameIndex] || '').trim()
      : '';

  const applicantName =
    nameIndex >= 0
      ? String(row[nameIndex] || '').trim()
      : '';

  return displayName || applicantName || '';
}