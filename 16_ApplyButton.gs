/**
 * 16_ApplyButton.gs
 * 応募ボタン送信
 *
 * 方針：
 * - 応募開始ボタンのFlex Messageを作成
 * - 設定USERID / 直接入力 / ユーザー管理選択行 へ送信
 * - 送信履歴はユーザー管理シートへ記録
 */

/**
 * 応募ボタンFlexを作成
 */
function createApplyButtonMessage_() {
  const startWord =
    getConfigValue_('START_WORD') ||
    DEFAULT_START_WORD;

  const title =
    getConfigValue_('APPLY_BUTTON_TITLE') ||
    '✨ ご応募はこちらからお願いします。';

  const note =
    getConfigValue_('APPLY_BUTTON_NOTE') ||
    '下のボタンから応募を開始できます。';

  const label =
    getConfigValue_('APPLY_BUTTON_LABEL') ||
    '応募する';

  const color =
    normalizeHexColor_(
      getConfigValue_('APPLY_BUTTON_COLOR') ||
      '#06C755'
    );

  return buildApplyButtonFlexMessage_({
    startWord,
    title,
    note,
    label,
    color
  });
}

/**
 * 応募ボタンFlex本体
 */
function buildApplyButtonFlexMessage_(params) {
  const label =
    params.label || '応募する';

  return {
    type: 'flex',
    altText: `✨ ${label}`,
    contents: {
      type: 'bubble',
      body: {
        type: 'box',
        layout: 'vertical',
        spacing: 'md',
        contents: [
          {
            type: 'text',
            text: params.title || '✨ ご応募はこちらからお願いします。',
            weight: 'bold',
            size: 'lg',
            wrap: true,
            align: 'center'
          },
          {
            type: 'text',
            text: params.note || '下のボタンから応募を開始できます。',
            size: 'sm',
            color: '#666666',
            wrap: true,
            align: 'center'
          },
          {
            type: 'button',
            style: 'primary',
            color: normalizeHexColor_(params.color || '#06C755'),
            height: 'sm',
            action: {
              type: 'message',
              label: `✨ ${label}`,
              text: params.startWord || DEFAULT_START_WORD
            }
          }
        ]
      }
    }
  };
}

/**
 * 応募ボタン案内を返信
 */
function sendApplyButtonGuide_(
  replyToken,
  userId,
  errorType
) {
  replyMessages_(
    replyToken,
    [
      {
        type: 'text',
        text: '✨ 応募をご希望の場合は、下のボタンから開始してください。'
      },
      createApplyButtonMessage_()
    ],
    errorType || 'APPLY_BUTTON_GUIDE_REPLY_ERROR',
    userId
  );
}

/**
 * 設定シートのUSERIDへ応募ボタン送信
 */
function sendApplyButtonToConfiguredUser() {
  const ui =
    SpreadsheetApp.getUi();

  const userId =
    String(
      getConfigValue_('APPLY_BUTTON_USER_ID') || ''
    ).trim();

  if (
    !userId ||
    userId === 'ここに送信先のLINEユーザーIDを入力'
  ) {
    ui.alert(
      '設定シートの APPLY_BUTTON_USER_ID に送信先USERIDを入力してください。'
    );
    return;
  }

  sendApplyButtonToUserId_(
    userId,
    '設定USERID',
    '',
    'PUSH_APPLY_BUTTON_CONFIGURED_ERROR'
  );

  ui.alert(
    `応募ボタンを送信しました。\nUSERID: ${userId}`
  );
}

/**
 * USERIDを入力して応募ボタン送信
 */
function sendApplyButtonToUser() {
  const ui =
    SpreadsheetApp.getUi();

  const prompt =
    ui.prompt(
      '応募ボタン送信',
      '送信先のLINEユーザーIDを入力してください。',
      ui.ButtonSet.OK_CANCEL
    );

  if (prompt.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const userId =
    String(prompt.getResponseText() || '').trim();

  if (!userId) {
    ui.alert('ユーザーIDが空です。');
    return;
  }

  sendApplyButtonToUserId_(
    userId,
    'USERID直接入力',
    '',
    'PUSH_APPLY_BUTTON_DIRECT_ERROR'
  );

  ui.alert('応募ボタンを送信しました。');
}

/**
 * ユーザー管理の選択行へ応募ボタン送信
 */
function sendApplyButtonToSelectedUser() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getActiveSheet();

  const ui =
    SpreadsheetApp.getUi();

  if (sheet.getName() !== SHEET_USER_MANAGEMENT) {
    ui.alert(
      'ユーザー管理シートで送信したい行を選択してください。'
    );
    return;
  }

  const row =
    sheet.getActiveRange().getRow();

  if (row <= 1) {
    ui.alert(
      '送信したいユーザーの行を選択してください。'
    );
    return;
  }

  ensureUserManagementHeader_();

  const headers =
    getSheetHeaders_(sheet);

  const userId =
    getCellValueByHeader_(
      sheet,
      headers,
      row,
      'ユーザーID'
    );

  const displayName =
    getCellValueByHeader_(
      sheet,
      headers,
      row,
      'LINE表示名'
    );

  const targetUserId =
    String(userId || '').trim();

  if (!targetUserId) {
    ui.alert('選択行にユーザーIDがありません。');
    return;
  }

  sendApplyButtonToUserId_(
    targetUserId,
    '選択ユーザー',
    String(displayName || '').trim(),
    'PUSH_APPLY_BUTTON_SELECTED_ERROR'
  );

  ui.alert(
    '選択したユーザーへ応募ボタンを送信しました。'
  );
}

/**
 * 応募ボタン送信共通処理
 */
function sendApplyButtonToUserId_(
  userId,
  method,
  displayName,
  errorType
) {
  pushMessages_(
    userId,
    [createApplyButtonMessage_()],
    errorType || 'PUSH_APPLY_BUTTON_ERROR'
  );

  markApplyButtonSent_(
    userId,
    method,
    displayName
  );
}

/**
 * 応募ボタン送信履歴をユーザー管理へ記録
 */
function markApplyButtonSent_(
  userId,
  method,
  displayName
) {
  if (!userId) return;

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    getOrCreateSheet_(
      ss,
      SHEET_USER_MANAGEMENT
    );

  ensureUserManagementHeader_();

  const headers =
    getSheetHeaders_(sheet);

  const requiredHeaders = [
    'ユーザーID',
    '応募ボタン送信',
    '応募ボタン送信日時',
    '応募ボタン送信方法'
  ];

  const missingHeaders =
    requiredHeaders.filter(header => {
      return headers.indexOf(header) === -1;
    });

  if (missingHeaders.length > 0) {
    throw new Error(
      'ユーザー管理シートの応募ボタン送信管理列が不足しています: ' +
      missingHeaders.join(' / ')
    );
  }

  const targetRow =
    findRowByHeaderValue_(
      sheet,
      'ユーザーID',
      userId,
      2
    ) ||
    Math.max(sheet.getLastRow() + 1, 2);

  if (!findRowByHeaderValue_(sheet, 'ユーザーID', userId, 2)) {
    setCellValueByHeader_(
      sheet,
      headers,
      targetRow,
      'ユーザーID',
      userId
    );
  }

  if (displayName) {
    setCellValueByHeader_(
      sheet,
      headers,
      targetRow,
      'LINE表示名',
      displayName
    );
  }

  setCellValueByHeader_(
    sheet,
    headers,
    targetRow,
    '応募ボタン送信',
    '済'
  );

  setCellValueByHeader_(
    sheet,
    headers,
    targetRow,
    '応募ボタン送信日時',
    new Date()
  );

  setCellValueByHeader_(
    sheet,
    headers,
    targetRow,
    '応募ボタン送信方法',
    method || ''
  );
}