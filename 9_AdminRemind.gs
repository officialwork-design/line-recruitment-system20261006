/**
 * 9_AdminRemind.gs
 * 応募No指定・対応管理行選択から管理者グループへ再メンション
 */

/**
 * 応募Noを手入力して再メンション
 */
function resendApplicationMentionByPrompt() {
  const ui = SpreadsheetApp.getUi();

  const noPrompt = ui.prompt(
    '応募Noを指定して再メンション',
    '再メンションしたい応募Noを入力してください。\n例：15',
    ui.ButtonSet.OK_CANCEL
  );

  if (noPrompt.getSelectedButton() !== ui.Button.OK) return;

  const applicationNo =
    String(noPrompt.getResponseText() || '').trim();

  if (!applicationNo) {
    ui.alert('応募Noが空です。');
    return;
  }

  const messagePrompt = ui.prompt(
    '追加文言',
    'メンションの上部に追加する文言を入力してください。\n空欄の場合は「進捗状況いかがですか？」になります。',
    ui.ButtonSet.OK_CANCEL
  );

  if (messagePrompt.getSelectedButton() !== ui.Button.OK) return;

  const customMessage =
    String(messagePrompt.getResponseText() || '').trim() ||
    '進捗状況いかがですか？';

  const applicationData =
    getApplicationRowDataByNo_(applicationNo);

  if (!applicationData) {
    ui.alert(`応募No.${applicationNo} が応募管理に見つかりません。`);
    return;
  }

  const text =
    buildApplicationReMentionText_(
      applicationData,
      customMessage
    );

  notifyAdminGroupAllText_(
    text,
    'APPLICATION_RE_MENTION_ERROR'
  );

  ui.alert(`応募No.${applicationNo} を再メンションしました。`);
}

/**
 * 対応管理で選択中の行から再メンション
 */
function resendApplicationMentionFromSelectedSupportRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  if (sheet.getName() !== SHEET_SUPPORT) {
    ui.alert('対応管理シートで再メンションしたい行を選択してください。');
    return;
  }

  const selectedRow =
    sheet.getActiveRange().getRow();

  if (selectedRow < 3) {
    ui.alert('再メンションしたい応募行を選択してください。');
    return;
  }

  const headers =
    sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(h => String(h || '').trim());

  const noCol =
    headers.indexOf('No') + 1;

  if (noCol <= 0) {
    ui.alert('対応管理シートに No 列がありません。');
    return;
  }

  const applicationNo =
    String(
      sheet
        .getRange(selectedRow, noCol)
        .getValue() || ''
    ).trim();

  if (!applicationNo) {
    ui.alert('選択行に応募Noがありません。');
    return;
  }

  if (applicationNo.indexOf('問い合わせ-') === 0) {
    ui.alert('問い合わせ行は再メンション対象外です。');
    return;
  }

  const messagePrompt = ui.prompt(
    '追加文言',
    'メンションの上部に追加する文言を入力してください。\n空欄の場合は「進捗状況いかがですか？」になります。',
    ui.ButtonSet.OK_CANCEL
  );

  if (messagePrompt.getSelectedButton() !== ui.Button.OK) return;

  const customMessage =
    String(messagePrompt.getResponseText() || '').trim() ||
    '進捗状況いかがですか？';

  const applicationData =
    getApplicationRowDataByNo_(applicationNo);

  if (!applicationData) {
    ui.alert(`応募No.${applicationNo} が応募管理に見つかりません。`);
    return;
  }

  const text =
    buildApplicationReMentionText_(
      applicationData,
      customMessage
    );

  notifyAdminGroupAllText_(
    text,
    'APPLICATION_RE_MENTION_SELECTED_ERROR'
  );

  ui.alert(`応募No.${applicationNo} を再メンションしました。`);
}

/**
 * 再メンション本文を作成
 */
function buildApplicationReMentionText_(
  applicationData,
  customMessage
) {
  const applicationNo =
    pickApplicationValue_(
      applicationData,
      ['No']
    );

  const receivedAt =
    pickApplicationValue_(
      applicationData,
      ['受信日時', '更新日']
    );

  const formattedDate =
    formatApplicationDateForMention_(receivedAt);

  const media =
    pickApplicationValue_(
      applicationData,
      ['募集媒体']
    );

  const mediaCategory =
    pickApplicationValue_(
      applicationData,
      ['募集媒体大分類']
    );

  const mediaOther =
    pickApplicationValue_(
      applicationData,
      ['その他媒体']
    );

  const mediaDisplay =
    mediaCategory === 'その他'
      ? (mediaOther || 'その他')
      : (media || '-');

  const applicantName =
    pickApplicationValue_(
      applicationData,
      ['名前', 'お名前', 'LINE表示名']
    );

  const age =
    pickApplicationValue_(
      applicationData,
      ['年齢']
    );

  const store =
    pickApplicationValue_(
      applicationData,
      ['希望店舗']
    );

  const workDays =
    pickApplicationValue_(
      applicationData,
      ['勤務日数']
    );

  const remarks =
    pickApplicationValue_(
      applicationData,
      ['備考欄']
    );

  return [
    '{all}',
    '',
    customMessage || '進捗状況いかがですか？',
    '',
    `【応募No.${applicationNo || '-'}】`,
    `🕒 ${formattedDate || '-'}`,
    '',
    '✨ 新規応募が入りました。',
    `🏬 ${getAdminNotifyStoreLine_(store)}`,
    '',
    '📋 応募内容',
    '',
    `📮 募集媒体：${mediaDisplay}`,
    `👤 お名前：${applicantName || '-'}`,
    `🎂 年齢：${age || '-'}`,
    `🏬 希望店舗：${store || '-'}`,
    `📅 勤務日数：${workDays || '-'}`,
    `📝 備考欄：${remarks || '-'}`
  ].join('\n');
}

/**
 * 応募管理データから最初に値が入っているものを取得
 */
function pickApplicationValue_(
  data,
  keys
) {
  for (let i = 0; i < keys.length; i++) {
    const value =
      data[keys[i]];

    if (value instanceof Date) {
      return value;
    }

    const text =
      String(value || '').trim();

    if (text) {
      return text;
    }
  }

  return '';
}

/**
 * 再メンション用の日付整形
 */
function formatApplicationDateForMention_(value) {
  if (!value) return '';

  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      'Asia/Tokyo',
      'yyyy/MM/dd HH:mm'
    );
  }

  const date =
    new Date(value);

  if (!isNaN(date.getTime())) {
    return Utilities.formatDate(
      date,
      'Asia/Tokyo',
      'yyyy/MM/dd HH:mm'
    );
  }

  return String(value || '');
}