/**
 * 8_AdminNotify.gs
 * 管理者LINEグループ通知・対応ステータスFlex
 */

function notifyApplicationToAdmin_(
  applicationNo,
  state,
  userId,
  displayName
) {
  const notificationData =
    buildAdminApplicationNotificationData_(
      applicationNo,
      state,
      userId,
      displayName
    );

  const text =
    buildAdminApplicationNotificationText_(
      notificationData
    );

  notifyAdminGroupAllText_(
    text,
    'APPLICATION_ADMIN_NOTIFY_ERROR'
  );
}

function buildAdminApplicationNotificationData_(
  applicationNo,
  state,
  userId,
  displayName
) {
  const extraAnswers =
    parseJsonSafe_(state.extraAnswers || '{}');

  const media =
    pickFirstValue_([
      state.media,
      extraAnswers['募集媒体']
    ]);

  const mediaCategory =
    pickFirstValue_([
      extraAnswers['募集媒体大分類']
    ]);

  const mediaOther =
    pickFirstValue_([
      extraAnswers['その他媒体']
    ]);

  const mediaDisplay =
    mediaCategory === 'その他'
      ? (mediaOther || 'その他')
      : (media || '-');

  const applicantName =
    pickFirstValue_([
      state.name,
      extraAnswers['名前'],
      extraAnswers['お名前'],
      extraAnswers['氏名'],
      displayName,
      state.displayName
    ]);

  const age =
    pickFirstValue_([
      state.age,
      extraAnswers['年齢']
    ]);

  const store =
    pickFirstValue_([
      state.store,
      extraAnswers['希望店舗']
    ]);

  const workDays =
    pickFirstValue_([
      extraAnswers['勤務日数'],
      extraAnswers['勤務日数・シフト'],
      extraAnswers['週何日'],
      extraAnswers['勤務希望日数']
    ]);

  const remarks =
    pickFirstValue_([
      state.remarks,
      extraAnswers['備考欄'],
      extraAnswers['備考'],
      extraAnswers['その他'],
      extraAnswers['質問7'],
      extraAnswers['質問8']
    ]);

  const applicationCount =
    countCompletedApplicationsByUserId_(userId);

  return {
    applicationNo,
    mediaDisplay,
    applicantName,
    age,
    store,
    workDays,
    remarks,
    applicationCount,
    nowText:
      Utilities.formatDate(
        new Date(),
        'Asia/Tokyo',
        'yyyy/MM/dd HH:mm'
      )
  };
}

function buildAdminApplicationNotificationText_(data) {
  const historyLines =
    Number(data.applicationCount || 0) > 1
      ? [
          '',
          '🔁 過去に応募したことがある応募者です。',
          `応募回数：${data.applicationCount}回目`
        ]
      : [];

  return [
    '{all}',
    '',
    `【応募No.${data.applicationNo || '-'}】`,
    `🕒 ${data.nowText || '-'}`,
    '',
    '✨ 新規応募が入りました。',
    `🏬 ${getAdminNotifyStoreLine_(data.store)}`,
    '',
    '📋 応募内容',
    '',
    `📮 募集媒体：${data.mediaDisplay || '-'}`,
    `👤 お名前：${data.applicantName || '-'}`,
    `🎂 年齢：${data.age || '-'}`,
    `🏬 希望店舗：${data.store || '-'}`,
    `📅 勤務日数：${data.workDays || '-'}`,
    `📝 備考欄：${data.remarks || '-'}`,
    ...historyLines
  ].join('\n');
}

function notifyInquiryToAdmin_(
  inquiryNo,
  applicationNo,
  userId,
  displayName,
  text
) {
  const now =
    Utilities.formatDate(
      new Date(),
      'Asia/Tokyo',
      'yyyy/MM/dd HH:mm'
    );

  const messageText = [
    '{all}',
    '',
    `【問い合わせNo.${inquiryNo || '-'}】`,
    `🕒 ${now}`,
    '',
    '💬 追加問い合わせが入りました。',
    `関連応募No：${applicationNo || '-'}`,
    `LINE名：${displayName || '-'}`,
    '',
    '📝 内容',
    `${text || '-'}`
  ].join('\n');

  notifyAdminGroupAllText_(
    messageText,
    'INQUIRY_ADMIN_NOTIFY_ERROR'
  );
}

function notifyAdminGroupAllText_(
  text,
  errorType
) {
  const groupId =
    getConfigValue_(CONFIG_ADMIN_GROUP_ID);

  if (!groupId) {
    errorLog_(
      'ADMIN_GROUP_ID_EMPTY',
      '設定シートの ADMIN_GROUP_ID が空です。',
      '',
      ''
    );
    return;
  }

  pushMessages_(
    groupId,
    [createAllMentionTextV2Message_(text)],
    errorType || 'ADMIN_GROUP_NOTIFY_ERROR'
  );
}

function createAllMentionTextV2Message_(text) {
  return {
    type: 'textV2',
    text: text || '{all}',
    substitution: {
      all: {
        type: 'mention',
        mentionee: {
          type: 'all'
        }
      }
    }
  };
}

function getAdminNotifyStoreLine_(store) {
  const value =
    String(store || '').trim();

  return value
    ? `${value}です。`
    : '希望店舗未入力です。';
}

function pickFirstValue_(values) {
  for (let i = 0; i < values.length; i++) {
    if (values[i] instanceof Date) {
      return values[i];
    }

    const value =
      String(values[i] || '').trim();

    if (value) {
      return value;
    }
  }

  return '';
}

function createSupportStatusFlexFromSupportData_(supportData) {
  const supportNo =
    supportData.supportNo || '';

  const buttonStatus =
    supportData.buttonStatus || '';

  const title =
    getSupportStatusReplyTitle_(buttonStatus);

  return {
    type: 'flex',
    altText: title,
    contents: {
      type: 'bubble',
      body: {
        type: 'box',
        layout: 'vertical',
        spacing: 'md',
        contents: [
          {
            type: 'text',
            text: title,
            weight: 'bold',
            size: 'md',
            wrap: true
          },
          {
            type: 'text',
            text: `応募No：${supportNo || '-'}`,
            size: 'sm',
            wrap: true
          },
          {
            type: 'text',
            text: `現在の状態：${buttonStatus || '-'}`,
            size: 'sm',
            wrap: true
          }
        ]
      }
    }
  };
}

function getSupportStatusReplyTitle_(buttonStatus) {
  const status =
    String(buttonStatus || '').trim();

  if (status === SUPPORT_STATUS_DONE) {
    return '✅ 対応完了にしました。';
  }

  if (status === SUPPORT_STATUS_IN_PROGRESS) {
    return '🔄 対応中にしました。';
  }

  return '対応状況を更新しました。';
}

function getSupportRowDataByNo_(supportNo) {
  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_SUPPORT);

  if (!sheet || sheet.getLastRow() < 3) {
    return createFallbackSupportData_(supportNo);
  }

  const headers =
    getSheetHeaders_(sheet);

  const noCol =
    headers.indexOf('No') + 1;

  if (noCol <= 0) {
    return createFallbackSupportData_(supportNo);
  }

  const values =
    sheet
      .getRange(3, 1, sheet.getLastRow() - 2, sheet.getLastColumn())
      .getValues();

  const targetNo =
    String(supportNo || '').trim();

  for (let i = 0; i < values.length; i++) {
    const rowNo =
      String(values[i][noCol - 1] || '').trim();

    if (rowNo === targetNo) {
      return buildSupportRowData_(
        headers,
        values[i],
        supportNo
      );
    }
  }

  return createFallbackSupportData_(supportNo);
}

function buildSupportRowData_(
  headers,
  row,
  supportNo
) {
  const data = {
    supportNo,
    title: '対応',
    lines: [],
    buttonStatus: ''
  };

  headers.forEach((header, index) => {
    if (header) {
      data[header] = row[index];
    }
  });

  data.buttonStatus =
    data['ステータス'] || '';

  data.lines = [
    `No：${supportNo}`,
    `ステータス：${data['ステータス'] || '-'}`,
    `LINE表示名：${data['LINE表示名'] || '-'}`,
    `名前：${data['名前'] || '-'}`,
    `面接担当：${data['面接担当'] || '-'}`,
    `合否：${data['合否'] || '-'}`,
    `面接：${data['面接'] || '-'}`,
    `体入：${data['体入'] || '-'}`,
    `本入：${data['本入'] || '-'}`
  ];

  return data;
}

function createFallbackSupportData_(supportNo) {
  return {
    supportNo,
    title: '対応',
    lines: [`No：${supportNo}`],
    buttonStatus: ''
  };
}

function getSupportStatusButton_(
  supportNo,
  buttonStatus
) {
  const status =
    String(buttonStatus || '').trim();

  if (status === SUPPORT_STATUS_DONE) {
    return {
      type: 'text',
      text: '✅ 対応完了済み',
      size: 'sm',
      color: '#666666',
      wrap: true
    };
  }

  if (status === SUPPORT_STATUS_IN_PROGRESS) {
    return {
      type: 'button',
      style: 'primary',
      color: '#4A4A4A',
      height: 'sm',
      action: {
        type: 'postback',
        label: '✅ 対応完了にする',
        data: `action=complete_support&supportNo=${encodeURIComponent(supportNo)}`
      }
    };
  }

  return {
    type: 'button',
    style: 'primary',
    color: '#06C755',
    height: 'sm',
    action: {
      type: 'postback',
      label: '対応中にする',
      data: `action=start_support&supportNo=${encodeURIComponent(supportNo)}`
    }
  };
}

function getSupportStatusNoteText_(buttonStatus) {
  const status =
    String(buttonStatus || '').trim();

  if (status === SUPPORT_STATUS_DONE) {
    return 'この応募は対応完了です。';
  }

  if (status === SUPPORT_STATUS_IN_PROGRESS) {
    return '対応中です。完了したらボタンを押してください。';
  }

  return '対応を開始する場合はボタンを押してください。';
}