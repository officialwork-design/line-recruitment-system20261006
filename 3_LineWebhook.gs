/**
 * 3_LineWebhook.gs
 * LINE Webhook入口・イベント振り分け
 *
 * 方針：
 * - doPost は必ず OK を返してWebhook落下を防ぐ
 * - イベント単位で try-catch し、1件の失敗で全体を止めない
 * - 個人チャット、グループ、ルームを分岐
 * - 再応募は応募途中のみ有効
 *
 * 方針（V2改修・2026/10）：
 * - 応募完了後は userState:${userId} が削除されるため、
 *   以後のメッセージは「state無し」として扱われる
 *   （通常の未応募ユーザーと同様、質問回答等は何もせず無視される）。
 * - 「応募開始ワード」（handleStartText_）とfollowイベント
 *   （handleFollowEvent_：友だち追加・再追加）のどちらから入っても、
 *   応募管理シートを正とする hasCompletedApplication_(userId) で
 *   過去応募済みかを確認し、該当する場合は新規応募フローを
 *   開始させない。
 * - 通常の新規ユーザーのfollowイベントは、従来通り応募フローを
 *   自動開始する。
 */

/**
 * LINE Webhook入口
 */
function doPost(e) {
  try {
    const events =
      parseLineWebhookEvents_(e);

    events.forEach(event => {
      handleLineEventSafely_(event);
    });

  } catch (error) {
    errorLogSafeForWebhook_(
      'WEBHOOK_ERROR',
      error.stack || error.message,
      '',
      'doPost'
    );
  }

  return ContentService
    .createTextOutput('OK')
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Webhookイベント配列を取得
 */
function parseLineWebhookEvents_(e) {
  const bodyText =
    e && e.postData
      ? e.postData.contents
      : '';

  if (!bodyText) {
    return [];
  }

  const body =
    JSON.parse(bodyText);

  return body.events || [];
}

/**
 * イベント単位の安全実行
 */
function handleLineEventSafely_(event) {
  try {
    handleLineEvent_(event);
  } catch (error) {
    errorLogSafeForWebhook_(
      'LINE_EVENT_ERROR',
      error.stack || error.message,
      getEventUserId_(event),
      JSON.stringify(event || {})
    );
  }
}

/**
 * イベント振り分け
 */
function handleLineEvent_(event) {
  if (!event || !event.type) return;

  logSourceIdIfEnabled_(event);

  if (isGroupOrRoomEvent_(event)) {
    handleGroupOrRoomEvent_(event);
    return;
  }

  const userId =
    getEventUserId_(event);

  const replyToken =
    getEventReplyToken_(event);

  if (event.type === 'follow') {
    handleFollowEvent_(
      userId,
      replyToken
    );
    return;
  }

  if (event.type === 'postback') {
    handlePostbackEvent_(
      event,
      userId,
      replyToken
    );
    return;
  }

  if (event.type === 'message') {
    handleMessageEvent_(
      event,
      userId,
      replyToken
    );
  }
}

/**
 * グループ・ルームイベント判定
 */
function isGroupOrRoomEvent_(event) {
  const sourceType =
    event &&
    event.source &&
    event.source.type
      ? event.source.type
      : '';

  return (
    sourceType === 'group' ||
    sourceType === 'room'
  );
}

/**
 * グループ・ルームイベント処理
 *
 * 現状：
 * - postback のみ処理
 * - 通常メッセージはグループID確認ログ用途以外は無視
 */
function handleGroupOrRoomEvent_(event) {
  if (event.type !== 'postback') {
    return;
  }

  handlePostbackEvent_(
    event,
    getEventUserId_(event),
    getEventReplyToken_(event)
  );
}

/**
 * postback処理
 */
function handlePostbackEvent_(
  event,
  userId,
  replyToken
) {
  const params =
    parsePostbackParams_(event);

  const action =
    params.action || '';

  const supportNo =
    params.supportNo || '';

  if (!supportNo) return;

  if (action === 'start_support') {
    handleSupportStatusPostback_(
      supportNo,
      SUPPORT_STATUS_IN_PROGRESS,
      replyToken,
      userId
    );
    return;
  }

  if (action === 'complete_support') {
    handleSupportStatusPostback_(
      supportNo,
      SUPPORT_STATUS_DONE,
      replyToken,
      userId
    );
  }
}

/**
 * 対応ステータスpostback共通処理
 */
function handleSupportStatusPostback_(
  supportNo,
  status,
  replyToken,
  userId
) {
  const updated =
    updateSupportStatusByNo_(
      supportNo,
      status
    );

  if (!updated) return;

  const supportData =
    getSupportRowDataByNo_(
      supportNo
    );

  supportData.buttonStatus =
    status;

  replyMessages_(
    replyToken,
    [createSupportStatusFlexFromSupportData_(supportData)],
    status === SUPPORT_STATUS_DONE
      ? 'SUPPORT_DONE_REPLY_ERROR'
      : 'SUPPORT_IN_PROGRESS_REPLY_ERROR',
    userId
  );
}

/**
 * postback data パース
 */
function parsePostbackParams_(event) {
  const data =
    event &&
    event.postback &&
    event.postback.data
      ? event.postback.data
      : '';

  return parseQueryString_(
    data
  );
}

/**
 * クエリ文字列パース
 */
function parseQueryString_(query) {
  const result = {};

  String(query || '')
    .split('&')
    .forEach(part => {
      const pair =
        part.split('=');

      const key =
        decodeURIComponent(pair[0] || '')
          .trim();

      const value =
        decodeURIComponent(
          pair.slice(1).join('=') || ''
        ).trim();

      if (key) {
        result[key] = value;
      }
    });

  return result;
}

/**
 * LINEグループID確認ログ
 */
function logSourceIdIfEnabled_(event) {
  const enabled =
    PropertiesService
      .getScriptProperties()
      .getProperty('ENABLE_GROUP_ID_LOG');

  if (enabled !== 'TRUE') return;

  const sourceId =
    getSourceIdFromEvent_(event);

  if (!sourceId) return;

  errorLog_(
    'LINE_SOURCE_ID',
    sourceId,
    getEventUserId_(event),
    event.source && event.source.type
      ? event.source.type
      : ''
  );
}

/**
 * sourceId取得
 */
function getSourceIdFromEvent_(event) {
  if (!event || !event.source) return '';

  if (event.source.groupId) return event.source.groupId;
  if (event.source.roomId) return event.source.roomId;
  if (event.source.userId) return event.source.userId;

  return '';
}

/**
 * userId取得
 */
function getEventUserId_(event) {
  return event &&
    event.source &&
    event.source.userId
      ? event.source.userId
      : '';
}

/**
 * replyToken取得
 */
function getEventReplyToken_(event) {
  return event && event.replyToken
    ? event.replyToken
    : '';
}

/**
 * 友だち追加時
 *
 * - 連絡先へ登録
 * - 処理ログ
 * - 応募フローを自動開始
 *
 * 方針（V2改修・2026/10）：
 * - 過去応募済みユーザー（応募管理シートを正とする
 *   hasCompletedApplication_(userId) で判定）が再度followした場合は、
 *   handleStartText_ と同様に新しい応募フローを開始しない。
 * - 通常の新規ユーザーのfollowイベントは、従来通り応募フローを
 *   自動開始する。
 */
function handleFollowEvent_(
  userId,
  replyToken
) {
  const profile =
    getLineProfile_(userId);

  const displayName =
    profile.displayName || '';

  upsertContactOnFollow_(
    userId,
    displayName
  );

  processLog_(
    'FOLLOW',
    '友だち追加されました。',
    userId,
    displayName
  );

  if (hasCompletedApplication_(userId)) {
    replyText_(
      replyToken,
      '既にご応募いただいております。ご応募ありがとうございました。',
      'APPLICATION_ALREADY_COMPLETED_REPLY_ERROR',
      userId
    );

    return;
  }

  startApplicationFlow_(
    userId,
    displayName,
    replyToken
  );
}

/**
 * メッセージ受信時
 */
function handleMessageEvent_(
  event,
  userId,
  replyToken
) {
  const message =
    event.message;

  if (!message) return;

  const state =
    getUserState_(userId);

  const displayName =
    resolveDisplayName_(
      userId,
      state
    );

  updateContactOnMessage_(
    userId,
    displayName
  );

  if (isFinishedApplicationState_(state)) {
    return;
  }

  if (message.type === 'text') {
    handleTextMessage_(
      message,
      userId,
      displayName,
      replyToken,
      state
    );
    return;
  }

  if (message.type === 'image') {
    handleImageMessage_(
      message,
      userId,
      displayName,
      replyToken,
      state
    );
    return;
  }

  handleUnsupportedMessageDuringPhoto_(
    state,
    replyToken,
    userId
  );
}

/**
 * 表示名取得
 */
function resolveDisplayName_(
  userId,
  state
) {
  if (state && state.displayName) {
    return state.displayName;
  }

  const profile =
    getLineProfile_(userId);

  return profile.displayName || '';
}

/**
 * 応募完了・キャンセル状態か
 */
function isFinishedApplicationState_(state) {
  return !!(
    state &&
    (
      state.status === STATUS_DONE ||
      state.status === STATUS_CANCEL
    )
  );
}

/**
 * テキストメッセージ処理
 */
function handleTextMessage_(
  message,
  userId,
  displayName,
  replyToken,
  state
) {
  const text =
    String(message.text || '').trim();

  if (isRestartWord_(text)) {
    handleRestartText_(
      userId,
      displayName,
      replyToken,
      state
    );
    return;
  }

  if (isStartWord_(text)) {
    handleStartText_(
      userId,
      displayName,
      replyToken,
      state
    );
    return;
  }

  if (!state) {
    return;
  }

  routeTextAnswerByState_(
    userId,
    displayName,
    replyToken,
    state,
    text
  );
}

/**
 * 再応募ワード処理
 *
 * 再応募は応募途中のみ有効。
 * 完了後は handleMessageEvent_ 側で return 済み。
 */
function handleRestartText_(
  userId,
  displayName,
  replyToken,
  state
) {
  if (!isApplicationInProgress_(state)) {
    return;
  }

  const profile =
    getLineProfile_(userId);

  const profileName =
    profile.displayName ||
    displayName ||
    '';

  startApplicationFlow_(
    userId,
    profileName,
    replyToken
  );
}

/**
 * 応募開始ワード処理
 *
 * 方針（V2改修・2026/10）：
 * - userState削除後も「完了済みユーザーは新規応募を開始できない」
 *   という現行仕様を維持するため、userStateの有無だけで判定せず、
 *   応募管理シートを正とする hasCompletedApplication_(userId) で
 *   過去応募済みかどうかを確認する。
 * - 応募途中ユーザーのルーティング（isApplicationInProgress_判定）は変更しない。
 */
function handleStartText_(
  userId,
  displayName,
  replyToken,
  state
) {
  const profile =
    getLineProfile_(userId);

  const profileName =
    profile.displayName ||
    displayName ||
    '';

  if (isApplicationInProgress_(state)) {
    replyText_(
      replyToken,
      `⚠️ すでに応募途中です。\nこのまま続けて回答してください。\n\n最初からやり直す場合は「${RESTART_WORD}」と入力してください。`,
      'APPLICATION_ALREADY_STARTED_REPLY_ERROR',
      userId
    );

    return;
  }

  if (hasCompletedApplication_(userId)) {
    replyText_(
      replyToken,
      '既にご応募いただいております。ご応募ありがとうございました。',
      'APPLICATION_ALREADY_COMPLETED_REPLY_ERROR',
      userId
    );

    return;
  }

  startApplicationFlow_(
    userId,
    profileName,
    replyToken
  );
}

/**
 * state別テキスト回答ルーティング
 */
function routeTextAnswerByState_(
  userId,
  displayName,
  replyToken,
  state,
  text
) {
  if (state.status === STATUS_WAIT_MEDIA) {
    handleMediaAnswer_(
      userId,
      displayName,
      replyToken,
      state,
      text
    );
    return;
  }

  if (state.status === STATUS_WAIT_MEDIA_OTHER) {
    handleMediaOtherAnswer_(
      userId,
      displayName,
      replyToken,
      state,
      text
    );
    return;
  }

  if (state.status === STATUS_WAITING) {
    handleTextAnswer_(
      userId,
      displayName,
      replyToken,
      state,
      text
    );
    return;
  }

  if (state.status === STATUS_WAIT_REMARKS) {
    handleRemarksAnswer_(
      userId,
      displayName,
      replyToken,
      state,
      text
    );
    return;
  }

  if (state.status === STATUS_WAIT_FACE_PHOTO) {
    replyText_(
      replyToken,
      getPhotoImageErrorMessage_(
        getFacePhotoMessage_(state)
      ),
      'FACE_PHOTO_TEXT_ERROR',
      userId
    );
    return;
  }

  if (state.status === STATUS_WAIT_FULL_BODY_PHOTO) {
    replyText_(
      replyToken,
      getPhotoImageErrorMessage_(
        getFullBodyPhotoMessage_(state)
      ),
      'FULL_BODY_PHOTO_TEXT_ERROR',
      userId
    );
  }
}

/**
 * 写真待ち中に未対応メッセージが来た場合
 */
function handleUnsupportedMessageDuringPhoto_(
  state,
  replyToken,
  userId
) {
  if (
    state &&
    state.status === STATUS_WAIT_FACE_PHOTO
  ) {
    replyText_(
      replyToken,
      getPhotoImageErrorMessage_(
        getFacePhotoMessage_(state)
      ),
      'FACE_PHOTO_UNSUPPORTED_MESSAGE_ERROR',
      userId
    );

    return;
  }

  if (
    state &&
    state.status === STATUS_WAIT_FULL_BODY_PHOTO
  ) {
    replyText_(
      replyToken,
      getPhotoImageErrorMessage_(
        getFullBodyPhotoMessage_(state)
      ),
      'FULL_BODY_PHOTO_UNSUPPORTED_MESSAGE_ERROR',
      userId
    );
  }
}