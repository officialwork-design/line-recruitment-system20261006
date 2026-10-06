/**
 * 4_LineApi.gs
 * LINE Messaging API 共通処理
 *
 * 方針：
 * - reply / push / multicast / broadcast を共通化
 * - replyToken切れ時は userId があれば push にフォールバック
 * - LINE APIエラーはエラーログへ保存
 * - Webhook落下防止用の safe log も維持
 */

/**
 * テキスト返信
 */
function replyText_(
  replyToken,
  text,
  errorType,
  fallbackUserId
) {
  replyMessages_(
    replyToken,
    [
      {
        type: 'text',
        text: String(text || '')
      }
    ],
    errorType || 'REPLY_TEXT_ERROR',
    fallbackUserId
  );
}

/**
 * 複数メッセージ返信
 */
function replyMessages_(
  replyToken,
  messages,
  errorType,
  fallbackUserId
) {
  const safeMessages =
    normalizeLineMessages_(messages);

  if (safeMessages.length === 0) {
    return;
  }

  if (!replyToken) {
    pushFallbackIfPossible_(
      fallbackUserId,
      safeMessages,
      `${errorType || 'LINE_REPLY_ERROR'}_PUSH_FALLBACK`
    );
    return;
  }

  try {
    postLineApi_(
      'https://api.line.me/v2/bot/message/reply',
      {
        replyToken,
        messages: safeMessages
      },
      errorType || 'LINE_REPLY_ERROR'
    );

  } catch (error) {
    if (
      isInvalidReplyTokenError_(error) &&
      fallbackUserId
    ) {
      pushFallbackIfPossible_(
        fallbackUserId,
        safeMessages,
        `${errorType || 'LINE_REPLY_ERROR'}_PUSH_FALLBACK`
      );

      return;
    }

    errorLog_(
      errorType || 'LINE_REPLY_ERROR',
      error.stack || error.message,
      fallbackUserId || '',
      ''
    );

    throw error;
  }
}

/**
 * replyToken切れ時のpushフォールバック
 */
function pushFallbackIfPossible_(
  userId,
  messages,
  errorType
) {
  if (!userId) return;

  pushMessages_(
    userId,
    messages,
    errorType || 'LINE_PUSH_FALLBACK_ERROR'
  );
}

/**
 * Push送信
 */
function pushMessages_(
  userId,
  messages,
  errorType
) {
  const safeMessages =
    normalizeLineMessages_(messages);

  if (!userId) {
    throw new Error('userId が空です。');
  }

  if (safeMessages.length === 0) {
    throw new Error('送信メッセージが空です。');
  }

  postLineApi_(
    'https://api.line.me/v2/bot/message/push',
    {
      to: userId,
      messages: safeMessages
    },
    errorType || 'LINE_PUSH_ERROR'
  );
}

/**
 * Multicast送信
 */
function multicastMessages_(
  userIds,
  messages,
  errorType
) {
  const safeUserIds =
    (userIds || [])
      .map(id => String(id || '').trim())
      .filter(id => id !== '');

  const safeMessages =
    normalizeLineMessages_(messages);

  if (safeUserIds.length === 0) {
    throw new Error('送信先 userIds が空です。');
  }

  if (safeMessages.length === 0) {
    throw new Error('送信メッセージが空です。');
  }

  postLineApi_(
    'https://api.line.me/v2/bot/message/multicast',
    {
      to: safeUserIds,
      messages: safeMessages
    },
    errorType || 'LINE_MULTICAST_ERROR'
  );
}

/**
 * Broadcast送信
 */
function broadcastMessages_(
  messages,
  errorType
) {
  const safeMessages =
    normalizeLineMessages_(messages);

  if (safeMessages.length === 0) {
    throw new Error('送信メッセージが空です。');
  }

  postLineApi_(
    'https://api.line.me/v2/bot/message/broadcast',
    {
      messages: safeMessages
    },
    errorType || 'LINE_BROADCAST_ERROR'
  );
}

/**
 * LINE API POST共通処理
 */
function postLineApi_(
  url,
  payload,
  errorType
) {
  const token =
    getLineChannelAccessToken_();

  const response =
    UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

  return handleLineApiResponse_(
    response,
    payload,
    errorType || 'LINE_API_ERROR'
  );
}

/**
 * LINE APIレスポンス処理
 */
function handleLineApiResponse_(
  response,
  payload,
  errorType
) {
  const statusCode =
    response.getResponseCode();

  const responseText =
    response.getContentText();

  if (statusCode >= 300) {
    errorLog_(
      errorType || 'LINE_API_ERROR',
      responseText,
      '',
      JSON.stringify(payload)
    );

    throw new Error(
      `${errorType || 'LINE_API_ERROR'} : ${responseText}`
    );
  }

  return responseText;
}

/**
 * LINEプロフィール取得
 */
function getLineProfile_(
  userId
) {
  if (!userId) return {};

  const token =
    getConfigValue_('LINE_CHANNEL_ACCESS_TOKEN');

  if (!isValidLineToken_(token)) {
    return {};
  }

  const url =
    `https://api.line.me/v2/bot/profile/${userId}`;

  try {
    const response =
      UrlFetchApp.fetch(url, {
        method: 'get',
        headers: {
          Authorization: `Bearer ${token}`
        },
        muteHttpExceptions: true
      });

    const statusCode =
      response.getResponseCode();

    const responseText =
      response.getContentText();

    if (statusCode !== 200) {
      errorLog_(
        'PROFILE_ERROR',
        responseText,
        userId,
        ''
      );

      return {};
    }

    return JSON.parse(responseText);

  } catch (error) {
    errorLog_(
      'PROFILE_ERROR',
      error.stack || error.message,
      userId,
      ''
    );

    return {};
  }
}

/**
 * LINEメッセージ画像等のコンテンツBlob取得
 */
function getMessageContentBlob_(
  messageId
) {
  if (!messageId) {
    throw new Error('messageId が空です。');
  }

  const token =
    getLineChannelAccessToken_();

  const url =
    `https://api-data.line.me/v2/bot/message/${messageId}/content`;

  const response =
    UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        Authorization: `Bearer ${token}`
      },
      muteHttpExceptions: true
    });

  const statusCode =
    response.getResponseCode();

  if (statusCode >= 300) {
    errorLog_(
      'GET_MESSAGE_CONTENT_ERROR',
      response.getContentText(),
      '',
      messageId
    );

    throw new Error(
      `画像取得失敗 : ${response.getContentText()}`
    );
  }

  return response.getBlob();
}

/**
 * LINEチャネルアクセストークン取得
 */
function getLineChannelAccessToken_() {
  const token =
    getConfigValue_('LINE_CHANNEL_ACCESS_TOKEN');

  if (!isValidLineToken_(token)) {
    errorLog_(
      'ERROR',
      'LINE_CHANNEL_ACCESS_TOKEN is empty',
      '',
      ''
    );

    throw new Error(
      '設定シートの LINE_CHANNEL_ACCESS_TOKEN を入力してください。'
    );
  }

  return token;
}

/**
 * LINEトークン妥当性チェック
 */
function isValidLineToken_(token) {
  const value =
    String(token || '').trim();

  return !!value &&
    value !== 'ここにLINEチャネルアクセストークンを入力' &&
    value !== 'ここにトークンを入力';
}

/**
 * LINEメッセージ配列正規化
 */
function normalizeLineMessages_(messages) {
  if (!messages) return [];

  const list =
    Array.isArray(messages)
      ? messages
      : [messages];

  return list.filter(message => {
    return !!(
      message &&
      typeof message === 'object' &&
      message.type
    );
  });
}

/**
 * replyToken切れ判定
 */
function isInvalidReplyTokenError_(error) {
  const errorMessage =
    String(
      error && error.message
        ? error.message
        : ''
    );

  return (
    errorMessage.indexOf('Invalid reply token') !== -1 ||
    errorMessage.indexOf('invalid reply token') !== -1
  );
}

/**
 * Webhook落下防止用エラーログ
 */
function errorLogSafeForWebhook_(
  type,
  content,
  userId,
  memo
) {
  try {
    errorLog_(
      type,
      content,
      userId,
      memo
    );
  } catch (error) {
    // webhook落下防止
  }
}