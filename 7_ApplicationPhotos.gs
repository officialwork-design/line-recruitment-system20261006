/**
 * 7_ApplicationPhotos.gs
 * 応募写真受付
 *
 * 方針：
 * - 顔写真 → 全身写真 → 備考欄 → 完了 の順で進行
 * - 現時点では画像ファイル本体は保存せず、受信状態だけ管理する
 * - 写真以外が送られた場合は再案内する
 */

/**
 * 顔写真依頼メッセージ
 */
function getFacePhotoMessage_(state) {
  return buildPhotoRequestMessage_(
    state,
    1,
    FACE_PHOTO_MESSAGE
  );
}

/**
 * 全身写真依頼メッセージ
 */
function getFullBodyPhotoMessage_(state) {
  return buildPhotoRequestMessage_(
    state,
    2,
    FULL_BODY_PHOTO_MESSAGE
  );
}

/**
 * 写真依頼メッセージ共通作成
 */
function buildPhotoRequestMessage_(
  state,
  photoStepOffset,
  message
) {
  const total =
    getTotalApplicationSteps_(state);

  const normalQuestions =
    getNormalQuestionsBeforePhoto_();

  const displayNumber =
    getQuestionOffset_(state) +
    normalQuestions.length +
    Number(photoStepOffset || 1);

  return `📸 質問 ${displayNumber} / ${total}\n\n${message}`;
}

/**
 * 写真以外が送られた場合のエラーメッセージ
 */
function getPhotoImageErrorMessage_(photoMessage) {
  return [
    '⚠️ 写真画像を送ってください。',
    PHOTO_FILE_NOTE,
    '',
    photoMessage
  ].join('\n');
}

/**
 * 画像メッセージ受信処理
 */
function handleImageMessage_(
  message,
  userId,
  displayName,
  replyToken,
  state
) {
  if (!state) {
    replyText_(
      replyToken,
      `応募をご希望の場合は「${DEFAULT_START_WORD}」と送信してください。`,
      'IMAGE_NO_STATE_REPLY_ERROR',
      userId
    );

    return;
  }

  if (!isWaitingPhotoStatus_(state.status)) {
    replyText_(
      replyToken,
      '⚠️ 現在、写真受付の状態になっていません。',
      'IMAGE_INVALID_STATUS_REPLY_ERROR',
      userId
    );

    return;
  }

  const updatedState =
    buildPhotoReceivedState_(
      state
    );

  if (state.status === STATUS_WAIT_FACE_PHOTO) {
    proceedToFullBodyPhoto_(
      userId,
      replyToken,
      updatedState
    );

    return;
  }

  if (state.status === STATUS_WAIT_FULL_BODY_PHOTO) {
    proceedAfterFullBodyPhoto_(
      userId,
      displayName,
      replyToken,
      updatedState
    );
  }
}

/**
 * 写真待ちステータスか判定
 */
function isWaitingPhotoStatus_(status) {
  return [
    STATUS_WAIT_FACE_PHOTO,
    STATUS_WAIT_FULL_BODY_PHOTO
  ].includes(
    String(status || '').trim()
  );
}

/**
 * 写真受信後の状態を作成
 */
function buildPhotoReceivedState_(state) {
  const newPhotoCount =
    Number(state.photoCount || 0) + 1;

  const updatedState = {
    ...state,
    photoCount: newPhotoCount,
    updatedAt: new Date().toISOString()
  };

  if (state.status === STATUS_WAIT_FACE_PHOTO) {
    updatedState.status =
      STATUS_WAIT_FULL_BODY_PHOTO;

    updatedState.facePhotoStatus =
      PHOTO_RECEIVED_TEXT;
  }

  if (state.status === STATUS_WAIT_FULL_BODY_PHOTO) {
    updatedState.fullBodyPhotoStatus =
      PHOTO_RECEIVED_TEXT;
  }

  return updatedState;
}

/**
 * 顔写真受信後、全身写真へ進む
 */
function proceedToFullBodyPhoto_(
  userId,
  replyToken,
  updatedState
) {
  upsertUserManagement_(
    updatedState
  );

  replyText_(
    replyToken,
    getFullBodyPhotoMessage_(updatedState),
    'FULL_BODY_PHOTO_REQUEST_ERROR',
    userId
  );
}

/**
 * 全身写真受信後、備考欄または完了へ進む
 */
function proceedAfterFullBodyPhoto_(
  userId,
  displayName,
  replyToken,
  updatedState
) {
  const remarksQuestion =
    getRemarksQuestion_();

  if (remarksQuestion) {
    proceedToRemarksQuestion_(
      userId,
      replyToken,
      updatedState,
      remarksQuestion
    );

    return;
  }

  completeApplication_(
    userId,
    displayName,
    replyToken,
    updatedState
  );
}

/**
 * 備考欄へ進む
 */
function proceedToRemarksQuestion_(
  userId,
  replyToken,
  updatedState,
  remarksQuestion
) {
  updatedState.status =
    STATUS_WAIT_REMARKS;

  updatedState.currentQuestionNo =
    remarksQuestion.order;

  updatedState.updatedAt =
    new Date().toISOString();

  upsertUserManagement_(
    updatedState
  );

  replyText_(
    replyToken,
    `${getRemarksProgressText_(updatedState)}${cleanQuestionText_(remarksQuestion.questionText)}`,
    'REMARKS_REQUEST_ERROR',
    userId
  );
}
