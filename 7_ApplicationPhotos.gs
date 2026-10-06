/**
 * 7_ApplicationPhotos.gs
 * 応募写真受付
 *
 * 方針（V2改修・2026/10：応募フロー簡略化）：
 * - 事前確認事項の質問（6_ApplicationQuestions.gs） → 写真1枚提出 → 完了
 *   の順で進行する。顔写真・全身写真の2枚提出は廃止した。
 * - 現時点では画像ファイル本体は保存せず、受信状態だけ管理する
 *   （Drive保存等は今回追加しない）。
 * - 写真以外が送られた場合は再案内する。
 * - 画像を1枚受信した時点で completeApplication_ を呼び、
 *   2枚目は要求しない。
 */

/**
 * 顔写真依頼メッセージ
 *
 * 【注】応募フロー簡略化（V2改修・2026/10）により未使用。
 * 写真提出メッセージは getApplicationPhotoMessage_ を使用する。
 * 既存データ互換性のため関数定義は削除せず残す。
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
 *
 * 【注】応募フロー簡略化（V2改修・2026/10）により未使用。
 * 既存データ互換性のため関数定義は削除せず残す。
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
 *
 * 【注】応募フロー簡略化（V2改修・2026/10）により、
 * getFacePhotoMessage_ / getFullBodyPhotoMessage_ 経由でのみ
 * 参照される（どちらも未使用）。新しい写真提出メッセージは
 * getApplicationPhotoMessage_ を使用する。
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
 * 写真提出メッセージ（応募フロー簡略化・2026/10追加：写真は1枚のみ）
 *
 * 写真提出は応募フロー上の最後の入力ステップのため、
 * 常に「質問 total / total」として表示する。
 */
function getApplicationPhotoMessage_(state) {
  const total =
    getTotalApplicationSteps_(state);

  const message = [
    '最後に、あなたのお写真を1枚送ってください📷',
    '顔がわかりやすい写真をお願いします！'
  ].join('\n');

  return `📸 質問 ${total} / ${total}\n\n${message}`;
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
 *
 * 方針（V2改修・2026/10：応募フロー簡略化）：
 * - 写真は1枚のみ。画像を1枚受信した時点で completeApplication_ を呼び、
 *   2枚目は要求しない（proceedToFullBodyPhoto_ 等は呼ばない）。
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

  completeApplication_(
    userId,
    displayName,
    replyToken,
    updatedState
  );
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
 *
 * 方針（V2改修・2026/10：応募フロー簡略化）：
 * - 写真は1枚のみ。受信した1枚を facePhotoStatus に記録する。
 * - fullBodyPhotoStatus は今回のフローでは使用しない
 *   （既存データ互換性のためフィールド自体は残す。常に未設定のまま）。
 * - ステータス遷移は行わない（受信後はこのまま completeApplication_
 *   を呼ぶため、呼び出し元の handleImageMessage_ 側で完結する）。
 */
function buildPhotoReceivedState_(state) {
  return {
    ...state,
    photoCount: Number(state.photoCount || 0) + 1,
    facePhotoStatus: PHOTO_RECEIVED_TEXT,
    updatedAt: new Date().toISOString()
  };
}

/**
 * 顔写真受信後、全身写真へ進む
 *
 * 【注】応募フロー簡略化（V2改修・2026/10）により、
 * handleImageMessage_ から呼ばれなくなり未使用。
 * 既存データ互換性のため関数定義は削除せず残す。
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
 *
 * 【注】応募フロー簡略化（V2改修・2026/10）により未使用。
 * 既存データ互換性のため関数定義は削除せず残す。
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
 *
 * 【注】応募フロー簡略化（V2改修・2026/10）により未使用。
 * 既存データ互換性のため関数定義は削除せず残す。
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
