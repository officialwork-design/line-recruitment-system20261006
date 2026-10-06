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
 * 全身写真依頼メッセージ
 *
 * 【注】現在の1枚写真フローでは新規に呼ばれることはないが、
 * 3_LineWebhook.gs の routeTextAnswerByState_ /
 * handleUnsupportedMessageDuringPhoto_ が、過去（2枚写真フロー時代）に
 * STATUS_WAIT_FULL_BODY_PHOTO のまま残っている可能性のある古い
 * userState を防御的に処理するために参照している。そのため削除しない
 * （後方互換のためのコード）。
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
 * 【注】getFullBodyPhotoMessage_（後方互換用）から参照される共通ヘルパー。
 * 新しい写真提出メッセージは getApplicationPhotoMessage_ を使用する。
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
 * 方針（1枚写真フロー）：
 * - 写真は1枚のみ。受信した1枚を facePhotoStatus に記録する。
 * - fullBodyPhotoStatus フィールドは2026/10の列削除に伴い廃止した
 *   （以前は「常に未設定のまま残す」互換フィールドだったが、今回の
 *   スキーマ簡素化で完全に削除した）。
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
