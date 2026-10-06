/**
 * 5_ApplicationFlow.gs
 * 応募フロー本体
 *
 * 方針：
 * - 応募開始 / 再応募判定
 * - 募集媒体 → 通常質問 → 顔写真 → 全身写真 → 備考欄 → 完了
 * - 応募完了時に年齢判定、自動返信、応募保存、対応履歴保存、状態更新
 */

/**
 * 応募開始ワード判定
 */
function isStartWord_(text) {
  const value =
    normalizeText_(text);

  const startWord =
    normalizeText_(
      getConfigValue_('START_WORD') ||
      DEFAULT_START_WORD
    );

  return value === startWord;
}

/**
 * 再応募ワード判定
 */
function isRestartWord_(text) {
  return normalizeText_(text) === normalizeText_(RESTART_WORD);
}

/**
 * 応募途中か判定
 */
function isApplicationInProgress_(state) {
  if (!state) return false;

  const status =
    String(state.status || '').trim();

  return [
    STATUS_WAIT_MEDIA,
    STATUS_WAIT_MEDIA_OTHER,
    STATUS_WAITING,
    STATUS_WAIT_FACE_PHOTO,
    STATUS_WAIT_FULL_BODY_PHOTO,
    STATUS_WAIT_REMARKS
  ].includes(status);
}

/**
 * テキスト正規化
 */
function normalizeText_(text) {
  return String(text || '')
    .replace(/\s/g, '')
    .replace(/　/g, '')
    .trim();
}

/**
 * 基本質問数
 *
 * 募集媒体 1
 * 通常質問 n
 * 顔写真 1
 * 全身写真 1
 * 備考欄 0 or 1
 */
function getBaseApplicationSteps_() {
  const normalQuestions =
    getNormalQuestionsBeforePhoto_();

  const remarksQuestion =
    getRemarksQuestion_();

  return 1 + normalQuestions.length + 2 + (remarksQuestion ? 1 : 0);
}

/**
 * 最大質問数
 *
 * 「その他媒体」分岐がある場合は +1
 */
function getMaxApplicationSteps_() {
  return getBaseApplicationSteps_() + 1;
}

/**
 * stateに保存された総質問数を優先
 */
function getTotalApplicationSteps_(state) {
  if (state && Number(state.totalSteps || 0) > 0) {
    return Number(state.totalSteps);
  }

  return getBaseApplicationSteps_();
}

/**
 * 写真前の通常質問
 */
function getNormalQuestionsBeforePhoto_() {
  return getActiveQuestions_()
    .filter(q => q.itemKey !== 'remarks')
    .sort((a, b) => Number(a.order) - Number(b.order));
}

/**
 * 備考欄質問
 */
function getRemarksQuestion_() {
  return getActiveQuestions_()
    .find(q => q.itemKey === 'remarks') || null;
}

/**
 * 応募フロー開始
 */
function startApplicationFlow_(
  userId,
  displayName,
  replyToken
) {
  updateContactApplicationStatus_(
    userId,
    '応募中',
    ''
  );

  const initialState =
    createInitialApplicationState_(
      userId,
      displayName
    );

  upsertUserManagement_(
    initialState
  );

  sendMediaQuestion_(
    replyToken,
    userId
  );
}

/**
 * 初期応募状態作成
 */
function createInitialApplicationState_(
  userId,
  displayName
) {
  return {
    userId,
    displayName,
    currentQuestionNo: '',
    status: STATUS_WAIT_MEDIA,
    name: '',
    store: '',
    media: '',
    mediaOtherFlow: 'FALSE',
    totalSteps: getMaxApplicationSteps_(),
    photoCount: 0,
    facePhotoStatus: '',
    fullBodyPhotoStatus: '',
    extraAnswers: '{}',
    applicationMessage: '',
    age: '',
    remarks: '',
    updatedAt: new Date().toISOString()
  };
}

/**
 * 最初の通常質問、または顔写真へ進む
 */
function goToFirstNormalQuestionOrPhoto_(
  userId,
  displayName,
  replyToken,
  state
) {
  const questions =
    getNormalQuestionsBeforePhoto_();

  if (questions.length === 0) {
    moveToFacePhoto_(
      userId,
      replyToken,
      state
    );

    return;
  }

  const firstQuestion =
    questions[0];

  const updatedState = {
    ...state,
    currentQuestionNo: firstQuestion.order,
    status: STATUS_WAITING,
    updatedAt: new Date().toISOString()
  };

  upsertUserManagement_(
    updatedState
  );

  sendQuestion_(
    replyToken,
    firstQuestion,
    userId,
    updatedState
  );
}

/**
 * 顔写真待ちへ進む
 */
function moveToFacePhoto_(
  userId,
  replyToken,
  state
) {
  const updatedState = {
    ...state,
    currentQuestionNo: '',
    status: STATUS_WAIT_FACE_PHOTO,
    updatedAt: new Date().toISOString()
  };

  upsertUserManagement_(
    updatedState
  );

  replyText_(
    replyToken,
    getFacePhotoMessage_(updatedState),
    'FACE_PHOTO_REQUEST_ERROR',
    userId
  );
}

/**
 * 通常質問の回答処理
 */
function handleTextAnswer_(
  userId,
  displayName,
  replyToken,
  state,
  text
) {
  const questions =
    getNormalQuestionsBeforePhoto_();

  const currentQuestion =
    findQuestionByOrder_(
      questions,
      state.currentQuestionNo
    );

  if (!currentQuestion) {
    handleQuestionStateError_(
      userId,
      displayName,
      replyToken,
      state
    );

    return;
  }

  if (
    !validateQuestionAnswer_(
      replyToken,
      currentQuestion,
      text,
      userId,
      state
    )
  ) {
    return;
  }

  const updatedState =
    applyAnswerToState_(
      state,
      currentQuestion,
      text,
      displayName
    );

  const nextQuestion =
    getNextQuestion_(
      questions,
      currentQuestion.order
    );

  if (!nextQuestion) {
    moveToFacePhoto_(
      userId,
      replyToken,
      updatedState
    );

    return;
  }

  moveToNextQuestion_(
    userId,
    replyToken,
    updatedState,
    nextQuestion
  );
}

/**
 * 質問状態エラー
 */
function handleQuestionStateError_(
  userId,
  displayName,
  replyToken,
  state
) {
  replyText_(
    replyToken,
    '⚠️ 応募フローの状態確認でエラーが発生しました。\n恐れ入りますが、もう一度「再応募」と送ってください。',
    'QUESTION_STATE_ERROR',
    userId
  );

  errorLog_(
    'QUESTION_STATE_ERROR',
    `currentQuestionNo=${state.currentQuestionNo}`,
    userId,
    displayName
  );
}

/**
 * 回答バリデーション
 */
function validateQuestionAnswer_(
  replyToken,
  question,
  text,
  userId,
  state
) {
  if (question.answerType === '選択ボタン') {
    const choices =
      getChoicesByItemKey_(question.itemKey);

    const validChoices =
      choices.map(choice => choice.sendText);

    if (!validChoices.includes(text)) {
      sendInvalidChoiceQuestion_(
        replyToken,
        question,
        choices,
        userId,
        state
      );

      return false;
    }
  }

  if (question.answerType === '文字入力') {
    if (!isValidTextAnswer_(text, question.required)) {
      replyText_(
        replyToken,
        [
          '⚠️ 必須項目です。',
          '入力をお願いします。',
          '',
          `${getQuestionProgressText_(question, state)}${cleanQuestionText_(question.questionText)}`
        ].join('\n'),
        'EMPTY_TEXT_ANSWER_ERROR',
        userId
      );

      return false;
    }
  }

  return true;
}

/**
 * 次の通常質問へ進む
 */
function moveToNextQuestion_(
  userId,
  replyToken,
  state,
  nextQuestion
) {
  const updatedState = {
    ...state,
    currentQuestionNo: nextQuestion.order,
    status: STATUS_WAITING,
    updatedAt: new Date().toISOString()
  };

  upsertUserManagement_(
    updatedState
  );

  sendQuestion_(
    replyToken,
    nextQuestion,
    userId,
    updatedState
  );
}

/**
 * 備考欄回答処理
 */
function handleRemarksAnswer_(
  userId,
  displayName,
  replyToken,
  state,
  text
) {
  const remarksQuestion =
    getRemarksQuestion_();

  if (!remarksQuestion) {
    completeApplication_(
      userId,
      displayName,
      replyToken,
      state
    );

    return;
  }

  if (!isValidTextAnswer_(text, remarksQuestion.required)) {
    replyText_(
      replyToken,
      [
        '⚠️ 必須項目です。',
        '入力をお願いします。',
        '',
        `${getRemarksProgressText_(state)}${cleanQuestionText_(remarksQuestion.questionText)}`
      ].join('\n'),
      'EMPTY_REMARKS_ANSWER_ERROR',
      userId
    );

    return;
  }

  const updatedState =
    applyAnswerToState_(
      state,
      remarksQuestion,
      text,
      displayName
    );

  completeApplication_(
    userId,
    displayName,
    replyToken,
    updatedState
  );
}

/**
 * 文字入力の必須チェック
 */
function isValidTextAnswer_(
  text,
  required
) {
  const value =
    String(text || '').trim();

  if (required && value === '') {
    return false;
  }

  return true;
}

/**
 * 回答内容を応募状態に反映
 */
function applyAnswerToState_(
  state,
  question,
  answer,
  displayName
) {
  const extraAnswers =
    parseJsonSafe_(state.extraAnswers || '{}');

  const saveName =
    question.saveColumnName ||
    question.itemName ||
    question.itemKey;

  extraAnswers[saveName] =
    answer;

  const normalizedApplicant =
    normalizeApplicantFieldsFromAnswer_(
      state,
      question,
      answer,
      saveName,
      extraAnswers
    );

  const mergedApplicationMessage =
    buildApplicationMessage_({
      currentMessage: state.applicationMessage || '',
      label: question.itemName || saveName || question.itemKey,
      answer
    });

  return {
    ...state,
    displayName: displayName || state.displayName || '',
    name: normalizedApplicant.name,
    store: normalizedApplicant.store,
    media: state.media || extraAnswers['募集媒体'] || '',
    photoCount: Number(state.photoCount || 0),
    facePhotoStatus: state.facePhotoStatus || '',
    fullBodyPhotoStatus: state.fullBodyPhotoStatus || '',
    extraAnswers: JSON.stringify(extraAnswers),
    applicationMessage: mergedApplicationMessage,
    age: normalizedApplicant.age,
    remarks: normalizedApplicant.remarks,
    updatedAt: new Date().toISOString()
  };
}

/**
 * 回答内容から主要項目を正規化
 */
function normalizeApplicantFieldsFromAnswer_(
  state,
  question,
  answer,
  saveName,
  extraAnswers
) {
  let name =
    state.name || '';

  let store =
    state.store || '';

  let age =
    state.age || '';

  let remarks =
    state.remarks || '';

  if (
    question.itemKey === 'name' ||
    saveName === 'お名前' ||
    saveName === '名前' ||
    saveName === '氏名'
  ) {
    name = answer;
    extraAnswers['名前'] = answer;
    extraAnswers['お名前'] = answer;
  }

  if (
    question.itemKey === 'store' ||
    saveName === '希望店舗'
  ) {
    store = answer;
  }

  if (
    question.itemKey === 'age' ||
    saveName === '年齢'
  ) {
    const parsedApplicant =
      parseApplicantNameAndAge_(answer);

    if (parsedApplicant.name) {
      name = parsedApplicant.name;
      extraAnswers['名前'] = parsedApplicant.name;
      extraAnswers['お名前'] = parsedApplicant.name;
    }

    if (parsedApplicant.age) {
      age = parsedApplicant.age;
      extraAnswers['年齢'] = parsedApplicant.age;
    } else {
      age = answer;
      extraAnswers['年齢'] = answer;
    }
  }

  if (
    question.itemKey === 'remarks' ||
    saveName === '備考欄'
  ) {
    remarks = answer;
  }

  return {
    name,
    store,
    age,
    remarks
  };
}

/**
 * 応募メッセージを追記
 */
function buildApplicationMessage_(
  params
) {
  const itemLabel =
    params.label ||
    (
      params.question
        ? (
            params.question.itemName ||
            params.question.saveColumnName ||
            params.question.itemKey
          )
        : ''
    );

  const line =
    `${itemLabel}：${params.answer}`;

  if (!params.currentMessage) return line;

  return `${params.currentMessage}\n${line}`;
}

/**
 * 応募完了処理
 *
 * 方針（V2改修・2026/10）：
 * - 応募データの保存を最優先する。LINEへの完了replyは保存の後に行う。
 * - reply送信（replyMessages_）に失敗しても、既に保存済みの
 *   応募データは巻き戻らない（保存とreplyを互いに独立させる）。
 * - reply失敗はログに残すが、例外を外へ伝播させない。
 */
function completeApplication_(
  userId,
  displayName,
  replyToken,
  state
) {
  const normalizedState =
    normalizeApplicationStateBeforeComplete_(
      state
    );

  let saveSucceeded = false;

  try {
    saveCompletedApplication_(
      userId,
      displayName,
      normalizedState
    );

    saveSucceeded = true;

  } catch (error) {
    processLog_(
      'APPLICATION_SAVE_ERROR',
      error.stack || error.message,
      userId,
      displayName
    );

    errorLog_(
      'APPLICATION_SAVE_ERROR',
      error.stack || error.message,
      userId,
      displayName
    );
  }

  try {
    replyMessages_(
      replyToken,
      buildCompleteReplyMessages_(normalizedState),
      'COMPLETE_REPLY_ERROR',
      userId
    );

  } catch (error) {
    // 保存結果（成功・失敗いずれの場合も）には影響させない。
    // replyMessages_ 内で既に errorLog_ 済みのため、ここでは
    // 保存結果とあわせた状況確認用のログのみ残して握りつぶす。
    processLog_(
      'COMPLETE_REPLY_FAILED_AFTER_SAVE',
      `reply失敗（保存結果: ${saveSucceeded ? '成功' : '失敗'}）: ${error.stack || error.message}`,
      userId,
      displayName
    );
  }
}

/**
 * 応募完了後の保存一式
 *
 * 方針（V2改修・2026/10）：
 * - saveApplicationFromState_ がロック内で応募No採番〜
 *   応募管理シート保存〜対応管理追加までを行い、応募Noを返す。
 * - Contact更新・対応履歴保存・userState削除は、
 *   応募管理シートへの保存が成功した場合のみ実行される
 *   （この関数全体が completeApplication_ 側の try-catch で
 *   保護されているため、途中で例外が出れば以降は実行されない）。
 * - 管理者LINE通知（外部API）は、保存処理がすべて完了した
 *   最後に呼ぶ。失敗しても保存結果には影響しない
 *   （notifyApplicationToAdminSafely_ 内部で握りつぶされる）。
 */
function saveCompletedApplication_(
  userId,
  displayName,
  normalizedState
) {
  const applicationNo =
    saveApplicationFromState_(
      normalizedState,
      userId,
      displayName
    );

  updateContactApplicationStatus_(
    userId,
    '応募完了',
    applicationNo
  );

  saveSupportHistory_({
    userId,
    displayName: displayName || normalizedState.displayName || '',
    type: '応募完了',
    message: normalizedState.applicationMessage || '',
    supportNo: '',
    applicationNo
  });

  markUserDone_(
    userId
  );

  notifyApplicationToAdminSafely_(
    applicationNo,
    normalizedState,
    userId,
    displayName
  );
}

/**
 * 名前＋年齢の自由入力を解析
 */
function parseApplicantNameAndAge_(answer) {
  const rawText =
    String(answer || '').trim();

  if (!rawText) {
    return {
      name: '',
      age: '',
      rawText: ''
    };
  }

  const normalizedText =
    normalizeNumberText_(rawText);

  const age =
    extractApplicantAge_(normalizedText);

  let name =
    normalizedText;

  if (age) {
    name = removeAgeFromApplicantText_(
      normalizedText,
      age
    );
  }

  return {
    name,
    age,
    rawText
  };
}

/**
 * 年齢抽出
 */
function extractApplicantAge_(text) {
  const ageWithSuffixMatch =
    text.match(/([0-9]{1,2})\s*(歳|才)/);

  if (ageWithSuffixMatch) {
    const ageNumber =
      Number(ageWithSuffixMatch[1]);

    if (isValidApplicantAge_(ageNumber)) {
      return String(ageNumber);
    }
  }

  const numberMatches =
    text.match(/[0-9]{1,2}/g) || [];

  for (let i = 0; i < numberMatches.length; i++) {
    const ageNumber =
      Number(numberMatches[i]);

    if (isValidApplicantAge_(ageNumber)) {
      return String(ageNumber);
    }
  }

  return '';
}

/**
 * 年齢部分を除去して名前を作る
 */
function removeAgeFromApplicantText_(
  text,
  age
) {
  return String(text || '')
    .replace(new RegExp(age + '\\s*(歳|才)?'), '')
    .replace(/[、,，]/g, ' ')
    .replace(/\r\n/g, ' ')
    .replace(/\r/g, ' ')
    .replace(/\n/g, ' ')
    .replace(/[ \t　]+/g, ' ')
    .trim();
}

/**
 * 全角数字を半角数字へ変換
 */
function normalizeNumberText_(text) {
  return String(text || '')
    .replace(/[０-９]/g, function(char) {
      return String.fromCharCode(
        char.charCodeAt(0) - 0xFEE0
      );
    });
}

/**
 * 応募者年齢として扱う範囲か判定
 */
function isValidApplicantAge_(age) {
  return !isNaN(age) && age >= 15 && age <= 60;
}

function normalizeApplicationStateBeforeComplete_(state) {
  const extraAnswers =
    parseJsonSafe_(state.extraAnswers || '{}');

  const nameSource =
    pickFirstValue_([
      state.name,
      extraAnswers['名前'],
      extraAnswers['お名前'],
      extraAnswers['氏名']
    ]);

  const ageSource =
    pickFirstValue_([
      state.age,
      extraAnswers['年齢']
    ]);

  const sourceText =
    [nameSource, ageSource]
      .filter(value => String(value || '').trim() !== '')
      .join(' ');

  const parsed =
    parseApplicantNameAndAge_(sourceText);

  const normalizedState = {
    ...state
  };

  if (parsed.name) {
    normalizedState.name =
      parsed.name;

    extraAnswers['名前'] =
      parsed.name;

    extraAnswers['お名前'] =
      parsed.name;
  } else if (nameSource) {
    normalizedState.name =
      nameSource;

    extraAnswers['名前'] =
      nameSource;

    extraAnswers['お名前'] =
      nameSource;
  }

  if (parsed.age) {
    normalizedState.age =
      parsed.age;

    extraAnswers['年齢'] =
      parsed.age;
  } else if (ageSource) {
    normalizedState.age =
      ageSource;

    extraAnswers['年齢'] =
      ageSource;
  }

  normalizedState.extraAnswers =
    JSON.stringify(extraAnswers);

  return normalizedState;
}

/**
 * 応募完了時の返信メッセージ
 */
function buildCompleteReplyMessages_(state) {
  const age =
    Number(state.age || 0);

  if (age > 0 && age < 18) {
    return [
      {
        type: 'text',
        text: [
          'ご応募ありがとうございます。',
          '',
          '大変申し訳ございませんが、当店では18歳未満の方のご応募はご遠慮いただいております。',
          'ご理解のほどよろしくお願いいたします。'
        ].join('\n')
      }
    ];
  }

  if (age === 18) {
    return [
      {
        type: 'text',
        text: [
          'ありがとうございます。',
          '',
          '念のため確認ですが、高校生ではございませんか？',
          '18歳未満及び高校生は雇用出来ないため、確認させていただいております。'
        ].join('\n')
      }
    ];
  }

  return [
    {
      type: 'text',
      text:
        getConfigValue_('COMPLETE_MESSAGE') ||
        '✅ ご応募ありがとうございました！\n\n内容を確認後、担当者よりご連絡いたします。'
    }
  ];
}