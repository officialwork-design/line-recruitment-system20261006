/**
 * 6_ApplicationQuestions.gs
 * 応募質問送信・媒体分岐・選択肢Flex
 */

/**
 * その他媒体フローか判定
 */
function isMediaOtherFlow_(state) {
  return !!(
    state &&
    String(state.mediaOtherFlow || '').toUpperCase() === 'TRUE'
  );
}

/**
 * 質問番号の表示オフセット
 */
function getQuestionOffset_(state) {
  return isMediaOtherFlow_(state) ? 2 : 1;
}

/**
 * 通常質問の進捗テキスト
 */
function getQuestionProgressText_(question, state) {
  const questions =
    getNormalQuestionsBeforePhoto_();

  const currentIndex =
    questions.findIndex(q =>
      Number(q.order) === Number(question.order)
    ) + 1;

  if (currentIndex <= 0) return '';

  const total =
    getTotalApplicationSteps_(state);

  const displayNumber =
    currentIndex + getQuestionOffset_(state);

  return `📮 質問 ${displayNumber} / ${total}\n\n`;
}

/**
 * 事前確認事項の質問文言（応募フロー簡略化・2026/10追加）
 *
 * 方針：
 * - 旧・備考欄（質問設定シートの remarks 行）を廃止し、
 *   募集媒体の質問（質問1）と同様に、質問設定シートに依存しない
 *   固定文言として送信する。
 * - 通常質問が何問あっても（0問でも）、この質問は必ず
 *   「写真提出の直前」に1回だけ実行される。
 * - 回答後は completeApplication_ を直接呼ばず、写真提出ステップ
 *   （moveToPhotoStep_）へ進む。
 */
const CONFIRM_QUESTION_TEXT =
  '事前に伝えておきたいことや、確認しておきたいことがあれば教えてください！\n特になければ「なし」と送ってください。';

/**
 * 事前確認事項の進捗テキスト
 *
 * 常に「写真提出の1つ前」の番号として表示する
 * （質問 total-1 / total）。
 */
function getConfirmProgressText_(state) {
  const total =
    getTotalApplicationSteps_(state);

  return `📮 質問 ${total - 1} / ${total}\n\n`;
}

/**
 * 事前確認事項の質問送信
 */
function sendConfirmQuestion_(
  replyToken,
  userId,
  state
) {
  replyText_(
    replyToken,
    `${getConfirmProgressText_(state)}${CONFIRM_QUESTION_TEXT}`,
    'CONFIRM_QUESTION_ERROR',
    userId
  );
}

/**
 * 事前確認事項の回答処理
 *
 * 方針（V2改修・2026/10：応募フロー簡略化）：
 * - 自由入力を受け付ける（「なし」を含む）。空回答のみ再入力を促す。
 * - 回答後は completeApplication_ を直接呼ばず、写真提出ステップへ進む。
 */
function handleConfirmAnswer_(
  userId,
  displayName,
  replyToken,
  state,
  text
) {
  if (!isValidTextAnswer_(text, true)) {
    replyText_(
      replyToken,
      [
        '⚠️ 必須項目です。',
        '入力をお願いします。特になければ「なし」と送ってください。',
        '',
        `${getConfirmProgressText_(state)}${CONFIRM_QUESTION_TEXT}`
      ].join('\n'),
      'EMPTY_CONFIRM_ANSWER_ERROR',
      userId
    );

    return;
  }

  const updatedState =
    applyConfirmAnswerToState_(
      state,
      text,
      displayName
    );

  moveToPhotoStep_(
    userId,
    replyToken,
    updatedState
  );
}

/**
 * 事前確認事項の回答をstateへ反映
 *
 * 既存の応募管理シート列「備考欄」をそのまま使用する
 * （新しい列は追加しない）。
 */
function applyConfirmAnswerToState_(
  state,
  answer,
  displayName
) {
  const extraAnswers =
    parseJsonSafe_(state.extraAnswers || '{}');

  extraAnswers['備考欄'] =
    answer;

  const mergedApplicationMessage =
    buildApplicationMessage_({
      currentMessage: state.applicationMessage || '',
      label: '事前確認事項',
      answer
    });

  return {
    ...state,
    displayName: displayName || state.displayName || '',
    remarks: answer,
    extraAnswers: JSON.stringify(extraAnswers),
    applicationMessage: mergedApplicationMessage,
    updatedAt: new Date().toISOString()
  };
}

/**
 * LINE表示用に質問文を整形
 */
function cleanQuestionText_(text) {
  return String(text || '')
    .replace(/※LINE表示では質問\s*\d+\s*になります。?/g, '')
    .replace(/※LINE表示では質問\s*[0-9０-９]+\s*になります。?/g, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/**
 * 質問1：募集媒体
 */
function sendMediaQuestion_(
  replyToken,
  userId
) {
  const total =
    getMaxApplicationSteps_();

  const questionText = [
    '✨ ご応募ありがとうございます！',
    '',
    '順番にご回答をお願いいたします。',
    '',
    `📮 質問 1 / ${total}`,
    '',
    '何を見て応募しましたか？'
  ].join('\n');

  processLog_(
    'QUESTION_1_SENT',
    '質問1を送信しました：募集媒体',
    userId,
    `total=${total}`
  );

  replyMessages_(
    replyToken,
    [createMediaQuestionMessage_(questionText)],
    'MEDIA_MESSAGE_ERROR',
    userId
  );
}

/**
 * 募集媒体の不正回答時
 */
function sendInvalidMediaQuestion_(
  replyToken,
  userId
) {
  const total =
    getMaxApplicationSteps_();

  const questionText = [
    '⚠️ 募集媒体はボタンから選択してください。',
    '',
    '✨ ご応募ありがとうございます！',
    '',
    '順番にご回答をお願いいたします。',
    '',
    `📮 質問 1 / ${total}`,
    '',
    '何を見て応募しましたか？'
  ].join('\n');

  replyMessages_(
    replyToken,
    [createMediaQuestionMessage_(questionText)],
    'INVALID_MEDIA_REPLY_ERROR',
    userId
  );
}

/**
 * 募集媒体Flex
 */
function createMediaQuestionMessage_(questionText) {
  return createButtonQuestionMessage_(
    '📮 質問 1',
    questionText,
    [
      {
        label: '☕ カフェるん',
        text: 'カフェるん',
        color: '#06C755'
      },
      {
        label: '📸 Instagram',
        text: 'Instagram広告',
        color: '#E4405F'
      },
      {
        label: '𝕏 X',
        text: 'X',
        color: '#111111'
      },
      {
        label: 'その他',
        text: 'その他',
        color: '#666666'
      }
    ]
  );
}

/**
 * 募集媒体回答処理
 */
function handleMediaAnswer_(
  userId,
  displayName,
  replyToken,
  state,
  text
) {
  processLog_(
    'QUESTION_1_ANSWER',
    `質問1の回答を受信しました：${text}`,
    userId,
    displayName
  );

  if (!MEDIA_CHOICES.includes(text)) {
    sendInvalidMediaQuestion_(
      replyToken,
      userId
    );

    return;
  }

  const updatedState =
    buildMediaAnsweredState_(
      userId,
      displayName,
      state,
      text
    );

  if (text === 'その他') {
    updatedState.status =
      STATUS_WAIT_MEDIA_OTHER;

    updatedState.currentQuestionNo =
      '';

    upsertUserManagement_(
      updatedState
    );

    sendMediaOtherQuestion_(
      replyToken,
      userId,
      updatedState
    );

    return;
  }

  goToFirstNormalQuestionOrPhoto_(
    userId,
    displayName,
    replyToken,
    updatedState
  );
}

/**
 * 募集媒体回答後のstate作成
 */
function buildMediaAnsweredState_(
  userId,
  displayName,
  state,
  text
) {
  const extraAnswers =
    parseJsonSafe_(state.extraAnswers || '{}');

  extraAnswers['募集媒体'] =
    text;

  return {
    ...state,
    userId,
    displayName: displayName || state.displayName || '',
    media: text,
    mediaOtherFlow: text === 'その他' ? 'TRUE' : 'FALSE',
    totalSteps: text === 'その他'
      ? getMaxApplicationSteps_()
      : getBaseApplicationSteps_(),
    extraAnswers: JSON.stringify(extraAnswers),
    applicationMessage: buildApplicationMessage_({
      currentMessage: state.applicationMessage || '',
      label: '募集媒体',
      answer: text
    }),
    updatedAt: new Date().toISOString()
  };
}

/**
 * 質問2：その他媒体
 */
function sendMediaOtherQuestion_(
  replyToken,
  userId,
  state
) {
  const total =
    getTotalApplicationSteps_(state);

  const questionText =
    `📮 質問 2 / ${total}\n\nその他の場合、何を見て応募しましたか？`;

  replyMessages_(
    replyToken,
    [createMediaOtherQuestionMessage_(questionText)],
    'MEDIA_OTHER_MESSAGE_ERROR',
    userId
  );
}

/**
 * その他媒体の不正回答時
 */
function sendInvalidMediaOtherQuestion_(
  replyToken,
  userId,
  state
) {
  const total =
    getTotalApplicationSteps_(state);

  const questionText = [
    '⚠️ 選択肢から選んでください。',
    'もう一度お願いします。',
    '',
    `📮 質問 2 / ${total}`,
    '',
    'その他の場合、何を見て応募しましたか？'
  ].join('\n');

  replyMessages_(
    replyToken,
    [createMediaOtherQuestionMessage_(questionText)],
    'INVALID_MEDIA_OTHER_MESSAGE_ERROR',
    userId
  );
}

/**
 * その他媒体Flex
 */
function createMediaOtherQuestionMessage_(questionText) {
  return createButtonQuestionMessage_(
    '📮 質問 2',
    questionText,
    [
      {
        label: '▶️ YouTube',
        text: 'YouTube',
        color: '#FF0000'
      },
      {
        label: '🎵 TikTok',
        text: 'TikTok',
        color: '#111111'
      },
      {
        label: '👥 紹介',
        text: '紹介',
        color: '#1E88E5'
      }
    ]
  );
}

/**
 * その他媒体回答処理
 */
function handleMediaOtherAnswer_(
  userId,
  displayName,
  replyToken,
  state,
  text
) {
  if (!MEDIA_OTHER_CHOICES.includes(text)) {
    sendInvalidMediaOtherQuestion_(
      replyToken,
      userId,
      state
    );

    return;
  }

  const updatedState =
    buildMediaOtherAnsweredState_(
      state,
      text
    );

  goToFirstNormalQuestionOrPhoto_(
    userId,
    displayName,
    replyToken,
    updatedState
  );
}

/**
 * その他媒体回答後のstate作成
 */
function buildMediaOtherAnsweredState_(
  state,
  text
) {
  const extraAnswers =
    parseJsonSafe_(state.extraAnswers || '{}');

  extraAnswers['募集媒体'] =
    text;

  extraAnswers['募集媒体大分類'] =
    'その他';

  extraAnswers['その他媒体'] =
    text;

  return {
    ...state,
    media: text,
    mediaOtherFlow: 'TRUE',
    totalSteps: getMaxApplicationSteps_(),
    extraAnswers: JSON.stringify(extraAnswers),
    applicationMessage: buildApplicationMessage_({
      currentMessage: state.applicationMessage || '',
      label: 'その他媒体',
      answer: text
    }),
    updatedAt: new Date().toISOString()
  };
}

/**
 * 通常質問送信
 */
function sendQuestion_(
  replyToken,
  question,
  userId,
  state
) {
  if (!question) return;

  const questionTextWithProgress =
    `${getQuestionProgressText_(question, state)}${cleanQuestionText_(question.questionText)}`;

  if (question.answerType === '選択ボタン') {
    sendChoiceQuestion_(
      replyToken,
      question,
      questionTextWithProgress,
      userId
    );

    return;
  }

  replyText_(
    replyToken,
    questionTextWithProgress,
    'QUESTION_TEXT_REPLY_ERROR',
    userId
  );
}

/**
 * 選択ボタン質問送信
 */
function sendChoiceQuestion_(
  replyToken,
  question,
  questionTextWithProgress,
  userId
) {
  const choices =
    getChoicesByItemKey_(question.itemKey);

  if (choices.length === 0) {
    errorLog_(
      'NO_CHOICES_ERROR',
      `選択肢が未設定です：${question.itemKey}`,
      userId,
      question.itemName || ''
    );

    replyText_(
      replyToken,
      `${questionTextWithProgress}\n\n⚠️ 選択肢が未設定です。`,
      'NO_CHOICES_ERROR',
      userId
    );

    return;
  }

  replyMessages_(
    replyToken,
    [createChoiceButtonMessage_(questionTextWithProgress, choices)],
    'CHOICE_MESSAGE_ERROR',
    userId
  );
}

/**
 * 選択肢不正時の文言
 */
function getInvalidChoiceMessageByQuestion_(question) {
  const itemKey =
    String(question.itemKey || '').trim();

  if (itemKey === 'store') {
    return '⚠️ 希望店舗はボタンから選択してください。';
  }

  if (itemKey === 'work_days') {
    return '⚠️ 勤務日数はボタンから選択してください。';
  }

  return '⚠️ 選択肢から選んでください。\nもう一度お願いします。';
}

/**
 * 選択肢不正時に同じ質問を再送
 */
function sendInvalidChoiceQuestion_(
  replyToken,
  question,
  choices,
  userId,
  state
) {
  const questionTextWithProgress =
    `${getInvalidChoiceMessageByQuestion_(question)}\n\n` +
    `${getQuestionProgressText_(question, state)}` +
    `${cleanQuestionText_(question.questionText)}`;

  if (!choices || choices.length === 0) {
    replyText_(
      replyToken,
      `${questionTextWithProgress}\n\n⚠️ 選択肢が未設定です。`,
      'INVALID_CHOICE_NO_CHOICES_ERROR',
      userId
    );

    return;
  }

  replyMessages_(
    replyToken,
    [createChoiceButtonMessage_(questionTextWithProgress, choices)],
    'INVALID_CHOICE_REPLY_ERROR',
    userId
  );
}

/**
 * 通常選択肢Flex
 */
function createChoiceButtonMessage_(
  questionText,
  choices
) {
  const buttonItems =
    choices.map(choice => ({
      label: formatChoiceLabel_(choice.choiceName),
      text: choice.sendText,
      color: '#06C755'
    }));

  return createButtonQuestionMessage_(
    questionText,
    questionText,
    buttonItems
  );
}

/**
 * ボタン型Flex共通作成
 */
function createButtonQuestionMessage_(
  altText,
  questionText,
  buttonItems
) {
  return {
    type: 'flex',
    altText: altText || '質問',
    contents: {
      type: 'bubble',
      body: {
        type: 'box',
        layout: 'vertical',
        spacing: 'md',
        contents: [
          {
            type: 'text',
            text: questionText,
            weight: 'bold',
            size: 'md',
            wrap: true
          },
          ...buttonItems.map(item => ({
            type: 'button',
            style: 'primary',
            color: item.color || '#06C755',
            height: 'sm',
            action: {
              type: 'message',
              label: item.label,
              text: item.text
            }
          }))
        ]
      }
    }
  };
}

/**
 * 選択肢ラベル整形
 */
function formatChoiceLabel_(text) {
  const value =
    String(text || '').trim();

  const emojiMap = {
    'はい': '⭕',
    'いいえ': '❌',
    'あり': '⭕',
    'なし': '❌',
    '未経験': '🔰',
    '経験なし': '🔰',
    '経験無し': '🔰',
    '経験あり': '☕',
    '経験有り': '☕',
    '相談したい': '💬',
    '週1〜2日': '📅',
    '週2〜3日': '📅',
    '週4〜5日': '📅'
  };

  const storeValues = [
    '大阪店',
    '新宿店',
    '秋葉原店',
    '複数店',
    '複数店(新宿店/秋葉原店)'
  ];

  if (storeValues.includes(value)) {
    return value;
  }

  const emoji =
    emojiMap[value];

  if (emoji) {
    return `${emoji} ${value}`;
  }

  return value;
}
