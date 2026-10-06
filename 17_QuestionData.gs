/**
 * 17_QuestionData.gs
 * 質問設定・選択肢設定の取得
 *
 * 方針：
 * - 質問設定シート、選択肢設定シートから有効データのみ取得
 * - CacheService で短時間キャッシュ
 * - 既存関数名は維持
 */

function getActiveQuestions_() {
  const cached =
    getJsonCache_('activeQuestions');

  if (cached) {
    return cached;
  }

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_QUESTIONS);

  if (!sheet) return [];

  const values =
    sheet.getDataRange().getValues();

  if (values.length <= 1) return [];

  const questions =
    values
      .slice(1)
      .filter(isActiveQuestionRow_)
      .map(questionRowToObject_)
      .sort((a, b) => Number(a.order) - Number(b.order));

  putJsonCache_(
    'activeQuestions',
    questions
  );

  return questions;
}

function getChoicesByItemKey_(itemKey) {
  const targetKey =
    String(itemKey || '').trim();

  if (!targetKey) return [];

  return getAllChoices_()
    .filter(choice => choice.itemKey === targetKey)
    .sort((a, b) => Number(a.order) - Number(b.order));
}

function getAllChoices_() {
  const cached =
    getJsonCache_('allChoices');

  if (cached) {
    return cached;
  }

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_CHOICES);

  if (!sheet) return [];

  const values =
    sheet.getDataRange().getValues();

  if (values.length <= 1) return [];

  const choices =
    values
      .slice(1)
      .filter(isActiveChoiceRow_)
      .map(choiceRowToObject_)
      .sort((a, b) => Number(a.order) - Number(b.order));

  putJsonCache_(
    'allChoices',
    choices
  );

  return choices;
}

function findQuestionByOrder_(
  questions,
  order
) {
  return (questions || []).find(q =>
    Number(q.order) === Number(order)
  ) || null;
}

function getNextQuestion_(
  questions,
  currentOrder
) {
  return (questions || []).find(q =>
    Number(q.order) > Number(currentOrder)
  ) || null;
}

/**
 * 質問設定の有効行判定
 *
 * 想定列：
 * A 順番
 * B 項目キー
 * C 項目名
 * D 質問文
 * E 回答形式
 * F 必須
 * G 保存列名
 * H 有効
 */
function isActiveQuestionRow_(row) {
  if (!row) return false;

  const order =
    row[0];

  if (order === '' || order === '順番') {
    return false;
  }

  if (isNaN(Number(order))) {
    return false;
  }

  const enabled =
    String(row[7] || '')
      .trim()
      .toUpperCase();

  return enabled === 'TRUE';
}

/**
 * 質問設定行をオブジェクト化
 */
function questionRowToObject_(row) {
  return {
    order: Number(row[0]),
    itemKey: String(row[1] || '').trim(),
    itemName: String(row[2] || '').trim(),
    questionText: String(row[3] || '').trim(),
    answerType: String(row[4] || '').trim(),
    required: String(row[5] || '').trim().toUpperCase() === 'TRUE',
    saveColumnName: String(row[6] || '').trim(),
    enabled: String(row[7] || '').trim().toUpperCase() === 'TRUE'
  };
}

/**
 * 選択肢設定の有効行判定
 *
 * 想定列：
 * A 項目キー
 * B 選択肢名
 * C 送信テキスト
 * D 表示順
 * E 有効
 */
function isActiveChoiceRow_(row) {
  if (!row) return false;

  const itemKey =
    String(row[0] || '').trim();

  if (!itemKey) {
    return false;
  }

  const enabled =
    String(row[4] || '')
      .trim()
      .toUpperCase();

  return enabled === 'TRUE';
}

/**
 * 選択肢設定行をオブジェクト化
 */
function choiceRowToObject_(row) {
  return {
    itemKey: String(row[0] || '').trim(),
    choiceName: String(row[1] || '').trim(),
    sendText: String(row[2] || '').trim(),
    order: Number(row[3] || 0),
    enabled: String(row[4] || '').trim().toUpperCase() === 'TRUE'
  };
}

/**
 * JSONキャッシュ取得
 */
function getJsonCache_(key) {
  const cache =
    CacheService.getScriptCache();

  const cached =
    cache.get(key);

  if (!cached) return null;

  try {
    return JSON.parse(cached);
  } catch (error) {
    cache.remove(key);
    return null;
  }
}

/**
 * JSONキャッシュ保存
 */
function putJsonCache_(
  key,
  value
) {
  CacheService
    .getScriptCache()
    .put(
      key,
      JSON.stringify(value),
      CACHE_SECONDS
    );
}