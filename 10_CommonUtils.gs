/**
 * 10_CommonUtils.gs
 * 共通ユーティリティ
 */

/**
 * シート取得。なければ作成。
 */
function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/**
 * JSON安全パース
 */
function parseJsonSafe_(text) {
  try {
    return JSON.parse(text || '{}');
  } catch (error) {
    return {};
  }
}

/**
 * ヘッダー名に一致する列へ値をセット
 */
function setRowValueByHeader_(
  headers,
  row,
  headerName,
  value
) {
  const normalizedHeaders =
    headers.map(h => String(h || '').trim());

  const index =
    normalizedHeaders.indexOf(
      String(headerName || '').trim()
    );

  if (index === -1) return;

  row[index] = value;
}

/**
 * 同期用ヘッダー正規化
 */
function normalizeSyncHeader_(value) {
  return String(value || '')
    .replace(/\s/g, '')
    .replace(/　/g, '')
    .trim();
}

/**
 * 設定値取得
 */
function getConfigValue_(key) {
  const cache =
    CacheService.getScriptCache();

  const cacheKey =
    `config:${key}`;

  const cached =
    cache.get(cacheKey);

  if (cached !== null) {
    return cached;
  }

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_CONFIG);

  if (!sheet) return '';

  const values =
    sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    const rowKey =
      String(values[i][0] || '').trim();

    if (rowKey === key) {
      const value =
        String(values[i][1] || '');

      cache.put(
        cacheKey,
        value,
        CACHE_SECONDS
      );

      return value;
    }
  }

  return '';
}

/**
 * 設定・質問・選択肢キャッシュをクリア
 * UIなしで内部処理から呼ぶ用
 */
function clearConfigCache_() {
  CacheService
    .getScriptCache()
    .removeAll([
      'activeQuestions',
      'allChoices',
      'config:LINE_CHANNEL_ACCESS_TOKEN',
      'config:START_WORD',
      'config:COMPLETE_MESSAGE',
      'config:APPLY_BUTTON_USER_ID',
      'config:APPLY_BUTTON_TITLE',
      'config:APPLY_BUTTON_NOTE',
      'config:APPLY_BUTTON_LABEL',
      'config:APPLY_BUTTON_COLOR',
      `config:${CONFIG_ADMIN_GROUP_ID}`
    ]);
}

/**
 * 処理ログ
 */
function processLog_(
  type,
  content,
  userId,
  memo
) {
  if (!ENABLE_PROCESS_LOG) return;

  const allowedTypes =
    typeof PROCESS_LOG_ALLOWED_TYPES !== 'undefined'
      ? PROCESS_LOG_ALLOWED_TYPES
      : [];

  if (!allowedTypes.includes(type)) {
    return;
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    getOrCreateSheet_(
      ss,
      SHEET_PROCESS_LOG
    );

  ensureLogHeader_(sheet);

  sheet.appendRow([
    new Date(),
    type,
    typeof content === 'string'
      ? content
      : JSON.stringify(content),
    userId || '',
    memo || ''
  ]);
}

/**
 * エラーログ
 */
function errorLog_(
  type,
  content,
  userId,
  memo
) {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    getOrCreateSheet_(
      ss,
      SHEET_ERROR_LOG
    );

  ensureLogHeader_(sheet);

  sheet.appendRow([
    new Date(),
    type,
    typeof content === 'string'
      ? content
      : JSON.stringify(content),
    userId || '',
    memo || ''
  ]);
}

/**
 * ログヘッダー保証
 */
function ensureLogHeader_(sheet) {
  if (!sheet) return;

  if (sheet.getLastRow() === 0) {
    sheet
      .getRange(1, 1, 1, 5)
      .setValues([[
        '日時',
        '種別',
        '内容',
        'ユーザーID',
        'メモ'
      ]]);
  }
}

/**
 * Hexカラー正規化
 */
function normalizeHexColor_(value) {
  const text =
    String(value || '').trim();

  if (/^#[0-9A-Fa-f]{6}$/.test(text)) {
    return text;
  }

  if (/^[0-9A-Fa-f]{6}$/.test(text)) {
    return `#${text}`;
  }

  return '#333333';
}

/**
 * XMLエスケープ
 */
function escapeXml_(text) {
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * 日時整形 yyyy/MM/dd HH:mm:ss
 */
function formatDateTime_(value) {
  if (!value) return '';

  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      'Asia/Tokyo',
      'yyyy/MM/dd HH:mm:ss'
    );
  }

  const date = new Date(value);

  if (!isNaN(date.getTime())) {
    return Utilities.formatDate(
      date,
      'Asia/Tokyo',
      'yyyy/MM/dd HH:mm:ss'
    );
  }

  return String(value || '');
}

/**
 * 日時整形 yyyy/MM/dd HH:mm
 */
function formatDateMinute_(value) {
  if (!value) return '';

  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      'Asia/Tokyo',
      'yyyy/MM/dd HH:mm'
    );
  }

  const date = new Date(value);

  if (!isNaN(date.getTime())) {
    return Utilities.formatDate(
      date,
      'Asia/Tokyo',
      'yyyy/MM/dd HH:mm'
    );
  }

  return String(value || '');
}

/**
 * 文字列化
 */
function toSafeString_(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value).trim();
}

/**
 * 空判定
 */
function isBlank_(value) {
  return toSafeString_(value) === '';
}

/**
 * 非空判定
 */
function notBlank_(value) {
  return !isBlank_(value);
}

/**
 * ヘッダーIndex取得
 */
function getHeaderIndex_(headers, headerName) {
  const normalizedHeaders =
    headers.map(h => String(h || '').trim());

  return normalizedHeaders.indexOf(
    String(headerName || '').trim()
  );
}

/**
 * ヘッダー列番号取得
 */
function getHeaderCol_(headers, headerName) {
  const index =
    getHeaderIndex_(
      headers,
      headerName
    );

  return index >= 0 ? index + 1 : 0;
}

/**
 * ヘッダー名からセル値取得
 */
function getCellValueByHeader_(
  sheet,
  headers,
  row,
  headerName
) {
  const col =
    getHeaderCol_(
      headers,
      headerName
    );

  if (col <= 0) return '';

  return sheet
    .getRange(row, col)
    .getValue();
}

/**
 * ヘッダー名からセル値セット
 */
function setCellValueByHeader_(
  sheet,
  headers,
  row,
  headerName,
  value
) {
  const col =
    getHeaderCol_(
      headers,
      headerName
    );

  if (col <= 0) return;

  sheet
    .getRange(row, col)
    .setValue(value);
}

/**
 * ヘッダー名・値から行検索
 */
function findRowByHeaderValue_(
  sheet,
  headerName,
  targetValue,
  startRow
) {
  if (!sheet) return 0;

  const lastRow =
    sheet.getLastRow();

  const lastCol =
    sheet.getLastColumn();

  const beginRow =
    startRow || 2;

  if (lastRow < beginRow) return 0;

  const headers =
    sheet
      .getRange(1, 1, 1, lastCol)
      .getValues()[0];

  const col =
    getHeaderCol_(
      headers,
      headerName
    );

  if (col <= 0) return 0;

  const values =
    sheet
      .getRange(
        beginRow,
        col,
        lastRow - beginRow + 1,
        1
      )
      .getValues();

  const target =
    String(targetValue || '').trim();

  for (let i = 0; i < values.length; i++) {
    const value =
      String(values[i][0] || '').trim();

    if (value === target) {
      return beginRow + i;
    }
  }

  return 0;
}

/**
 * シートヘッダー取得
 */
function getSheetHeaders_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) {
    return [];
  }

  return sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(h => String(h || '').trim());
}

/**
 * アラート表示
 */
function alert_(message) {
  SpreadsheetApp
    .getUi()
    .alert(String(message || ''));
}