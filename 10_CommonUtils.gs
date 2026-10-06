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
 * 表形式シートの2行目に「最新V2説明行」を安全に保証する共通ヘルパー。
 *
 * 方針（2026/10改訂・説明行統一）：
 * 最終仕様は「1行目=最新ヘッダー／2行目=各列の説明／3行目以降=実データ」。
 * ただし再実行のたびに説明行を増殖させず、かつ既存の実データ行を
 * 絶対に説明行で上書きしないことを最優先する。
 *
 * 判定ロジック：
 * - 2行目が存在しない（新規シート）→ そのまま2行目に説明行を書く
 *   （保護すべきデータが無い）。
 * - 2行目が空白行 → 同上、そのまま2行目に説明行を書く。
 * - 2行目が「今回書こうとしている説明行」と完全一致 → 既に説明行が
 *   設定済みとみなし、行を増やさずその場で上書き（内容的には無変化、
 *   または文言更新のみ）。
 * - 上記のいずれにも該当しない → 既存の実データ行である可能性がある
 *   ため、必ず「データ」側として扱い、2行目の手前に新しい行を挿入して
 *   から説明行を書く（既存データは3行目以降へ安全に押し下げられる）。
 *   判定に迷うケースも必ずこちら側（行を挿入する側）に倒す。
 *
 * @param {Sheet} sheet 対象シート
 * @param {Array} descriptionValues 最新の説明行（ヘッダーと同じ列数）
 * @return {boolean} true = 既存データ保護のため行を挿入した／false = 挿入なし
 */
function ensureDescriptionRow_(sheet, descriptionValues) {
  if (!sheet || !descriptionValues || descriptionValues.length === 0) {
    return false;
  }

  const columnCount = descriptionValues.length;
  const lastRow = sheet.getLastRow();

  let insertedRow = false;

  if (lastRow >= 2) {
    const currentRow2 =
      sheet.getRange(2, 1, 1, columnCount).getValues()[0];

    const currentRow2Trimmed =
      currentRow2.map(v => String(v || '').trim());

    const isBlankRow2 =
      currentRow2Trimmed.every(v => v === '');

    const isAlreadyDescriptionRow =
      !isBlankRow2 &&
      descriptionValues.every(
        (v, i) => String(v || '').trim() === currentRow2Trimmed[i]
      );

    if (!isBlankRow2 && !isAlreadyDescriptionRow) {
      // 2行目は既存の実データとみなし、保護のため手前に新しい行を挿入する。
      // （挿入された新しい2行目は、挿入元のセルの書式・入力規則を
      // 引き継ぐ場合があるため、この後の clearDataValidations() で
      // 必ず解除する）。
      sheet.insertRowBefore(2);
      insertedRow = true;
    }
  }

  const descriptionRange =
    sheet.getRange(2, 1, 1, columnCount);

  // 2行目は説明専用とし、データ用の入力規則（プルダウン・チェックボックス等）
  // を一切残さない。旧バージョンで2行目がデータ行だった頃に設定された
  // 入力規則がセルに残ったままになっている場合があり、それを解除せずに
  // 説明文を書き込むと「入力規則に違反しています」というエラーで
  // setup自体が失敗する（2026/10・質問設定E2で発生した不具合）。
  // clearDataValidations() は入力規則のみを解除し、背景色・文字色・
  // 折り返しなどの書式は変更しない。
  descriptionRange.clearDataValidations();
  descriptionRange.setValues([descriptionValues]);

  return insertedRow;
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
/**
 * エラーログ・処理ログ共通の説明行（2行目）配列
 * （11_SetupSpreadsheet.gs の setupErrorLogSheetForRecruit_ /
 * setupProcessLogSheetForRecruit_ と必ず一致させること）
 */
function getLogDescriptionRow_() {
  return [
    '自動',
    'システム内部管理',
    'エラー内容／ログ内容',
    'LINEから取得（任意）',
    '補足メモ（任意）'
  ];
}

/**
 * ログシート（エラーログ・処理ログ）のヘッダー保証
 *
 * 方針（2026/10改訂・説明行統一）：1行目＝最新ヘッダー／2行目＝最新説明／
 * 3行目以降＝ログ（追記専用）。appendRow() は常に最終行の次へ追記する
 * ため、2行目に説明行を挿入しても追記処理自体は影響を受けない。
 * 2行目は ensureDescriptionRow_ により、既存のログデータ行を保護
 * しながら安全に挿入・更新する。
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

  ensureDescriptionRow_(sheet, getLogDescriptionRow_());
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

/**
 * 応募管理・対応管理のNo列の値が「有効なデータ行」かどうかを判定する。
 *
 * 対応管理・対応履歴では、No列の空でない値のうち "No"（ヘッダー文字列）
 * や "自動"（説明文）のような非データ文字列が紛れ込むことがあるため、
 * 「空でなければデータ行とみなす」という単純な判定だとヘッダー・説明行
 * まで誤ってデータ行として扱ってしまう。
 *
 * 有効な対応管理Noは「数値」または「問い合わせ-」で始まる文字列のみ
 * なので、それ以外（"No" や "自動" など）は確実に除外する。
 */
function isLikelySupportNo_(value) {
  const text = String(value || '').trim();

  if (!text) return false;

  if (text.indexOf('問い合わせ-') === 0) return true;

  return !isNaN(Number(text));
}