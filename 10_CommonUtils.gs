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

/**
 * ============================================================
 * V2ヘッダー追記方式 共通ユーティリティ（2026/10 V2移行改修）
 * ============================================================
 *
 * 方針：
 * - 旧バージョン（本番からコピーされた可能性のある既存ヘッダー・
 *   既存データ）は一切削除・上書き・移動しない。
 * - V2のヘッダーは「既存データの最終行の次」に追記する。
 *   シートが空（新規）の場合はそのまま1行目へ作成する。
 * - setupを何度再実行してもV2ヘッダーを増殖させない
 *   （既に存在する場合はそれを検出して再利用する）。
 * - 識別方法：シート内でヘッダー配列（各セル文字列）が完全一致する
 *   「最も下側の行」をV2ヘッダーとして扱う。
 *   Script Propertiesに行番号をキャッシュして毎回のフルスキャンを
 *   避けるが、キャッシュは必ず実際のシート内容と突き合わせて検証し、
 *   一致しない場合はスキャンし直す（スプレッドシートを複製した場合、
 *   コピー後の新しいスクリプトプロジェクトにはScript Propertiesの
 *   キャッシュが存在しないため、この検証つきスキャンが唯一の
 *   信頼できる手段となる）。
 */

/**
 * ヘッダー配列同士が完全一致するか判定
 * （前後の空白を除去したうえで文字列として比較）
 */
function headersEqual_(a, b) {
  if (!a || !b || a.length !== b.length) return false;

  for (let i = 0; i < b.length; i++) {
    if (String(a[i] || '').trim() !== String(b[i] || '').trim()) {
      return false;
    }
  }

  return true;
}

/**
 * シート内で、指定したヘッダー配列と完全一致する最も下側の行番号を探す。
 * 見つからない場合は0を返す。
 */
function findLatestHeaderRow_(sheet, expectedHeaders) {
  if (!sheet || !expectedHeaders || expectedHeaders.length === 0) {
    return 0;
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 1) return 0;

  const values =
    sheet
      .getRange(1, 1, lastRow, expectedHeaders.length)
      .getValues();

  for (let r = values.length - 1; r >= 0; r--) {
    if (headersEqual_(values[r], expectedHeaders)) {
      return r + 1;
    }
  }

  return 0;
}

/**
 * V2ヘッダー行を解決する（Script Propertiesキャッシュ＋検証つきフォール
 * バックスキャン）。見つからない場合は0を返す（＝V2ヘッダー未作成）。
 */
function resolveV2HeaderRow_(sheet, expectedHeaders) {
  if (!sheet) return 0;

  const cacheKey = `v2HeaderRow:${sheet.getName()}`;
  const props = PropertiesService.getScriptProperties();
  const cached = Number(props.getProperty(cacheKey) || 0);

  if (cached > 0 && cached <= sheet.getLastRow()) {
    const cachedValues =
      sheet
        .getRange(cached, 1, 1, expectedHeaders.length)
        .getValues()[0];

    if (headersEqual_(cachedValues, expectedHeaders)) {
      return cached;
    }
  }

  const found = findLatestHeaderRow_(sheet, expectedHeaders);

  if (found > 0) {
    props.setProperty(cacheKey, String(found));
  }

  return found;
}

/**
 * V2ヘッダーブロックを保証する（冪等）。
 *
 * - 既にV2ヘッダー（expectedHeadersと完全一致する行）が存在する場合は
 *   何もせず、その行番号を返す。
 * - 存在しない場合は、現在の最終行の次（シートが空なら1行目）に
 *   ヘッダー行（＋noteRowが指定されていれば説明行）を新規作成する。
 *   既存データは一切変更しない。
 *
 * 戻り値：{ headerRow, dataStartRow, created }
 *   headerRow    … ヘッダー行の行番号
 *   dataStartRow … データ入力開始行（ヘッダーのみなら headerRow+1、
 *                  説明行ありなら headerRow+2）
 *   created      … 今回新規作成したか（false＝既存のV2ヘッダーを再利用）
 */
function ensureV2HeaderBlock_(sheet, expectedHeaders, noteRow) {
  const existing = resolveV2HeaderRow_(sheet, expectedHeaders);
  const blockSize = noteRow ? 2 : 1;

  if (existing > 0) {
    return {
      headerRow: existing,
      dataStartRow: existing + blockSize,
      created: false
    };
  }

  const headerRow = sheet.getLastRow() + 1;

  sheet
    .getRange(headerRow, 1, 1, expectedHeaders.length)
    .setValues([expectedHeaders]);

  if (noteRow) {
    sheet
      .getRange(headerRow + 1, 1, 1, noteRow.length)
      .setValues([noteRow]);
  }

  PropertiesService
    .getScriptProperties()
    .setProperty(`v2HeaderRow:${sheet.getName()}`, String(headerRow));

  return {
    headerRow,
    dataStartRow: headerRow + blockSize,
    created: true
  };
}

/**
 * 応募管理・対応管理のNo列の値が「有効なデータ行」かどうかを判定する。
 *
 * V2ヘッダー追記方式では、シートの途中（旧データと新データの境目）に
 * もう1組のヘッダー行・説明行が挟まる。これらの行のNo列には
 * "No"（ヘッダー文字列）や "自動"（説明文）といった非空文字列が
 * 入っているため、「空でなければデータ行とみなす」という単純な判定だと
 * ヘッダー・説明行まで誤ってデータ行として扱ってしまう。
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