/**
 * 19_ContactMessage.gs
 * 連絡先シートから対象者へLINEメッセージ送信
 *
 * 方針：
 * - 送信対象は連絡先シートの「状態」「応募状況」で抽出
 * - LINE送信は1件ずつ実行
 * - 送信失敗しても全体処理は止めない
 * - 送信結果は成功/失敗件数で表示
 */

/**
 * 連絡先にLINEメッセージを送信
 */
function sendLineMessageToContactsByPrompt() {
  const ui =
    SpreadsheetApp.getUi();

  const typePrompt =
    ui.prompt(
      '送信対象',
      [
        '送信対象を入力してください。',
        '',
        '1 = 友だち追加済み',
        '2 = やりとりあり',
        '3 = 応募中',
        '4 = 応募完了',
        '5 = 未応募',
        '6 = 全員'
      ].join('\n'),
      ui.ButtonSet.OK_CANCEL
    );

  if (
    typePrompt.getSelectedButton() !==
    ui.Button.OK
  ) {
    return;
  }

  const type =
    String(
      typePrompt.getResponseText() || ''
    ).trim();

  const targetLabel =
    getContactMessageTargetLabel_(type);

  if (!targetLabel) {
    ui.alert('送信対象が不正です。');
    return;
  }

  const messagePrompt =
    ui.prompt(
      '送信メッセージ',
      [
        `送信対象：${targetLabel}`,
        '',
        '送信したいメッセージを入力してください。'
      ].join('\n'),
      ui.ButtonSet.OK_CANCEL
    );

  if (
    messagePrompt.getSelectedButton() !==
    ui.Button.OK
  ) {
    return;
  }

  const text =
    String(
      messagePrompt.getResponseText() || ''
    ).trim();

  if (!text) {
    ui.alert('メッセージが空です。');
    return;
  }

  const targets =
    getTargetContactsByType_(type);

  if (targets.length === 0) {
    ui.alert('送信対象が0件でした。');
    return;
  }

  const confirm =
    ui.alert(
      'LINE送信確認',
      [
        `送信対象：${targetLabel}`,
        `対象件数：${targets.length}件`,
        '',
        'この内容でLINEメッセージを送信しますか？',
        '',
        '※送信後の取り消しはできません。'
      ].join('\n'),
      ui.ButtonSet.OK_CANCEL
    );

  if (confirm !== ui.Button.OK) {
    return;
  }

  const result =
    sendLineMessageToContacts_(
      targets,
      text
    );

  ui.alert(
    [
      '送信が完了しました。',
      '',
      `対象：${targetLabel}`,
      `成功：${result.successCount}件`,
      `失敗：${result.errorCount}件`
    ].join('\n')
  );
}

/**
 * 対象者へLINE送信
 */
function sendLineMessageToContacts_(
  targets,
  text
) {
  let successCount = 0;
  let errorCount = 0;

  targets.forEach(contact => {
    try {
      pushMessages_(
        contact.userId,
        [
          {
            type: 'text',
            text
          }
        ],
        'CONTACT_BROADCAST_ERROR'
      );

      successCount++;

      Utilities.sleep(120);

    } catch (error) {
      errorCount++;

      errorLog_(
        'CONTACT_MESSAGE_SEND_ERROR',
        error.stack || error.message,
        contact.userId,
        contact.displayName || ''
      );
    }
  });

  return {
    successCount,
    errorCount
  };
}

/**
 * 送信対象取得
 *
 * V2移行（2026/10）：LINEメッセージの送信対象は、本番からコピーされた
 * 可能性のある連絡先シートの旧データ領域を含めず、必ずV2データ領域
 * （V2ヘッダーの次の行以降）のみから選定する。V2ヘッダーが見つからない
 * 場合は、安全側に倒して送信対象なし（0件）として扱う。
 */
function getTargetContactsByType_(type) {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(SHEET_CONTACTS);

  if (!sheet) {
    throw new Error(
      `シートが見つかりません: ${SHEET_CONTACTS}`
    );
  }

  const headerRow =
    resolveV2HeaderRow_(sheet, getContactHeaderDefinition_());

  if (headerRow === 0) {
    return [];
  }

  const dataStartRow =
    headerRow + 2;

  const lastRow =
    sheet.getLastRow();

  if (lastRow < dataStartRow) {
    return [];
  }

  const headers =
    sheet
      .getRange(
        headerRow,
        1,
        1,
        sheet.getLastColumn()
      )
      .getValues()[0]
      .map(h => String(h || '').trim());

  const values =
    sheet
      .getRange(
        dataStartRow,
        1,
        lastRow - dataStartRow + 1,
        sheet.getLastColumn()
      )
      .getValues();

  const userIdCol =
    headers.indexOf('ユーザーID');

  const displayNameCol =
    headers.indexOf('LINE表示名');

  const statusCol =
    headers.indexOf('状態');

  const applyStatusCol =
    headers.indexOf('応募状況');

  if (userIdCol === -1) {
    throw new Error(
      '連絡先シートにユーザーID列がありません。'
    );
  }

  return values
    .map(row => createContactTargetFromRow_(
      row,
      userIdCol,
      displayNameCol,
      statusCol,
      applyStatusCol
    ))
    .filter(contact => isTargetContactByType_(
      contact,
      type
    ));
}

/**
 * 行データから送信対象オブジェクト作成
 */
function createContactTargetFromRow_(
  row,
  userIdCol,
  displayNameCol,
  statusCol,
  applyStatusCol
) {
  return {
    userId:
      String(
        row[userIdCol] || ''
      ).trim(),

    displayName:
      displayNameCol >= 0
        ? String(
            row[displayNameCol] || ''
          ).trim()
        : '',

    status:
      statusCol >= 0
        ? String(
            row[statusCol] || ''
          ).trim()
        : '',

    applyStatus:
      applyStatusCol >= 0
        ? String(
            row[applyStatusCol] || ''
          ).trim()
        : ''
  };
}

/**
 * 送信対象判定
 */
function isTargetContactByType_(
  contact,
  type
) {
  if (!contact || !contact.userId) {
    return false;
  }

  switch (String(type)) {
    case '1':
      return contact.status === '友だち追加済み';

    case '2':
      return contact.status === 'やりとりあり';

    case '3':
      return contact.applyStatus === '応募中';

    case '4':
      return contact.applyStatus === '応募完了';

    case '5':
      return contact.applyStatus === '未応募';

    case '6':
      return true;

    default:
      return false;
  }
}

/**
 * 送信対象ラベル取得
 */
function getContactMessageTargetLabel_(type) {
  const labels = {
    '1': '友だち追加済み',
    '2': 'やりとりあり',
    '3': '応募中',
    '4': '応募完了',
    '5': '未応募',
    '6': '全員'
  };

  return labels[String(type)] || '';
}

function installUnappliedFollowupTrigger() {
  const triggers =
    ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (
      trigger.getHandlerFunction() ===
      'sendUnappliedFollowupMessages'
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp
    .newTrigger('sendUnappliedFollowupMessages')
    .timeBased()
    .everyDays(1)
    .atHour(13)
    .create();

  SpreadsheetApp
    .getUi()
    .alert('未応募者追客の自動送信トリガーを設定しました。毎日13時台に実行されます。');
}

/**
 * V2移行（2026/10）：追客（未応募者へのリマインド）の自動送信対象も、
 * 旧データ領域（本番からコピーされた可能性がある連絡先）には一切
 * 送信しない。必ずV2データ領域のみを対象にする。V2ヘッダーが見つから
 * ない場合は、安全側に倒して何もしない。
 */
function sendUnappliedFollowupMessages() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(SHEET_CONTACTS);

  if (!sheet) {
    throw new Error(`シートが見つかりません: ${SHEET_CONTACTS}`);
  }

  ensureContactHeader_();

  const headerRow =
    resolveV2HeaderRow_(sheet, getContactHeaderDefinition_());

  if (headerRow === 0) {
    return;
  }

  const dataStartRow =
    headerRow + 2;

  const lastRow =
    sheet.getLastRow();

  if (lastRow < dataStartRow) {
    return;
  }

  const headers =
    sheet
      .getRange(headerRow, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(h => String(h || '').trim());

  const values =
    sheet
      .getRange(dataStartRow, 1, lastRow - dataStartRow + 1, sheet.getLastColumn())
      .getValues();

  let successCount = 0;
  let skipCount = 0;
  let errorCount = 0;

  values.forEach((row, index) => {
    const rowNumber =
      index + dataStartRow;

    const target =
      buildFollowupTargetFromRow_(
        headers,
        row,
        rowNumber
      );

    const step =
      getUnappliedFollowupStep_(
        target
      );

    if (!step) {
      skipCount++;
      return;
    }

    try {
      pushMessages_(
        target.userId,
        [
          {
            type: 'text',
            text: getUnappliedFollowupMessageByStep_(step)
          }
        ],
        `UNAPPLIED_FOLLOWUP_${step}_ERROR`
      );

      updateContactFollowupStatus_(
        sheet,
        headers,
        rowNumber,
        getFollowupStatusByStep_(step)
      );

      successCount++;

      Utilities.sleep(120);

    } catch (error) {
      updateContactFollowupStatus_(
        sheet,
        headers,
        rowNumber,
        '送信失敗'
      );

      errorLog_(
        'UNAPPLIED_FOLLOWUP_SEND_ERROR',
        error.stack || error.message,
        target.userId || '',
        `row=${rowNumber}, step=${step}`
      );

      errorCount++;
    }
  });

  processLog_(
    'UNAPPLIED_FOLLOWUP_DONE',
    `成功=${successCount}, スキップ=${skipCount}, 失敗=${errorCount}`,
    '',
    ''
  );
}

function buildFollowupTargetFromRow_(
  headers,
  row,
  rowNumber
) {
  return {
    rowNumber,
    userId:
      getRowValueByHeaderName_(
        headers,
        row,
        'ユーザーID'
      ),
    displayName:
      getRowValueByHeaderName_(
        headers,
        row,
        'LINE表示名'
      ),
    followedAt:
      getRowValueByHeaderName_(
        headers,
        row,
        '友だち追加日時'
      ),
    status:
      getRowValueByHeaderName_(
        headers,
        row,
        '状態'
      ),
    applyStatus:
      getRowValueByHeaderName_(
        headers,
        row,
        '応募状況'
      ),
    followupStatus:
      getRowValueByHeaderName_(
        headers,
        row,
        '追客ステータス'
      )
  };
}

function getUnappliedFollowupStep_(
  target
) {
  if (!target || !target.userId) {
    return '';
  }

  if (
    String(target.applyStatus || '').trim() !== '未応募'
  ) {
    return '';
  }

  const status =
    String(target.status || '').trim();

  if (
    status === 'やりとりあり' ||
    status === '応募中' ||
    status === '応募完了'
  ) {
    return '';
  }

  const followupStatus =
    String(target.followupStatus || '').trim() || '未開始';

  if (
    followupStatus === '停止' ||
    followupStatus === '送信失敗' ||
    followupStatus === '7日送信済'
  ) {
    return '';
  }

  const elapsedDays =
    getElapsedDaysFromDate_(
      target.followedAt
    );

  if (
    elapsedDays >= 7 &&
    followupStatus === '5日送信済'
  ) {
    return '7日';
  }

  if (
    elapsedDays >= 5 &&
    followupStatus === '2日送信済'
  ) {
    return '5日';
  }

  if (
    elapsedDays >= 2 &&
    (
      followupStatus === '未開始' ||
      followupStatus === ''
    )
  ) {
    return '2日';
  }

  return '';
}

function getUnappliedFollowupMessageByStep_(
  step
) {
  if (step === '2日') {
    return [
      'ご登録ありがとうございます。',
      '',
      'ご応募をご希望の場合は、このLINEからそのまま応募を進められます。',
      '気になる点があれば、お気軽にご相談ください。'
    ].join('\n');
  }

  if (step === '5日') {
    return [
      'その後、ご応募についてご不明点はありませんか？',
      '',
      'ご応募をご希望の場合は、このLINEからそのまま応募を進められます。',
      '気になる点があれば、お気軽にご相談ください。'
    ].join('\n');
  }

  if (step === '7日') {
    return [
      'ご応募をご検討中でしたら、こちらのLINEからそのまま応募できます。',
      '',
      'タイミングが合う時に、お気軽にご連絡ください。'
    ].join('\n');
  }

  return '';
}

function getFollowupStatusByStep_(
  step
) {
  if (step === '2日') {
    return '2日送信済';
  }

  if (step === '5日') {
    return '5日送信済';
  }

  if (step === '7日') {
    return '7日送信済';
  }

  return '';
}

function updateContactFollowupStatus_(
  sheet,
  headers,
  row,
  status
) {
  if (!sheet || !headers || !row || !status) {
    return;
  }

  const followupCol =
    headers.indexOf('追客ステータス') + 1;

  const updatedAtCol =
    headers.indexOf('更新日') + 1;

  if (followupCol > 0) {
    sheet
      .getRange(row, followupCol)
      .setValue(status);
  }

  if (updatedAtCol > 0) {
    sheet
      .getRange(row, updatedAtCol)
      .setValue(new Date());
  }
}

function getElapsedDaysFromDate_(
  value
) {
  if (!value) {
    return 0;
  }

  const date =
    value instanceof Date
      ? value
      : new Date(value);

  if (isNaN(date.getTime())) {
    return 0;
  }

  const now =
    new Date();

  const diffMs =
    now.getTime() - date.getTime();

  return Math.floor(
    diffMs / (1000 * 60 * 60 * 24)
  );
}

function getRowValueByHeaderName_(
  headers,
  row,
  headerName
) {
  const index =
    headers.indexOf(headerName);

  if (index === -1) {
    return '';
  }

  return row[index];
}

