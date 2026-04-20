/**
 * SetupScript.gs - 初回セットアップ用スクリプト
 *
 * 使い方:
 *   1. 新規スプレッドシートを作成し、本プロジェクトを「拡張機能 → Apps Script」で開く
 *   2. すべての .gs / .html ファイルをこのプロジェクトにコピー（または clasp push）
 *   3. このファイルの setupAll() 関数を GAS エディタから実行
 *      （初回は権限承認のダイアログが出ます）
 *   4. 完了後、Settings シートを編集して各種値（Journal_Name、chiefEditorEmail 等）を設定
 *   5. ウェブアプリとしてデプロイ（公開: 全員 / 実行: 自分）
 *
 * 個別実行も可能:
 *   - setupSheets()     : 必要なシートを作成し、ヘッダー行を書き込む
 *   - setupSettings()   : Settings シートに既定値を投入
 *   - setupDecisions()  : Decisions シートに判定テンプレートのサンプルを投入
 *   - setupScriptProperty() : SPREADSHEET_ID をスクリプトプロパティに保存
 */

// ===== 各シートのヘッダー定義 =====

const SETUP_HEADERS = {
  Manuscripts: [
    'ID', 'MS_ID', 'MsVer', 'MsVerHex', 'key', 'eicKey',
    'CA_Name', 'CA_Email', 'CA_Affiliation', 'Address',
    'AuthorsJP', 'AuthorsEN', 'ccEmails',
    'MS_Type', 'TitleJP', 'TitleEN', 'RunningTitle',
    'AbstractJP', 'AbstractEN', 'LetterToEditor',
    'English_editing', 'Reprint request',
    'submittedFiles', 'folderUrl', 'receiptFolderUrl',
    'Ver_No', 'Submitted_At',
    // 判定・受理関連（FeedbackModule / FinalReviewModule で書き込み）
    'score', 'openComments', 'accepted', 'sentBackAt', 'resultFolderUrl',
    // 最終確認ルート（FinalReviewModule で書き込み）
    'managingEditorAuthorComment', 'managingEditorInternalComment',
    'managingEditorFileUrl', 'managingEditorSentAt',
    'eicFinalComment', 'eicProductionComment',
    'eicFinalFileUrl', 'eicFinalDecision',
    'finalStatus', 'isAccepted'
  ],
  Editor_log: [
    'MsVer', 'MsVerRevHex', 'editorKey',
    'Editor_Name', 'Editor_Email',
    'Ask_At', 'Answer_At', 'edtOk', 'Message',
    'Score', 'ConfidentialMessage',
    'reportPdfUrl', 'reportWordUrl', 'reportFolderUrl',
    'reportGoogleDocId', 'reportCommentPdfUrl',
    'reportAttachmentsFolderUrl',
    'firstReminded', 'secondReminded', 'thirdReminded'
  ],
  Review_log: [
    'MsVerRevHex', 'MsVer', 'reviewKey',
    'Editor_Name', 'Editor_Email',
    'Rev_Name', 'Rev_Email',
    'Ask_At', 'Answer_At', 'Received_At',
    'revOk', 'Review_Deadline', 'Score', 'Message', 'ConfidentialMessage',
    'openCommentsId', 'confidentialCommentsId',
    'reviewerUploadFolderUrl', 'reviewMaterialsFolderUrl',
    'firstReminded', 'secondReminded', 'thirdReminded'
  ],
  Emails: [
    'to', 'cc', 'bcc', 'subject', 'htmlBody',
    'attachmentFileIds', 'savedAt', 'logText'
  ],
  Log: ['Timestamp', 'Message'],
  Log_archive: ['Timestamp', 'Message']
};

// ===== Settings シートの既定値 =====

const SETUP_SETTINGS_DEFAULTS = [
  ['Journal_Name',          'My Journal'],
  ['Editor_Name',           'Editor-in-Chief'],
  ['chiefEditorEmail',      ''],
  ['managingEditorEmail',   ''],
  ['productionEditorEmail', ''],
  ['eicAdminKey',           Utilities.getUuid().replace(/-/g, '')],
  ['SUBFOLDER',             'Journal Files'],
  ['reviewMaterialsFolder', ''],
  ['submissionBccEmails',   ''],
  ['Resubmittion_Limit',    '8 weeks'],
  ['Review_Period',         '21'],
  ['firstReminderDays',     '7'],
  ['secondReminderDays',    '14'],
  ['thirdReminderDays',     '21'],
  ['mailFooter',            '<hr><p style="font-size:12px;color:#666;">This is an automated message. / これは自動配信メールです。</p>']
];

const SETUP_MANUSCRIPT_TYPES_DEFAULTS = [
  ['Original Article',  'O'],
  ['Review Article',    'R'],
  ['Short Communication', 'S'],
  ['Letter',            'L']
];

// ===== Decisions シートの既定テンプレート =====

const SETUP_DECISIONS_DEFAULTS = {
  headers: ['ShortExplanation', 'IsAccepted', 'Resubmit', 'Mail text'],
  rows: [
    ['Accept', 'yes', 'no',
     'Dear {{authorName}},\n\nWe are pleased to inform you that your manuscript "{{englishTitle}}" ({{manuscriptID}}) has been accepted for publication in {{Journal_Name}}.\n\nSincerely,\n{{Editor_Name}}'],
    ['Accept with Resubmission', 'yes', 'yes',
     'Dear {{authorName}},\n\nYour manuscript "{{englishTitle}}" ({{manuscriptID}}) has been provisionally accepted for {{Journal_Name}}, subject to minor revisions.\n\nPlease submit the revised version by {{dueDate}} via:\n{{formlink}}\n\nSincerely,\n{{Editor_Name}}'],
    ['Major Revision', 'no', 'yes',
     'Dear {{authorName}},\n\nYour manuscript "{{englishTitle}}" ({{manuscriptID}}) requires major revisions before it can be considered for publication in {{Journal_Name}}.\n\nPlease submit the revised version within {{Resubmittion_Limit}} via:\n{{formlink}}\n\nSincerely,\n{{Editor_Name}}'],
    ['Minor Revision', 'no', 'yes',
     'Dear {{authorName}},\n\nYour manuscript "{{englishTitle}}" ({{manuscriptID}}) requires minor revisions.\n\nPlease submit the revised version within {{Resubmittion_Limit}} via:\n{{formlink}}\n\nSincerely,\n{{Editor_Name}}'],
    ['Reject', 'no', 'no',
     'Dear {{authorName}},\n\nWe regret to inform you that your manuscript "{{englishTitle}}" ({{manuscriptID}}) is not suitable for publication in {{Journal_Name}}.\n\nSincerely,\n{{Editor_Name}}']
  ]
};

// ===== メインのセットアップ関数 =====

/**
 * 全てのセットアップ処理をまとめて実行する。
 * GAS エディタから 1 度だけ実行する想定。
 */
function setupAll() {
  const messages = [];
  messages.push('=== セットアップ開始 / Setup started ===');

  try {
    messages.push(setupScriptProperty());
    messages.push(setupSheets());
    messages.push(setupSettings());
    messages.push(setupManuscriptTypes());
    messages.push(setupDecisions());
    messages.push('');
    messages.push('=== セットアップ完了 / Setup complete ===');
    messages.push('次のステップ / Next steps:');
    messages.push('  1. Settings シートで Journal_Name、chiefEditorEmail などを設定してください。');
    messages.push('  2. ウェブアプリとしてデプロイしてください（公開: 全員 / 実行: 自分）。');
    messages.push('  3. デプロイ URL を著者・編集者に共有してください。');
  } catch (err) {
    messages.push('[ERROR] ' + err.message);
    messages.push(err.stack || '');
  }

  const result = messages.join('\n');
  Logger.log(result);
  try { SpreadsheetApp.getUi().alert(result); } catch (_) { /* UI 不可時は無視 */ }
  return result;
}

/**
 * 現在開いているスプレッドシートの ID をスクリプトプロパティに保存する。
 */
function setupScriptProperty() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('アクティブなスプレッドシートがありません。スプレッドシートにバインドされたスクリプトとして実行してください。 / No active spreadsheet. Run as a bound script.');
  }
  const ssId = ss.getId();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ssId);
  return '[OK] SPREADSHEET_ID をスクリプトプロパティに保存しました: ' + ssId;
}

/**
 * 必要な全シートを作成し、ヘッダー行を書き込む。既存のシートはそのまま残す。
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lines = [];

  Object.keys(SETUP_HEADERS).forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      lines.push('  [NEW] ' + sheetName + ' シートを作成しました。');
    } else {
      lines.push('  [EXISTS] ' + sheetName + ' シートは既に存在します。');
    }

    // ヘッダー行が空（または欠けている）場合のみヘッダーを書き込む
    const headers = SETUP_HEADERS[sheetName];
    const lastCol = sheet.getLastColumn();
    const existingHeaders = lastCol > 0
      ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim())
      : [];

    const missing = headers.filter(h => !existingHeaders.some(eh => eh.toLowerCase() === h.toLowerCase()));

    if (existingHeaders.filter(Boolean).length === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f1f5f9');
      sheet.setFrozenRows(1);
      lines.push('    → ヘッダー行を書き込みました（' + headers.length + ' 列）。');
    } else if (missing.length > 0) {
      let nextCol = lastCol + 1;
      missing.forEach(h => {
        sheet.getRange(1, nextCol).setValue(h).setFontWeight('bold').setBackground('#f1f5f9');
        nextCol++;
      });
      lines.push('    → 不足ヘッダー ' + missing.length + ' 列を追加しました: ' + missing.join(', '));
    } else {
      lines.push('    → ヘッダー行は既に揃っています。');
    }
  });

  // Settings シートも作成（ヘッダー扱いではないので別途）
  if (!ss.getSheetByName(SETTINGS_SHEET_NAME)) {
    const settingsSheet = ss.insertSheet(SETTINGS_SHEET_NAME);
    settingsSheet.getRange('A1').setValue('Settings').setFontWeight('bold').setFontSize(14);
    settingsSheet.getRange('A3').setValue('Key').setFontWeight('bold').setBackground('#f1f5f9');
    settingsSheet.getRange('B3').setValue('Value').setFontWeight('bold').setBackground('#f1f5f9');
    settingsSheet.getRange('E3').setValue('MS Type Label').setFontWeight('bold').setBackground('#f1f5f9');
    settingsSheet.getRange('F3').setValue('Prefix').setFontWeight('bold').setBackground('#f1f5f9');
    settingsSheet.setFrozenRows(3);
    lines.push('  [NEW] Settings シートを作成しました。');
  } else {
    lines.push('  [EXISTS] Settings シートは既に存在します。');
  }

  // Decisions シート
  if (!ss.getSheetByName(DECISION_MAIL_SHEET_NAME)) {
    const decSheet = ss.insertSheet(DECISION_MAIL_SHEET_NAME);
    decSheet.getRange('A1').setValue('Decisions / 判定テンプレート').setFontWeight('bold').setFontSize(14);
    lines.push('  [NEW] Decisions シートを作成しました。');
  } else {
    lines.push('  [EXISTS] Decisions シートは既に存在します。');
  }

  return '[OK] シートのセットアップ完了:\n' + lines.join('\n');
}

/**
 * Settings シートに既定値を書き込む（既存の値は上書きしない）。
 */
function setupSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) throw new Error('Settings シートが存在しません。先に setupSheets() を実行してください。');

  // 既存のキー一覧（A5:A100）を取得
  const existingKeys = sheet.getRange('A5:A100').getValues()
    .map(row => String(row[0] || '').trim().toLowerCase())
    .filter(Boolean);

  let nextRow = 5 + existingKeys.length;
  let added = 0;

  SETUP_SETTINGS_DEFAULTS.forEach(([key, value]) => {
    if (existingKeys.includes(key.toLowerCase())) return; // 既存はスキップ
    sheet.getRange(nextRow, 1).setValue(key);
    sheet.getRange(nextRow, 2).setValue(value);
    nextRow++;
    added++;
  });

  return '[OK] Settings シート: ' + added + ' 件の既定値を追加しました（既存値は保持）。';
}

/**
 * 原稿種別（E5:F18）に既定値を書き込む。
 */
function setupManuscriptTypes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) throw new Error('Settings シートが存在しません。');

  const existingTypes = sheet.getRange('E5:E18').getValues()
    .map(row => String(row[0] || '').trim().toLowerCase())
    .filter(Boolean);

  let nextRow = 5 + existingTypes.length;
  let added = 0;

  SETUP_MANUSCRIPT_TYPES_DEFAULTS.forEach(([label, prefix]) => {
    if (existingTypes.includes(label.toLowerCase())) return;
    if (nextRow > 18) return;
    sheet.getRange(nextRow, 5).setValue(label);
    sheet.getRange(nextRow, 6).setValue(prefix);
    nextRow++;
    added++;
  });

  return '[OK] 原稿種別: ' + added + ' 件の既定値を追加しました。';
}

/**
 * Decisions シートに既定の判定テンプレートを投入する。
 * 既にデータがある場合はスキップする（上書きしない）。
 */
function setupDecisions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DECISION_MAIL_SHEET_NAME);
  if (!sheet) throw new Error('Decisions シートが存在しません。');

  // 既にヘッダー行（ShortExplanation を含む）があるか確認
  const data = sheet.getDataRange().getValues();
  const hasHeader = data.some(row =>
    row.some(cell => String(cell).toLowerCase().trim() === 'shortexplanation')
  );
  if (hasHeader) {
    return '[SKIP] Decisions シートには既にデータがあります（既存データ保持）。';
  }

  // ヘッダー行を 3 行目に書き込む（A1 にはタイトルがある想定）
  sheet.getRange(3, 1, 1, SETUP_DECISIONS_DEFAULTS.headers.length)
    .setValues([SETUP_DECISIONS_DEFAULTS.headers])
    .setFontWeight('bold').setBackground('#f1f5f9');
  sheet.setFrozenRows(3);

  // データ行
  const rows = SETUP_DECISIONS_DEFAULTS.rows;
  sheet.getRange(4, 1, rows.length, SETUP_DECISIONS_DEFAULTS.headers.length).setValues(rows);

  // Mail text 列を折り返し表示にする
  sheet.getRange(4, 4, rows.length, 1).setWrap(true);
  sheet.setColumnWidth(4, 600);

  return '[OK] Decisions シート: ' + rows.length + ' 件のテンプレートを投入しました。';
}

/**
 * スプレッドシートを開いた時にカスタムメニューを追加する。
 * onOpen はシンプル（unauth）なので setup ではなく案内のみを行う。
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Manuscript System')
      .addItem('初回セットアップ / Setup All', 'setupAll')
      .addSeparator()
      .addItem('シートのみ作成 / Setup Sheets', 'setupSheets')
      .addItem('既定値投入 / Setup Settings', 'setupSettings')
      .addItem('原稿種別投入 / Setup MS Types', 'setupManuscriptTypes')
      .addItem('判定テンプレート投入 / Setup Decisions', 'setupDecisions')
      .addSeparator()
      .addItem('Decisions 診断 / Diagnose Decisions', 'diagnoseDecisionsSheet')
      .addToUi();
  } catch (e) {
    // バインドされていない場合などは黙って無視
  }
}
