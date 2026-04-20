/**
 * Code.gs - 統一論文投稿システム バックエンドロジック
 * 役割: ルーティング、データ取得、ビジネスロジックの統合
 */

// グローバル定数・変数
const SETTINGS_SHEET_NAME = 'Settings';
const MANUSCRIPTS_SHEET_NAME = 'Manuscripts';
const EDITOR_LOG_SHEET_NAME = 'Editor_log';
const REVIEW_LOG_SHEET_NAME = 'Review_log';
const PENDING_EMAILS_SHEET_NAME = 'Emails';
const DECISION_MAIL_SHEET_NAME = 'Decisions';
const LOG_SHEET_NAME = 'Log';
const ARCHIVE_SHEET_NAME = 'Archive';
const LOG_ARCHIVE_SHEET_NAME = 'Log_archive';

// リファクタリングされたモジュールの参照
// Note: GASではモジュールインポートができないため、関数はグローバルスコープで定義されます
// ManuscriptDataHandlers.js の関数は別ファイルで定義され、GASによって自動的に読み込まれます

/**
 * ウェブアプリのエントリポイント
 * URLパラメータ: key, editorKey, reviewKey に基づき画面を切り替える
 */
function doGet(e) {
  // デバッグ: すべてのパラメータを表示
  console.log('All URL parameters:', JSON.stringify(e.parameter));
  
  const key = e.parameter.key || '';
  const eicKey = e.parameter.eicKey || '';
  const editorKey = e.parameter.editorKey || '';
  const reviewKey = e.parameter.reviewKey || '';
  const managingEditorKey = e.parameter.managingEditorKey || '';
  const testParam = e.parameter.test || '';
  
  console.log('test parameter:', testParam);

  const settings = getSettings();
  const journalName = settings.Journal_Name || 'Journal Submission System';

  const template = HtmlService.createTemplateFromFile('index');
  template.journalName = journalName;

  // テンプレート変数の初期化
  template.role = 'unknown';
  template.key = '';
  template.showResponseForm = 'false';
  template.initialMsData = 'null';
  try { template.initialSettings = JSON.stringify(settings).replace(/<\//g, '<\\/'); } catch(e) { template.initialSettings = '{}'; }
  template.webAppUrl = ScriptApp.getService().getUrl();

  // テストモード判定（デバッグ情報付き）
  console.log('Test parameter value:', testParam, 'type:', typeof testParam);
  template.isTestMode = (testParam === 'true') ? 'true' : 'false';
  console.log('isTestMode template variable:', template.isTestMode);

  // 査読アコーディオンを初期展開するか（査読完了通知メールからのリンクで使用）
  template.openReviews = (e.parameter.openReviews === '1') ? 'true' : 'false';

  // ロールの判定
  try {
    if (reviewKey) {
      const k = reviewKey;
      template.role = 'Reviewer';
      template.key = k;
      const msData = getManuscriptData('reviewer', k);
      try { template.initialMsData = JSON.stringify(msData || null).replace(/<\//g, '<\\/'); } catch(e) {}
      if (msData && (!msData.revOk || msData.revOk === '')) {
        template.showResponseForm = 'true';
      }
    } else if (editorKey) {
      template.role = 'Editor';
      template.key = editorKey;
      const msData = getManuscriptData('editor', editorKey);
      try { template.initialMsData = JSON.stringify(msData || null).replace(/<\//g, '<\\/'); } catch(e) {}
      if (msData && (!msData.edtOk || msData.edtOk === '')) {
        template.showResponseForm = 'true';
      }
    } else if (managingEditorKey) {
      // 編集幹事専用キー
      const msData = getManuscriptData('managing-editor', managingEditorKey);
      if (msData) {
        template.role = 'Managing-Editor';
        template.key = managingEditorKey;
        try { template.initialMsData = JSON.stringify(msData).replace(/<\//g, '<\\/'); } catch(e) {}
      } else {
        template.role = 'unknown';
      }
    } else if (eicKey) {
      // eicAdminKey と一致する場合は全体進捗一覧ページへ
      if (settings.eicAdminKey && eicKey === String(settings.eicAdminKey).trim()) {
        template.role = 'Eic-Overview';
        template.key  = eicKey;
        try {
          const ssIdForOverview = getSpreadsheetId();
          const allMs = getEicAllMsData(ssIdForOverview);
          template.initialMsData = JSON.stringify(allMs).replace(/<\//g, '<\\/');
        } catch(eOverview) {
          console.error('getEicAllMsData error:', eOverview);
          template.initialMsData = '[]';
        }
      } else {
        // 編集委員長専用キー：著者キーと完全に分離されたキーでロールをサーバー側で検証
        const msData = getManuscriptData('eic', eicKey);
        if (msData) {
          template.role = 'Editor-in-chief';
          template.key = eicKey;
          try { template.initialMsData = JSON.stringify(msData).replace(/<\//g, '<\\/'); } catch(e) {}
        } else {
          template.role = 'unknown';
        }
      }
    } else if (key) {
      // 著者キーは常に著者ロール（サーバー側で確証済み）
      template.role = 'Author';
      template.key = key;
      try { template.initialMsData = JSON.stringify(getManuscriptData('author', key) || null).replace(/<\//g, '<\\/'); } catch(e) {}
    } else {
      template.role = 'New-Author';
    }
  } catch (err) {
    console.error('doGet role detection error:', err);
    template.role = 'unknown';
  }
  
  return template.evaluate()
    .setTitle(journalName)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * index.html内でのインクルード用ユーティリティ
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * スプレッドシートIDを取得（キャッシュ付き最適化版）
 * 1) スクリプトプロパティ SPREADSHEET_ID → 2) fassSetting.json の順で取得
 */
function getSpreadsheetId() {
  return getSpreadsheetIdOptimized();
}


/**
 * Settingsシートから全般設定を取得
 */
function getSettings() {
  const ssId = getSpreadsheetId();
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(SETTINGS_SHEET_NAME);
  const data = sheet.getRange('A5:B100').getValues();
  const settings = {};
  
  data.forEach(row => {
    if (row[0]) {
      settings[row[0]] = row[1];
    }
  });
  
  return settings;
}

/**
 * Decisions シートを名前で取得する。
 * DECISION_MAIL_SHEET_NAME で見つからない場合、よくある別名もフォールバックで試す。
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function _getDecisionsSheet(ssId) {
  const ss = SpreadsheetApp.openById(ssId);
  const candidates = [DECISION_MAIL_SHEET_NAME, 'Decisions', 'DecisionMail', 'Decision'];
  for (const name of candidates) {
    const sheet = ss.getSheetByName(name);
    if (sheet) return sheet;
  }
  // 全シートを走査して ShortExplanation ヘッダーを含むシートを探す（最終手段）
  const allSheets = ss.getSheets();
  for (const sheet of allSheets) {
    const parsed = _findDecisionsSheetRows(sheet);
    if (parsed) {
      writeLog('[INFO] Decisions シートを自動検出しました: "' + sheet.getName() + '"');
      return sheet;
    }
  }
  return null;
}

/**
 * Decisionsシートの全データから ShortExplanation ヘッダー行を動的に探して返す
 * @returns {{ headerRowIdx: number, headers: string[], data: any[][] } | null}
 */
function _findDecisionsSheetRows(sheet) {
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  const headerRowIdx = data.findIndex(row =>
    row.some(cell => String(cell).toLowerCase().trim() === 'shortexplanation')
  );
  if (headerRowIdx === -1) return null;
  const headers = data[headerRowIdx].map(h => String(h).trim());
  return { headerRowIdx, headers, data };
}

/**
 * Decisions シートの設定状態を診断する。
 * GAS エディタから手動実行して Logger.log の出力を確認してください。
 */
function diagnoseDecisionsSheet() {
  const ssId = getSpreadsheetId();
  const ss = SpreadsheetApp.openById(ssId);
  const lines = [];
  lines.push('=== Decisions シート診断 ===');
  lines.push('DECISION_MAIL_SHEET_NAME = "' + DECISION_MAIL_SHEET_NAME + '"');

  // 1. シート検索
  const sheet = _getDecisionsSheet(ssId);
  if (!sheet) {
    lines.push('[NG] Decisions シートが見つかりません。');
    lines.push('スプレッドシート内のシート一覧:');
    ss.getSheets().forEach(s => lines.push('  - "' + s.getName() + '"'));
    const msg = lines.join('\n');
    Logger.log(msg);
    writeLog(msg);
    return msg;
  }
  lines.push('[OK] シート発見: "' + sheet.getName() + '"');

  // 2. ヘッダー行検索
  const parsed = _findDecisionsSheetRows(sheet);
  if (!parsed) {
    lines.push('[NG] ShortExplanation ヘッダー行が見つかりません。');
    const msg = lines.join('\n');
    Logger.log(msg);
    writeLog(msg);
    return msg;
  }
  lines.push('[OK] ヘッダー行: ' + (parsed.headerRowIdx + 1) + '行目');
  lines.push('     ヘッダー列: ' + JSON.stringify(parsed.headers.filter(h => h !== '')));

  // 3. 列チェック
  const findHeaderIdx = (keyword) => {
    const kw = keyword.toLowerCase();
    return parsed.headers.findIndex(h => String(h).toLowerCase().trim() === kw);
  };
  const sIdx = findHeaderIdx('ShortExplanation');
  const aIdx = findHeaderIdx('IsAccepted');
  const rIdx = findHeaderIdx('Resubmit');
  const tIdx = findHeaderIdx('Mail text');
  lines.push(sIdx !== -1 ? '[OK] ShortExplanation 列あり' : '[NG] ShortExplanation 列なし');
  lines.push(aIdx !== -1 ? '[OK] IsAccepted 列あり' : '[NG] IsAccepted 列なし');
  lines.push(rIdx !== -1 ? '[OK] Resubmit 列あり' : '[NG] Resubmit 列なし');
  lines.push(tIdx !== -1 ? '[OK] Mail text 列あり' : '[NG] Mail text 列なし');

  // 4. データ行
  const dataRows = parsed.data.slice(parsed.headerRowIdx + 1).filter(r => String(r[sIdx] || '').trim() !== '');
  lines.push('データ行数: ' + dataRows.length);
  dataRows.forEach((r, i) => {
    const se = String(r[sIdx] || '').trim();
    const ia = aIdx !== -1 ? String(r[aIdx] || '').trim() : '(列なし)';
    const re = rIdx !== -1 ? String(r[rIdx] || '').trim() : '(列なし)';
    lines.push('  [' + (i + 1) + '] ShortExplanation="' + se + '" IsAccepted="' + ia + '" Resubmit="' + re + '"');
  });

  // 5. getScoreOptions と一致するか確認
  const scoreOptions = getScoreOptions();
  lines.push('');
  lines.push('getScoreOptions() の結果: ' + JSON.stringify(scoreOptions));
  const sheetValues = dataRows.map(r => String(r[sIdx] || '').trim());
  const match = JSON.stringify(scoreOptions) === JSON.stringify(sheetValues);
  lines.push(match ? '[OK] ドロップダウン値とシート値が一致' : '[WARN] ドロップダウン値とシート値が不一致 → isScoreAccepted が false になる原因');

  const msg = lines.join('\n');
  Logger.log(msg);
  writeLog(msg);
  return msg;
}

/**
 * Decisionsシートの ShortExplanation 列からスコア選択肢を取得
 * 空欄・シート未存在の場合はデフォルト値にフォールバック
 */
function getScoreOptions() {
  try {
    const ssId = getSpreadsheetId();

    // spreadsheetCache に Decisions シートが展開済みなら API 呼び出し不要
    const cachedData = spreadsheetCache.getSheetData(ssId, DECISION_MAIL_SHEET_NAME);
    if (cachedData) {
      const allRows  = [cachedData.headers, ...cachedData.rows];
      const parsed   = _findDecisionsSheetRowsFromArray(allRows);
      if (parsed) {
        const { headerRowIdx, headers, data } = parsed;
        const sIdx = headers.findIndex(h => String(h).toLowerCase().trim() === 'shortexplanation');
        if (sIdx !== -1) {
          const options = data.slice(headerRowIdx + 1)
            .map(row => String(row[sIdx]).trim())
            .filter(v => v !== '');
          if (options.length > 0) return options;
        }
      }
    }

    // キャッシュになければ従来通りシートから直接取得
    const sheet = _getDecisionsSheet(ssId);
    if (!sheet) throw new Error('Decisions sheet not found');
    const parsed = _findDecisionsSheetRows(sheet);
    if (!parsed) throw new Error('ShortExplanation header not found in Decisions sheet');
    const { headerRowIdx, headers, data } = parsed;
    const findHeaderIdx = (keyword) => {
      const kw = keyword.toLowerCase();
      return headers.findIndex(h => String(h).toLowerCase().trim() === kw);
    };
    const sIdx = findHeaderIdx('ShortExplanation');
    if (sIdx === -1) throw new Error('ShortExplanation column could not be found');
    const options = data.slice(headerRowIdx + 1)
      .map(row => String(row[sIdx]).trim())
      .filter(v => v !== '');
    return options.length > 0 ? options : ['Accept', 'Minor Revision', 'Major Revision', 'Reject'];
  } catch(e) {
    Logger.log('getScoreOptions fallback: ' + e.message);
    return ['Accept', 'Minor Revision', 'Major Revision', 'Reject'];
  }
}

/**
 * オープンコメント Google Doc のコピー先頭にヘッダーブロックを挿入する。
 *
 * 呼び出し元は必ず「原本のコピー」の Body を渡すこと。
 * 原本そのものへの書き込みを避けることで、EIC が PDF 送信後も
 * 元の Google Doc を再編集できる状態を保つ。
 *
 * 挿入順（上から）: 雑誌名 → 発行日時 → 水平線 → 判定
 *
 * @param {GoogleAppsScript.Document.Body} body      コピー先 Document の Body
 * @param {string} journalName  雑誌名
 * @param {string} score        EIC の判定スコア
 * @param {string} dateStr      発行日時文字列
 */
function insertCommentDocHeader(body, journalName, score, dateStr) {
  var styleTitle = {};
  styleTitle[DocumentApp.Attribute.FONT_SIZE]        = 12;
  styleTitle[DocumentApp.Attribute.BOLD]             = true;
  styleTitle[DocumentApp.Attribute.FONT_FAMILY]      = 'Arial';

  var styleSmall = {};
  styleSmall[DocumentApp.Attribute.FONT_SIZE]        = 10;
  styleSmall[DocumentApp.Attribute.BOLD]             = false;
  styleSmall[DocumentApp.Attribute.FONT_FAMILY]      = 'Arial';
  styleSmall[DocumentApp.Attribute.FOREGROUND_COLOR] = '#555555';

  var styleLabel = {};
  styleLabel[DocumentApp.Attribute.FONT_SIZE]        = 10.5;
  styleLabel[DocumentApp.Attribute.BOLD]             = true;
  styleLabel[DocumentApp.Attribute.FONT_FAMILY]      = 'Arial';

  // インデックス 0 に逆順で挿入することで、最終的に上から
  // [雑誌名] → [日付] → [HR] → [判定] → [元のコンテンツ] の順になる
  var decPara = body.insertParagraph(0, '判定 / Decision: ' + (score || ''));
  decPara.setAttributes(styleLabel);

  body.insertHorizontalRule(0);

  var datePara = body.insertParagraph(0, dateStr || '');
  datePara.setAttributes(styleSmall);
  datePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  var titlePara = body.insertParagraph(0, journalName || '');
  titlePara.setAttributes(styleTitle);
  titlePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
}

/**
 * キャッシュ済み 2D 配列から _findDecisionsSheetRows 相当の処理を行うヘルパー
 */
function _findDecisionsSheetRowsFromArray(data) {
  if (!data || data.length === 0) return null;
  const headerRowIdx = data.findIndex(row =>
    row.some(cell => String(cell).toLowerCase().trim() === 'shortexplanation')
  );
  if (headerRowIdx === -1) return null;
  return {
    headerRowIdx,
    headers: data[headerRowIdx].map(h => String(h).trim()),
    data
  };
}

/**
 * Settingsシートから原稿種別と接頭辞を取得 (E5:F18)
 */
function getManuscriptTypes() {
  const ssId = getSpreadsheetId();
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(SETTINGS_SHEET_NAME);
  const data = sheet.getRange('E5:F18').getValues();
  const types = [];
  
  data.forEach(row => {
    if (row[0]) {
      types.push({
        label: row[0],
        prefix: row[1]
      });
    }
  });
  
  return types;
}

/**
 * Google Doc の本文テキストを取得するヘルパー
 * ID が空・無効の場合は空文字を返す
 */
function readDocText(docId) {
  if (!docId || docId === 'nofile' || docId === '') return '';
  try {
    return DocumentApp.openById(docId).getBody().getText().trim();
  } catch(e) {
    Logger.log('readDocText error (' + docId + '): ' + e);
    return '';
  }
}

/**
 * デバッグ用: シートのヘッダー一覧を返す
 */
function getSheetHeaders(sheetName) {
  const ssId = getSpreadsheetId();
  if (!ssId) return ['ERROR: No spreadsheet ID'];
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(sheetName || MANUSCRIPTS_SHEET_NAME);
  if (!sheet) return ['ERROR: Sheet not found - ' + sheetName];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

/**
 * 論文データの取得 (keyに基づきManuscripts, Editor_log, Review_logから統合)
 * リファクタリング版: ロール別ハンドラーに分割
 */
function getManuscriptData(role, key) {
  // リファクタリング版の関数を呼び出す
  return getManuscriptDataRefactored(role, key);
}

/**
 * キーに一致する全行をオブジェクトの配列として返す汎用関数
 * 最適化版: キャッシュを使用し、API呼び出しを最小化
 */
function findAllRecordsByKey(ssId, sheetName, keyColName, keyValue) {
  return findAllRecordsByKeyOptimized(ssId, sheetName, keyColName, keyValue);
}

/**
 * キーに一致する行をオブジェクトとして返す汎用関数
 * 最適化版: キャッシュを使用し、API呼び出しを最小化
 */
function findRecordByKey(ssId, sheetName, keyColName, keyValue) {
  return findRecordByKeyOptimized(ssId, sheetName, keyColName, keyValue);
}

/**
 * 承諾・辞退リンクのクリックを処理
 */
/**
 * 担当編集者・査読者の受諾/辞退回答を処理するAPI
 */
function apiSubmitInvitationResponse(role, key, answer, message) {
  const ssId = getSpreadsheetId();
  const settings = getSettings();
  
  const r = (role || '').toLowerCase();
  
  if (r === 'editor') {
    const editorLog = findRecordByKey(ssId, EDITOR_LOG_SHEET_NAME, 'editorKey', key);
    if (!editorLog) throw new Error('Editor record not found.');

    // 既に辞退・取消・受諾済みの招待に対する二重回答を防ぐ。
    // （同じメールアドレスに新旧2通の招待が届いた場合など、
    //   古いURLから再回答されると誤った行に書き込まれる事故を防止）
    const existingEdtOk = String(editorLog.edtOk || '').trim();
    if (existingEdtOk === 'ng' || existingEdtOk === 'cancelled') {
      throw new Error('この招待はすでに辞退済みまたは取消済みのため、回答を受け付けできません。'
                    + '最新の依頼メールに記載のURLをご確認ください。\n'
                    + 'This invitation has already been declined or cancelled. '
                    + 'Please use the URL in the most recent invitation email.');
    }
    if (existingEdtOk === 'ok') {
      throw new Error('この招待はすでに受諾済みのため、再度の回答は受け付けできません。\n'
                    + 'This invitation has already been accepted.');
    }

    // 1. Logに記録
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
    updateLogCell(ssId, EDITOR_LOG_SHEET_NAME, 'editorKey', key, {
      'edtOk': answer,
      'Answer_At': now,
      'Message': message
    });
    
    writeLog(`Editor Response: ${editorLog.MsVer} by ${editorLog.Editor_Email} - ${answer}`);
    
    // 2. 通知送信
    const { msId } = parseMsVer(editorLog.MsVer);
    const ms = findRecordByKey(ssId, MANUSCRIPTS_SHEET_NAME, 'MS_ID', msId);
    
    // 委員長へ通知（受諾・辞退いずれも委員長ダッシュボードへのリンクを付与）
    const eicActionUrl = (ms && ms.eicKey)
      ? ScriptApp.getService().getUrl() + '?eicKey=' + ms.eicKey
      : null;
    const eicActionLabel = answer === 'ok'
      ? 'Open EIC Dashboard / 編集委員長ダッシュボードを開く'
      : 'Assign New Editor / 新しい候補者を指名する';
    const eicNoteHtml = answer === 'ok' ? `
      <p style="margin-top:1.5rem;">Please open your EIC dashboard using the button below to check the manuscript status.</p>
      <hr style="border:none; border-top:1px solid #e2e8f0; margin: 20px 0;">
      <p>以下のボタンより編集委員長ダッシュボードを開き、原稿の状況をご確認ください。</p>
    ` : null;
    if (settings.chiefEditorEmail) {
      sendAssignmentResponseNotificationToRequester(settings.chiefEditorEmail, 'Editor-in-Chief', editorLog.Editor_Name, ms || {}, answer, message, settings, eicActionUrl, eicActionLabel, eicNoteHtml);
    } else {
      Logger.log('apiSubmitInvitationResponse: chiefEditorEmail が設定されていないため EIC への通知をスキップします');
    }
    
    // 3. 受諾した場合はWelcomeメール送信
    if (answer === 'ok') {
      sendEditorWelcomeEmail(editorLog, settings);
    }
    
  } else if (r === 'reviewer') {
    const reviewLog = findRecordByKey(ssId, REVIEW_LOG_SHEET_NAME, 'reviewKey', key);
    if (!reviewLog) throw new Error('Reviewer record not found.');

    // 既に辞退・取消・受諾済みの招待に対する二重回答を防ぐ。
    // （同じメールアドレスに新旧2通の招待が届いた場合など、
    //   古いURLから再回答されると誤った行に書き込まれる事故を防止）
    const existingRevOk = String(reviewLog.revOk || '').trim();
    if (existingRevOk === 'ng' || existingRevOk === 'cancelled') {
      throw new Error('この招待はすでに辞退済みまたは取消済みのため、回答を受け付けできません。'
                    + '最新の依頼メールに記載のURLをご確認ください。\n'
                    + 'This invitation has already been declined or cancelled. '
                    + 'Please use the URL in the most recent invitation email.');
    }
    if (existingRevOk === 'ok') {
      throw new Error('この招待はすでに受諾済みのため、再度の回答は受け付けできません。\n'
                    + 'This invitation has already been accepted.');
    }

    // 1. Logに記録
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
    updateLogCell(ssId, REVIEW_LOG_SHEET_NAME, 'reviewKey', key, {
      'revOk': answer,
      'Answer_At': now,
      'Message': message
    });
    
    writeLog(`Reviewer Response: ${reviewLog.MsVer} by ${reviewLog.Rev_Email} - ${answer}`);
    
    // 2. 通知送信
    const { msId: reviewMsId } = parseMsVer(reviewLog.MsVer);
    const ms = findRecordByKey(ssId, MANUSCRIPTS_SHEET_NAME, 'MS_ID', reviewMsId);
    
    // 担当編集者へ通知（承諾・辞退いずれも担当編集者メニューへのリンクを付与）
    let editorActionUrl   = null;
    let editorActionLabel = null;
    {
      // Review_log には editorKey がないため、Editor_log から MsVer + edtOk==='ok' で検索
      const acceptedEdLog = findAllRecordsByKey(ssId, EDITOR_LOG_SHEET_NAME, 'MsVer', reviewLog.MsVer || '')
        .find(r => String(r.edtOk || '').trim() === 'ok');
      if (acceptedEdLog) {
        const editorKey = acceptedEdLog.editorKey;
        if (editorKey) {
          editorActionUrl = ScriptApp.getService().getUrl() + '?editorKey=' + editorKey;
        }
      }
    }
    if (answer === 'ok') {
      editorActionLabel = 'Check Reviewer Status / 査読者の割当状況を確認・追加する';
    } else if (answer === 'ng') {
      editorActionLabel = 'Assign New Candidate / 新しい候補者を指名する';
    }
    sendAssignmentResponseNotificationToRequester(reviewLog.Editor_Email, reviewLog.Editor_Name, reviewLog.Rev_Name, ms || {}, answer, message, settings, editorActionUrl, editorActionLabel);
    
    // 3. 受諾した場合はWelcomeメール送信
    if (answer === 'ok') {
      sendReviewerWelcomeEmail(reviewLog, ms || {}, settings);
    }
  }
  
  return { success: true };
}

/**
 * 未回答の査読者招待を取消す
 */
function apiCancelReviewerInvitation(reviewKey) {
  const ssId = getSpreadsheetId();
  const reviewLog = findRecordByKey(ssId, REVIEW_LOG_SHEET_NAME, 'reviewKey', reviewKey);
  if (!reviewLog) throw new Error('Reviewer record not found.');

  const revOk = String(reviewLog.revOk || '').trim();
  const receivedAt = String(reviewLog.Received_At || '').trim();

  if (revOk === 'ng' || revOk === 'cancelled') {
    throw new Error('この招待はすでに辞退済みまたは取消済みのため操作できません。/ This invitation has already been declined or cancelled.');
  }
  if (receivedAt !== '') {
    throw new Error('査読結果が提出済みのため取り消しできません。/ Cannot cancel: the review has already been submitted.');
  }

  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  updateLogCell(ssId, REVIEW_LOG_SHEET_NAME, 'reviewKey', reviewKey, {
    'revOk':     'cancelled',
    'Answer_At': now
  });

  Logger.log('Reviewer invitation cancelled: ' + reviewKey + ' (' + (reviewLog.Rev_Name || '') + ')');
  return { success: true };
}

function apiCancelEditorAssignment(editorKey) {
  const ssId = getSpreadsheetId();
  const editorLog = findRecordByKey(ssId, EDITOR_LOG_SHEET_NAME, 'editorKey', editorKey);
  if (!editorLog) throw new Error('Editor record not found.');

  const edtOk = String(editorLog.edtOk || '').trim();
  if (edtOk === 'ng' || edtOk === 'cancelled') {
    throw new Error('この割当はすでに辞退済みまたは取消済みのため操作できません。/ This assignment has already been declined or cancelled.');
  }

  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  updateLogCell(ssId, EDITOR_LOG_SHEET_NAME, 'editorKey', editorKey, {
    'edtOk':     'cancelled',
    'Answer_At': now
  });

  Logger.log('Editor assignment cancelled: ' + editorKey + ' (' + (editorLog.Editor_Name || '') + ')');
  return { success: true };
}





/**
 * 汎用的なログ更新関数
 */
function updateLogCell(ssId, sheetName, keyColName, keyValue, updates) {
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // 大文字小文字・前後スペースを無視して列を探すヘルパー
  const findColIndex = (name) => {
    const lowerName = name.toLowerCase().trim();
    return headers.findIndex(h => String(h).toLowerCase().trim() === lowerName);
  };

  const keyIdx = findColIndex(keyColName);
  if (keyIdx === -1) {
    Logger.log("Column " + keyColName + " not found in " + sheetName);
    return;
  }
  
  const trimmedKeyValue = String(keyValue).trim();
  const rowIdx = data.findIndex(r => String(r[keyIdx]).trim() === trimmedKeyValue);
  if (rowIdx === -1) {
    // デバッグ用：実際に格納されている値の先頭数件を出力
    const sampleValues = data.slice(1, Math.min(6, data.length))
      .map((r, i) => '  row' + (i + 1) + ': [' + JSON.stringify(String(r[keyIdx])).slice(0, 60) + ']')
      .join('\n');
    Logger.log('[updateLogCell] Value "' + trimmedKeyValue.slice(0, 40) + '..." not found in column "' + keyColName + '" of sheet "' + sheetName + '". Sample values:\n' + sampleValues);
    return;
  }
  
  Object.keys(updates).forEach(colName => {
    let colIdx = findColIndex(colName);
    if (colIdx === -1) {
      // 列が存在しない場合はシート末尾に列ヘッダーを追加
      colIdx = headers.length;
      sheet.getRange(1, colIdx + 1).setValue(colName);
      headers.push(colName); // ローカルの headers 配列も同期
      Logger.log("Column " + colName + " added to " + sheetName);
    }
    sheet.getRange(rowIdx + 1, colIdx + 1).setValue(updates[colName]);
  });
  SpreadsheetApp.flush();
  spreadsheetCache.invalidate(ssId, sheetName);
}

/**
 * 依頼元（委員長or編集者）への結果通知メール
 */
function sendAssignmentResponseNotificationToRequester(toEmail, toName, responderName, ms, answer, message, settings, actionUrl, actionLabel, noteHtml) {
  const status = answer === 'ok' ? 'ACCEPTED' : 'DECLINED';
  const paperTitle = (ms.TitleJP && ms.TitleEN) ? ms.TitleJP + ' / ' + ms.TitleEN : (ms.TitleJP || ms.TitleEN || 'Unknown Title');

  const subject = `[${settings.Journal_Name}] 招待への回答 / Invitation Response (${status}): ${responderName} — ${ms.MsVer || 'Manuscript'}`;

  // アクションノート（ボタンの上に表示する案内文）
  // noteHtml が渡された場合はそれを優先する（呼び出し元でロール別に制御）
  let actionNoteHtml = noteHtml || '';
  if (!actionNoteHtml) {
    if (actionUrl) {
      if (answer === 'ok') {
        actionNoteHtml = `
          <p style="margin-top:1.5rem;">Please use the button below to check the current reviewer assignment status and add new candidates if needed.</p>
          <hr style="border:none; border-top:1px solid #e2e8f0; margin: 20px 0;">
          <p>以下のボタンから査読者の割当状況を確認し、必要に応じて新たな候補者を追加してください。</p>
        `;
      } else if (answer === 'ng') {
        actionNoteHtml = `
          <p style="margin-top:1.5rem;">Please use the button below to assign a new candidate.</p>
          <hr style="border:none; border-top:1px solid #e2e8f0; margin: 20px 0;">
          <p>つきましては、改めて候補者を指名し、依頼メールをお送りください。以下のボタンよりご対応いただけます。</p>
        `;
      }
    }
  }

  const bodyHtml = `
    <p>We would like to inform you that <strong>${responderName}</strong> has <strong>${status.toLowerCase()}</strong> the invitation for manuscript <strong>${ms.MsVer || ''}</strong>.</p>
    <table style="width:100%; font-size: 14px; border-collapse: collapse; margin: 20px 0;">
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee; width: 30%;">Manuscript / 原稿</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${ms.MsVer || ''}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Type / 種別</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${ms.MS_Type || ''}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Title / タイトル</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${paperTitle}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Response / 回答</th><td style="padding: 8px; border-bottom: 1px solid #eee; font-weight:bold; color:${answer === 'ok' ? '#059669' : '#dc2626'}">${answer === 'ok' ? 'ACCEPT / 受諾' : 'DECLINE / 辞退'}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Message / メッセージ</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${(message || '').replace(/\n/g, '<br>') || '(No message / メッセージなし)'}</td></tr>
    </table>
    <p><strong>${responderName}</strong> 殿より、原稿 <strong>${ms.MsVer || ''}</strong> の依頼に対して <strong>${answer === 'ok' ? '受諾' : '辞退'}</strong> の回答がありました。</p>
    ${actionNoteHtml}
  `;

  const resolvedLabel = actionLabel || (answer === 'ng' ? 'Assign New Candidate / 新しい候補者を指名する' : 'Open Editor Menu / 担当編集メニューを開く');

  const html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting: `Dear Dr. ${toName},`,
    bodyHtml: bodyHtml,
    buttonUrl:   actionUrl || undefined,
    buttonLabel: actionUrl ? resolvedLabel : undefined,
    footerHtml: settings.mailFooter || ''
  });

  sendEmailSafe({ to: toEmail, subject, htmlBody: html },
    'Assignment Response Notification to ' + toName + ': ' + (ms.MsVer || ''));
}

/**
 * 査読コメントを遅延取得する API
 * buildReviewerList では Doc ID のみ返し、フロントエンドがこの関数を
 * on-demand で呼び出してコメント本文を取得する。
 */
function apiGetReviewComments(reviewKey) {
  const ssId = getSpreadsheetId();
  const reviewLog = findRecordByKey(ssId, REVIEW_LOG_SHEET_NAME, 'reviewKey', reviewKey);
  if (!reviewLog) return { error: 'Review record not found.' };

  return {
    openCommentsText: readDocText(String(reviewLog.openCommentsId || '').trim()),
    confidentialCommentsText: readDocText(String(reviewLog.confidentialCommentsId || '').trim())
  };
}


/**
 * 原稿をアーカイブシートに移動する (デスクリジェクト・取下げ用)
 * 審査前の却下など、メインのダッシュボードから除外したい場合に実行します。
 * 
 * @param {string} key 著者キー (Manuscripts.key)
 */
function apiArchiveManuscript(key) {
  const ssId = getSpreadsheetId();
  const ss = SpreadsheetApp.openById(ssId);
  const msSheet = ss.getSheetByName(MANUSCRIPTS_SHEET_NAME);
  
  if (!msSheet) throw new Error('Manuscripts sheet not found.');

  // 1. Manuscripts シートから該当レコードを取得
  // (ManuscriptDataHandlers.js の汎用関数を使用)
  const msData = findRecordByKey(ssId, MANUSCRIPTS_SHEET_NAME, 'key', key);
  if (!msData) throw new Error('原稿データが見つかりませんでした。データが既に移動されている可能性があります。 / Manuscript not found.');
  
  const headers = msSheet.getRange(1, 1, 1, msSheet.getLastColumn()).getValues()[0];
  const data = msSheet.getDataRange().getValues();
  const keyIdx = headers.findIndex(h => String(h).toLowerCase().trim() === 'key');
  const targetRowIdx = data.findIndex(r => String(r[keyIdx]) === String(key));
  
  if (targetRowIdx === -1) {
    throw new Error('Manuscripts シート内に行が見つかりませんでした。 / Target row not found in Manuscripts sheet.');
  }

  // 2. Archive シートの準備（定数 ARCHIVE_SHEET_NAME または 'Archive'）
  const archiveName = ARCHIVE_SHEET_NAME || 'Archive';
  let archiveSheet = ss.getSheetByName(archiveName);
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet(archiveName);
    // Manuscripts シートと同じヘッダーをコピー
    archiveSheet.appendRow(headers);
    archiveSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f4f6');
    archiveSheet.setFrozenRows(1);
  } else {
    // 既存のアーカイブシートがある場合、ヘッダーの同期を確認（簡易）
    const archHeaders = archiveSheet.getRange(1, 1, 1, archiveSheet.getLastColumn()).getValues()[0];
    if (archHeaders.length < headers.length) {
      archiveSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
  
  // 3. データの移動 (行データを構築して追加)
  const rowDataArray = headers.map(h => {
    const val = msData[String(h).trim()];
    return val === undefined ? '' : val;
  });
  archiveSheet.appendRow(rowDataArray);
  
  // 4. Manuscripts シートから削除 (行番号は 1-indexed)
  msSheet.deleteRow(targetRowIdx + 1);
  
  // 5. Google ドライブ上のフォルダをリネーム（任意）
  // フォルダ名に [ARCHIVED] を付けて視覚的に判別しやすくする
  if (msData.folderUrl || msData.folderID) {
    try {
      let folderId = msData.folderID;
      if (!folderId && msData.folderUrl) {
         const match = msData.folderUrl.match(/[-\w]{25,}/);
         if (match) folderId = match[0];
      }
      
      if (folderId) {
        const folder = DriveApp.getFolderById(folderId);
        const oldName = folder.getName();
        if (!oldName.includes('[ARCHIVED]')) {
          folder.setName('[ARCHIVED] ' + oldName);
        }
      }
    } catch (e) {
      Logger.log('Folder rename failed during archive: ' + e);
    }
  }
  
  writeLog(`Manuscript Archived: MS_ID=${msData.MS_ID}, ver=${msData.MsVer}, Author=${msData.CA_Name}`);
  
  // キャッシュを無効化
  spreadsheetCache.invalidate(ssId, MANUSCRIPTS_SHEET_NAME);
  spreadsheetCache.invalidate(ssId, archiveName);
  
  return { success: true, msId: msData.MS_ID };
}

function writeLog(text) {
  const ssId = getSpreadsheetId();
  if (!ssId) return;
  const ss = SpreadsheetApp.openById(ssId);
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    sheet.getRange(1, 1, 1, 2).setValues([['Timestamp', 'Message']])
      .setFontWeight('bold').setBackground('#f1f5f9');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([new Date(), text]);
}
