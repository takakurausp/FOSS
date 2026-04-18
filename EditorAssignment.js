/**
 * EditorAssignment.gs - 担当編集者選定・依頼ロジック
 */

/**
 * 編集委員長が担当編集者を指名・依頼するAPI
 * @param {Object} data { msKey, editorName, editorEmail, letterToEditor }
 */
function apiAssignEditor(data) {
  // 入力バリデーション
  validateRequiredString(data.msKey,       '原稿キー (msKey)');
  validateRequiredString(data.editorName,  '担当編集者名 (editorName)');
  validateEmail(data.editorEmail,          '担当編集者メールアドレス (editorEmail)');
  // letterToEditor は任意項目のため検証不要

  const ssId = getSpreadsheetId();
  const settings = getSettings();

  // 1. 論文情報を取得（著者キーまたは編集委員長専用キーによる検索）
  let ms = findRecordByKey(ssId, MANUSCRIPTS_SHEET_NAME, 'key', data.msKey);
  if (!ms) ms = findRecordByKey(ssId, MANUSCRIPTS_SHEET_NAME, 'eicKey', data.msKey);
  if (!ms) throw new Error('Manuscript not found for key: ' + data.msKey);

  // 2. 現在の担当編集者候補の状態を確認
  const msVer = ms['MsVer'] || '';
  const editorLogs = findAllRecordsByKey(ssId, EDITOR_LOG_SHEET_NAME, 'MsVer', msVer);
  
  // 辞退していない（未回答または承諾）候補者がいるか確認
  const hasActiveCandidate = editorLogs.some(log => {
    const edtOk = String(log.edtOk || '').trim();
    return edtOk === '' || edtOk === 'ok'; // 未回答または承諾
  });
  
  if (hasActiveCandidate) {
    throw new Error('既に担当編集者候補が割り当てられており、まだ辞退していません。新たな候補者の指名はできません。/ An editor candidate has already been assigned and has not declined yet. You cannot assign a new candidate.');
  }

  // 3. Editor用キーの発行
  const msVerRevHex = getMsVerRevHexFromMsVer(msVer);
  const editorKey = msVerRevHex + Utilities.getUuid().replace(/-/g, '');  // CSPRNG (128 bit)

  // 4. Editor_logへの記録
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  addEditorLogEntry(ssId, {
    MsVer: msVer,
    MsVerRevHex: msVerRevHex,
    editorKey: editorKey,
    Editor_Name: data.editorName,
    Editor_Email: data.editorEmail,
    Ask_At: now
  });

  // 5. 担当編集者への依頼メール送信
  sendEditorRequestEmail(ms, data, editorKey, settings);

  writeLog(`Editor Assigned: ${msVer} by Chief Editor (Target: ${data.editorName})`);

  return { success: true, editorKey: editorKey };
}

/**
 * MsVer（例: JJEEZ3-1）から16進数の EditorRevHex を生成
 */
function getMsVerRevHexFromMsVer(msVer) {
  const { msId, verNo: verNum } = parseMsVer(msVer);
  const numericId = parseInt(msId.replace(/[^0-9]/g, ''), 10) || 0;
  
  const hexA = ('0000' + numericId.toString(16)).slice(-4);
  const hexB = verNum.toString(16);
  return 'H' + hexA + hexB + '0'; // 末尾の0はeditor/reviewer用サフィックス
}

/**
 * Editor_logシートに新規エントリを追加
 */
function addEditorLogEntry(ssId, entry) {
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(EDITOR_LOG_SHEET_NAME);
  if (!sheet) {
    Logger.log('Editor_log sheet not found.');
    return;
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newRow = new Array(headers.length).fill('');
  
  headers.forEach((h, i) => {
    const key = Object.keys(entry).find(k => k.toLowerCase() === String(h).toLowerCase());
    if (key !== undefined) newRow[i] = entry[key];
  });
  
  sheet.appendRow(newRow);
  SpreadsheetApp.flush();
  spreadsheetCache.invalidate(ssId, EDITOR_LOG_SHEET_NAME);
  Logger.log('[addEditorLogEntry] Row appended. editorKey=' + (entry.editorKey || '(none)'));
}

/**
 * 担当編集者候補への依頼メール送信
 */
function sendEditorRequestEmail(ms, data, editorKey, settings) {
  const webAppUrl = ScriptApp.getService().getUrl();
  const editorUrl = webAppUrl + '?editorKey=' + editorKey;
  
  const paperTitle = escHtml([ms['TitleJP'], ms['TitleEN']].filter(Boolean).join(' / '));
  const subject = `[${settings.Journal_Name}] 担当編集者ご就任のご依頼 / Invitation for Responsible Editor (${ms['MsVer']})`;

  const bodyHtml = `
    <p>You have been invited to serve as the responsible editor for the following manuscript. Please respond to this invitation by clicking the button below.</p>
    <p>以下の原稿について、担当編集者（Responsible Editor）への就任をお願いしたく存じます。内容をご確認の上、以下のボタンより受諾または辞退のご回答をお願いいたします。</p>
    <div style="background:#f1f5f9; padding:15px; border-radius:8px; margin:20px 0;">
      <p><strong>Message from Chief Editor:</strong><br>${escHtml(data.letterToEditor || '').replace(/\n/g, '<br>')}</p>
    </div>
    <table style="width:100%; font-size: 14px; border-collapse: collapse; margin: 20px 0;">
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee; width: 30%;">ID</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${escHtml(ms['MsVer'])}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Type</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${escHtml(ms['MS_Type'] || '')}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Title</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${paperTitle}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Abstract (JP)</th><td style="padding: 8px; border-bottom: 1px solid #eee; white-space: pre-wrap;">${escHtml(ms['AbstractJP'] || 'N/A')}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Abstract (EN)</th><td style="padding: 8px; border-bottom: 1px solid #eee; white-space: pre-wrap;">${escHtml(ms['AbstractEN'] || 'N/A')}</td></tr>
    </table>
  `;

  const html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting: `Dear Dr. ${data.editorName},`,
    bodyHtml: bodyHtml,
    buttonUrl: editorUrl,
    buttonLabel: 'Respond to Invitation / 依頼に回答する',
    footerHtml: settings.mailFooter || ''
  });

  sendEmailSafe({ to: data.editorEmail, subject, htmlBody: html },
    'Editor Invitation: ' + ms['MsVer'] + ' to ' + data.editorName);
}

/**
 * 担当編集者が承諾した後に送るウェルカムメール（メンテ用リンク含む）
 */
function sendEditorWelcomeEmail(editorLog, settings) {
  const webAppUrl = ScriptApp.getService().getUrl();
  const edName = editorLog.Editor_Name || editorLog.editorName;
  const edEmail = editorLog.Editor_Email || editorLog.editorEmail;
  const editorKey = editorLog.editorKey;
  const editorMenuUrl = webAppUrl + '?editorKey=' + editorKey;
  
  const subject = `[${settings.Journal_Name}] 担当編集者ご就任の確認 / Assignment Confirmed: ${editorLog.MsVer}`;
  const bodyHtml = `
    <p>Thank you for accepting the assignment as the responsible editor for manuscript <strong>${editorLog.MsVer}</strong>.</p>
    <p>Please use the button below to access your editor management menu. Start by selecting reviewers, then manage the peer review process as it progresses.</p>
    <hr style="border:none; border-top:1px solid #e2e8f0; margin: 20px 0;">
    <p>原稿 <strong>${editorLog.MsVer}</strong> の担当編集者をお引き受けいただき、誠にありがとうございます。</p>
    <p>以下のボタンより担当編集者専用メニューにアクセスしてください。まず査読者の選定を行い、その後、同メニューから査読の進行管理を行うことができます。</p>
  `;

  const html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting: `Dear Dr. ${edName},`,
    bodyHtml: bodyHtml,
    buttonUrl: editorMenuUrl,
    buttonLabel: 'Open Editor Menu / 編集担当メニューを開く',
    footerHtml: settings.mailFooter || ''
  });

  sendEmailSafe({ to: edEmail, subject, htmlBody: html },
    'Editor Welcome: ' + editorLog.MsVer + ' to ' + edName);
}
