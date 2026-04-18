/**
 * FinalReviewModule.js - 編集幹事・委員長最終確認フロー（受理ルート）
 *
 * 【前提となる Manuscripts シートへの追加列】
 * - managingEditorKey           : 編集幹事アクセス用ユニークキー
 * - managingEditorAuthorComment : 著者宛コメント（編集幹事作成）
 * - managingEditorInternalComment : 委員長・印刷担当者宛コメント（編集幹事作成、内部用）
 * - managingEditorFileUrl       : 編集幹事アップロードファイルのフォルダ URL
 * - managingEditorSentAt        : 編集幹事が委員長へ送信した日時
 * - eicFinalComment             : 委員長の最終コメント
 * - eicFinalFileUrl             : 委員長アップロードファイルのフォルダ URL
 * - finalStatus                 : 'final_review' | 'in_production' | ''
 *
 * 【Settings シートへの追加設定】
 * - managingEditorEmail  : 編集幹事のメールアドレス
 * - productionEditorEmail: 印刷担当者のメールアドレス
 */

/**
 * 編集幹事が委員長にレビューを送信するAPI
 * data: { managingEditorKey, authorComment, internalComment, reportGoogleDocId, files[] }
 *
 * 編集幹事のオープンコメントは、担当編集者・査読者のコメントが既に書き込まれた
 * Google Docs（reportGoogleDocId）にセクションを追加する形で保存する。
 * EIC はこの Google Docs を編集・確認したうえで PDF 化して著者に送付する。
 */
function apiSubmitManagingEditorReview(data) {
  const ssId = getSpreadsheetId();
  const settings = getSettings();

  const msData = getManuscriptData('managing-editor', data.managingEditorKey);
  if (!msData) throw new Error('原稿が見つかりません。/ Manuscript not found.');

  const now = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm');

  // 書き込み済み原稿ファイルを Drive に保存
  let fileUrl = '';
  if (data.files && data.files.length > 0) {
    const verFolder = getManuscriptVerFolder(msData, settings);
    const meFolder = driveFolderCache.getOrCreateFolder(verFolder, 'managing-editor');
    data.files.forEach(function(file) {
      const blob = Utilities.newBlob(Utilities.base64Decode(file.content), file.mimeType, file.name);
      meFolder.createFile(blob);
    });
    meFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    fileUrl = meFolder.getUrl();
  }

  // 編集幹事の著者宛コメントをオープンコメント集 Google Docs に追記
  // （フロントエンドから渡された reportGoogleDocId を優先し、
  //   未渡しの場合は EditorLog から取得する）
  var docId = (data.reportGoogleDocId || '').trim();
  if (!docId && msData._editorList && msData._editorList.length > 0) {
    for (var ei = 0; ei < msData._editorList.length; ei++) {
      var eDoc = (msData._editorList[ei].reportGoogleDocId || '').trim();
      if (eDoc) { docId = eDoc; break; }
    }
  }
  if (docId && data.authorComment && data.authorComment.trim()) {
    try {
      var doc  = DocumentApp.openById(docId);
      var body = doc.getBody();
      // セクション区切り
      body.appendHorizontalRule();
      var header = body.appendParagraph('Section 3: Managing Editor\'s Comments / 編集幹事コメント');
      header.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.appendParagraph(data.authorComment.trim());
      doc.saveAndClose();
    } catch (docErr) {
      Logger.log('ME comment append to Google Doc failed: ' + docErr.message);
    }
  }

  // Manuscripts シートを更新
  updateLogCell(ssId, MANUSCRIPTS_SHEET_NAME, 'key', msData.key, {
    'managingEditorAuthorComment':   data.authorComment   || '',
    'managingEditorInternalComment': data.internalComment || '',
    'managingEditorFileUrl':         fileUrl,
    'managingEditorSentAt':          now
  });

  // 担当編集者レポート＋編集幹事コメントを合わせた最終確認レポートPDFを生成
  var reportPdfBlob = null;
  try {
    reportPdfBlob = createFinalReviewReport(msData, data, settings, ssId);
  } catch (pdfErr) {
    Logger.log('createFinalReviewReport failed: ' + pdfErr.message);
  }

  // 委員長へ通知
  _sendManagingEditorReviewToEIC(msData, data, fileUrl, settings, reportPdfBlob);

  writeLog('Managing Editor Review Submitted: ' + (msData.MsVer || ''));
  return { success: true };
}

/**
 * 委員長が最終アクション（ルートa/b/c）を実行するAPI
 * data: { eicKey, route ('a'|'b'|'c'), eicComment, files[] }
 */
function apiEicFinalAction(data) {
  const ssId = getSpreadsheetId();
  const settings = getSettings();

  const msData = getManuscriptData('eic', data.eicKey);
  if (!msData) throw new Error('原稿が見つかりません。/ Manuscript not found.');

  // eic-final フォルダを作成し、EIC添付ファイルをアップロード（全ルート共通）
  const verFolder = getManuscriptVerFolder(msData, settings);
  const eicFolder = driveFolderCache.getOrCreateFolder(verFolder, 'eic-final');
  const hasFiles = data.files && data.files.length > 0;

  if (hasFiles) {
    data.files.forEach(function(file) {
      const decoded = Utilities.base64Decode(file.content);
      eicFolder.createFile(Utilities.newBlob(decoded, file.mimeType, file.name));
    });
  }

  // ルートB・C 用: ファイルがあれば今の時点でフォルダURLを確定する
  // ルートA はコメントPDF保存後に上書きする
  var eicFileUrl = hasFiles
    ? (eicFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW),
       eicFolder.getUrl())
    : '';

  // 委員長コメント・ファイルURL・判定名を Manuscripts に保存（全ルート共通）
  // ルートA はコメントPDF保存後に eicFinalFileUrl だけ再更新する
  updateLogCell(ssId, MANUSCRIPTS_SHEET_NAME, 'key', msData.key, {
    'eicFinalComment':      data.eicAuthorComment     || '',
    'eicProductionComment': data.eicProductionComment || '',
    'eicFinalFileUrl':      eicFileUrl,
    'eicFinalDecision':     data.decision             || ''
  });

  var route = data.route;

  if (route === 'a') {
    // ルートa: 投稿者に差し戻す — sentBackAt / score / accepted を記録
    var dtRouteA = getDecisionTemplates(ssId, data.decision || '');
    var nowRouteA = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm');
    updateLogCell(ssId, MANUSCRIPTS_SHEET_NAME, 'key', msData.key, {
      'sentBackAt': nowRouteA,
      'score':      data.decision || '',
      'accepted':   dtRouteA.isAccepted ? 'yes' : 'no'
    });

    // コメントPDFを eic-final フォルダへ保存（著者向けフォルダに統合）
    if (data.commentDocId) {
      try {
        var journalNameA = (settings && settings.Journal_Name) ? settings.Journal_Name : '';
        var nowStrA = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
        var msVerA = msData.MsVer || '';

        // ① 原本の先頭にヘッダーを挿入
        try {
          var commentDocA = DocumentApp.openById(data.commentDocId);
          insertCommentDocHeader(commentDocA.getBody(), journalNameA, data.decision || '', nowStrA);
          commentDocA.saveAndClose();
        } catch(eInsertA) {
          Logger.log('Header insertion failed (route a): ' + eInsertA.message);
        }

        // ② 原本から PDF を取得
        var commentPdfBlob = DriveApp.getFileById(data.commentDocId).getAs(MimeType.PDF);
        commentPdfBlob.setName('Open-Comments-' + msVerA + '.pdf');

        // ③ PDF を eic-final フォルダに保存して共有リンクを発行
        var savedCommentPdf = eicFolder.createFile(commentPdfBlob);
        savedCommentPdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        var commentPdfUrlA = savedCommentPdf.getUrl();
        if (data.commentEditorKey) {
          updateLogCell(ssId, EDITOR_LOG_SHEET_NAME, 'editorKey', data.commentEditorKey,
            { 'reportCommentPdfUrl': commentPdfUrlA });
        }

        // PDFが追加されたのでフォルダURLを確定して Manuscripts を更新
        eicFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        eicFileUrl = eicFolder.getUrl();
        updateLogCell(ssId, MANUSCRIPTS_SHEET_NAME, 'key', msData.key, { 'eicFinalFileUrl': eicFileUrl });

      } catch(e) {
        Logger.log('Comment PDF export failed (route a): ' + e.message);
      }
    }

    _sendFinalRouteAToAuthor(msData, data, eicFileUrl, settings, ssId);
    _notifyManagingEditorOfEicRoute(msData, 'a', data.eicAuthorComment, data.eicProductionComment, data.decision || '', settings, ssId);

  } else if (route === 'b') {
    // ルートb: 印刷担当者に送付、ステータスを in_production へ
    updateLogCell(ssId, MANUSCRIPTS_SHEET_NAME, 'key', msData.key, {
      'finalStatus': 'in_production',
      'accepted':    'yes'
    });
    _sendFinalRouteBToProductionEditor(msData, data, eicFileUrl, settings);
    _notifyAuthorOfAcceptance(msData, data, settings, ssId);
    _notifyManagingEditorOfEicRoute(msData, 'b', data.eicAuthorComment, data.eicProductionComment, data.decision || '', settings, ssId);

  } else if (route === 'c') {
    // ルートc: 担当編集者に差し戻し（再判定依頼）
    _resetEditorScoreForMsVer(ssId, msData.MsVer || '');
    updateLogCell(ssId, MANUSCRIPTS_SHEET_NAME, 'key', msData.key, {
      'finalStatus':                   '',
      'managingEditorKey':             '',
      'managingEditorAuthorComment':   '',
      'managingEditorInternalComment': '',
      'managingEditorFileUrl':         '',
      'managingEditorSentAt':          ''
    });
    _sendFinalRouteCToEditor(msData, data, eicFileUrl, settings, ssId);
    _notifyManagingEditorOfEicRoute(msData, 'c', data.eicAuthorComment, data.eicProductionComment, data.decision || '', settings, ssId);
  }

  writeLog('EIC Final Action: ' + (msData.MsVer || '') + ' - Route: ' + route);
  return { success: true };
}

/* ─── プライベートヘルパー ──────────────────────────────────── */

/**
 * 承諾済み担当編集者の Score と Received_At をリセット（ルートc用）
 */
function _resetEditorScoreForMsVer(ssId, msVer) {
  if (!msVer) return;
  var sheet = SpreadsheetApp.openById(ssId).getSheetByName(EDITOR_LOG_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var findCol = function(name) {
    return headers.findIndex(function(h) { return String(h).toLowerCase().trim() === name.toLowerCase().trim(); });
  };
  var msVerIdx = findCol('MsVer');
  var edtOkIdx = findCol('edtOk');
  var scoreIdx = findCol('Score');
  var rcvAtIdx = findCol('Received_At');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][msVerIdx]).trim() === msVer &&
        String(data[i][edtOkIdx]).trim() === 'ok') {
      if (scoreIdx !== -1) sheet.getRange(i + 1, scoreIdx + 1).setValue('');
      if (rcvAtIdx !== -1) sheet.getRange(i + 1, rcvAtIdx + 1).setValue('');
      break;
    }
  }
  SpreadsheetApp.flush();
  spreadsheetCache.invalidate(ssId, EDITOR_LOG_SHEET_NAME);
}

/**
 * 担当編集者レポート＋編集幹事コメントを合わせた最終確認レポートPDFを生成し、
 * Drive の working フォルダに保存して blob を返す。
 *
 * @param {Object} msData  getManagingEditorManuscriptData() の戻り値
 * @param {Object} data    apiSubmitManagingEditorReview() の data（authorComment, internalComment）
 * @param {Object} settings
 * @param {string} ssId
 * @returns {Blob|null}  生成した PDF blob。失敗時は null。
 */
function createFinalReviewReport(msData, data, settings, ssId) {
  const journalName = (settings && settings.Journal_Name) ? settings.Journal_Name : 'Journal';
  const now = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm');
  const esc = s => String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const nl2br = s => esc(s).replace(/\n/g, '<br>');

  const sectionTitle = title =>
    `<div style="background:#1e40af; color:#fff; padding:6px 12px; margin:24px 0 10px; border-radius:4px; font-size:13px; font-weight:bold;">${esc(title)}</div>`;
  const infoRow = (labelJp, labelEn, value) =>
    `<tr>
      <th style="text-align:left; padding:5px 10px; border:1px solid #d1d5db; background:#f3f4f6; width:32%; font-size:11px; vertical-align:top; white-space:nowrap;">
        ${esc(labelJp)}<br><span style="font-size:9.5px; color:#6b7280;">${esc(labelEn)}</span>
      </th>
      <td style="padding:5px 10px; border:1px solid #d1d5db; font-size:11px; vertical-align:top; white-space:pre-wrap;">${esc(value)}</td>
    </tr>`;
  const infoRowHtml = (labelJp, labelEn, valueHtml) =>
    `<tr>
      <th style="text-align:left; padding:5px 10px; border:1px solid #d1d5db; background:#f3f4f6; width:32%; font-size:11px; vertical-align:top; white-space:nowrap;">
        ${esc(labelJp)}<br><span style="font-size:9.5px; color:#6b7280;">${esc(labelEn)}</span>
      </th>
      <td style="padding:5px 10px; border:1px solid #d1d5db; font-size:11px; vertical-align:top;">${valueHtml}</td>
    </tr>`;
  const commentBox = (label, text, bgColor, borderColor) =>
    `<div style="margin:6px 0; padding:8px 12px; background:${bgColor}; border:1px solid ${borderColor}; border-radius:6px;">
      <p style="margin:0 0 4px; font-size:10px; font-weight:bold; color:#374151;">${esc(label)}</p>
      <p style="margin:0; font-size:11px; white-space:pre-wrap;">${nl2br(text) || '<span style="color:#9ca3af;">(なし / None)</span>'}</p>
    </div>`;

  // 承諾済み担当編集者を取得
  const editorList = msData._editorList || [];
  const acceptedEd = editorList.find(e => e.edtOk === 'ok') || editorList[0] || {};

  // 査読者コメント（Google Docs から本文を取得）
  const reviewLogLines = getFilteredReviewLog(ssId, msData.MsVer || '');

  let reviewerSections = '';
  reviewLogLines.forEach(function(rev, idx) {
    reviewerSections +=
      `<div style="margin-bottom:16px; padding:10px 14px; border:1px solid #e2e8f0; border-radius:8px; break-inside:avoid;">
        <p style="margin:0 0 6px; font-size:13px; font-weight:bold;">Reviewer #${idx + 1}: ${esc(rev.Rev_Name)}</p>
        <table style="width:100%; border-collapse:collapse; margin-bottom:8px;">
          ${infoRow('判定スコア', 'Score', rev.Score || '')}
          ${infoRow('査読結果提出日', 'Submitted', String(rev.Received_At || ''))}
        </table>
        ${commentBox('オープンコメント / Open Comments', rev.openCommentsText || '', '#f8fafc', '#e2e8f0')}
        ${commentBox('🔒 コンフィデンシャルコメント / Confidential Comments (not shared with authors)', rev.confidentialCommentsText || '', '#fffbeb', '#fcd34d')}
        ${rev.folderUrl && rev.folderUrl !== 'nofile'
          ? `<p style="margin:6px 0 0; font-size:10px;"><a href="${esc(rev.folderUrl)}">📁 添付ファイルフォルダ / Attached Files Folder</a></p>` : ''}
      </div>`;
  });

  const pdfHtml = `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<style>
  * { box-sizing: border-box; }
  body { font-family: 'Hiragino Kaku Gothic Pro','Meiryo','Arial',sans-serif; font-size:11px; color:#111827; margin:0; padding:28px 36px; line-height:1.55; }
  h1  { font-size:18px; margin:0; color:#fff; }
  h2  { font-size:14px; margin:10px 0 4px; color:#1e40af; }
  table { width:100%; border-collapse:collapse; margin-bottom:10px; }
  .meta { font-size:10px; color:#6b7280; margin-bottom:16px; }
  .page-break { page-break-before: always; }
</style>
</head>
<body>
  <div style="background:#1e40af; padding:18px 20px; margin:-28px -36px 24px; border-bottom:4px solid #1d4ed8;">
    <h1>${esc(journalName)}</h1>
    <p style="color:#bfdbfe; margin:4px 0 0; font-size:12px;">Final Review Report / 最終確認レポート</p>
  </div>

  <p class="meta">原稿番号 / Manuscript No.: <strong>${esc(msData.MsVer || '')}</strong> &nbsp;|&nbsp; 発行日時 / Issued: <strong>${now}</strong></p>

  ${sectionTitle('原稿情報 / Manuscript Information')}
  <table>
    ${infoRow('原稿番号', 'Manuscript ID', msData.MsVer || '')}
    ${infoRow('原稿種別', 'Type', msData.MS_Type || '')}
    ${infoRow('タイトル（日）', 'Title (JP)', msData.TitleJP || '')}
    ${infoRow('タイトル（英）', 'Title (EN)', msData.TitleEN || '')}
    ${infoRow('著者（日）', 'Authors (JP)', msData.AuthorsJP || '')}
    ${infoRow('著者（英）', 'Authors (EN)', msData.AuthorsEN || '')}
    ${infoRow('責任著者', 'Corresponding Author', (msData.CA_Name || '') + (msData.CA_Email ? ' (' + msData.CA_Email + ')' : ''))}
    ${infoRow('投稿日時', 'Submitted at', msData.Submitted_At || '')}
  </table>

  ${sectionTitle('担当編集者の推薦 / Responsible Editor\'s Recommendation')}
  <table>
    ${infoRow('担当編集者', 'Responsible Editor', (acceptedEd.Editor_Name || '') + (acceptedEd.Editor_Email ? ' (' + acceptedEd.Editor_Email + ')' : ''))}
    ${infoRowHtml('推薦スコア', 'Recommended Score', `<strong style="color:#1e40af; font-size:14px;">${esc(acceptedEd.Score || '')}</strong>`)}
    ${infoRow('推薦提出日時', 'Submitted at', acceptedEd.Answer_At || '')}
  </table>
  ${commentBox('オープンコメント / Open Comments (for authors)', acceptedEd.Message || '', '#f8fafc', '#e2e8f0')}
  ${commentBox('🔒 コンフィデンシャルコメント / Confidential Comments (for EIC only)', acceptedEd.ConfidentialMessage || '', '#fffbeb', '#fcd34d')}

  ${reviewLogLines.length > 0 ? sectionTitle('査読結果 / Peer Review Results (' + reviewLogLines.length + ' reviewers)') : ''}
  ${reviewerSections}

  ${sectionTitle('編集幹事の確認 / Managing Editor\'s Review')}
  <table>
    ${infoRow('編集幹事確認日時', 'Reviewed at', now)}
  </table>
  ${commentBox('著者宛コメント / Comment for Author', data.authorComment || '', '#f0fdf4', '#86efac')}
  ${commentBox('🔒 内部用コメント / Internal Comment (for EIC / production staff)', data.internalComment || '', '#fffbeb', '#fcd34d')}
</body>
</html>`;

  const pdfBlob = HtmlService.createHtmlOutput(pdfHtml).getBlob().getAs(MimeType.PDF);

  // Drive の working フォルダに保存
  try {
    const verFolder = getManuscriptVerFolder(msData, settings);
    const workingFolder = driveFolderCache.getOrCreateFolder(verFolder, 'working');
    const pdfFile = workingFolder.createFile(pdfBlob).setName('Final-Report-' + (msData.MsVer || '') + '.pdf');
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return pdfFile.getBlob().setName('Final-Report-' + (msData.MsVer || '') + '.pdf');
  } catch (saveErr) {
    Logger.log('createFinalReviewReport: Drive save failed: ' + saveErr.message + ' — using in-memory blob');
    pdfBlob.setName('Final-Report-' + (msData.MsVer || '') + '.pdf');
    return pdfBlob;
  }
}

/**
 * 編集幹事 → 委員長 通知メール
 */
function _sendManagingEditorReviewToEIC(msData, data, fileUrl, settings, reportPdfBlob) {
  if (!settings.chiefEditorEmail) {
    Logger.log('_sendManagingEditorReviewToEIC: chiefEditorEmail 未設定のためスキップ');
    return;
  }
  var webAppUrl = ScriptApp.getService().getUrl();
  var eicLink = webAppUrl + '?eicKey=' + (msData.eicKey || '');
  var paperTitle = _buildPaperTitle(msData);

  var bodyHtml =
    '<p>Managing editor has completed the initial review of the accepted manuscript. Please access the EIC dashboard below to take the final action.</p>' +
    '<p>編集幹事より、受理原稿の最終確認レビューが送信されました。以下のボタンから委員長ダッシュボードにアクセスし、最終アクションをご実施ください。</p>' +
    '<table style="width:100%; font-size:14px; border-collapse:collapse; margin:20px 0;">' +
      '<tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee; width:30%;">原稿番号 / MS ID</th>' +
          '<td style="padding:8px; border-bottom:1px solid #eee;">' + escHtml(msData.MsVer || '') + '</td></tr>' +
      '<tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">タイトル / Title</th>' +
          '<td style="padding:8px; border-bottom:1px solid #eee;">' + escHtml(paperTitle) + '</td></tr>' +
      '<tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">著者宛コメント<br><span style="font-size:0.85em;font-weight:normal;">Comment for Author</span></th>' +
          '<td style="padding:8px; border-bottom:1px solid #eee; white-space:pre-wrap;">' + escHtml(data.authorComment || '(なし / None)') + '</td></tr>' +
      '<tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">内部用コメント<br><span style="font-size:0.85em;font-weight:normal;">Internal Comment</span></th>' +
          '<td style="padding:8px; border-bottom:1px solid #eee; white-space:pre-wrap;">' + escHtml(data.internalComment || '(なし / None)') + '</td></tr>' +
    '</table>';

  var html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting: 'Dear Editor-in-Chief / 編集委員長殿,',
    bodyHtml: bodyHtml,
    buttonUrl: eicLink,
    buttonLabel: 'Open EIC Dashboard / 委員長ダッシュボードを開く',
    footerHtml: settings.mailFooter || ''
  });

  var emailOptions = {
    to: settings.chiefEditorEmail,
    subject: '[' + settings.Journal_Name + '] 編集幹事より最終確認済み / Managing Editor Review Ready: ' + (msData.MsVer || ''),
    htmlBody: html
  };
  if (reportPdfBlob) emailOptions.attachments = [reportPdfBlob];

  sendEmailSafe(emailOptions, 'Managing Editor Review to EIC: ' + (msData.MsVer || ''));
}

/**
 * ルートa: 著者への通知メール
 * DecisionMailテンプレート＋編集幹事・委員長コメントを合わせて送付
 * EIC添付ファイル・コメントPDFは eic-final フォルダの 1 リンクにまとめて送付する
 */
function _sendFinalRouteAToAuthor(msData, data, eicFileUrl, settings, ssId) {
  var webAppUrl = ScriptApp.getService().getUrl();
  var authorUrl = webAppUrl + '?key=' + (msData.key || '');

  var meAuthorComment = msData.managingEditorAuthorComment || '';
  var eicComment      = data.eicAuthorComment || '';
  // DecisionMailシートからテンプレートを取得
  var decisionTemplates = getDecisionTemplates(ssId || getSpreadsheetId(), data.decision || '');
  var resubmissionUrl = authorUrl;

  // 再投稿期限日の計算
  var limitStrA = String(settings.Resubmittion_Limit || '8 weeks').toLowerCase();
  var weeksMatchA = limitStrA.match(/(\d+)\s*weeks?/);
  var weeksA = weeksMatchA ? parseInt(weeksMatchA[1]) : 8;
  var dueDateObjA = new Date();
  dueDateObjA.setDate(dueDateObjA.getDate() + (weeksA * 7));
  var dueDateStrA = Utilities.formatDate(dueDateObjA, 'JST', 'yyyy/MM/dd');

  var replacements = {
    'authorName':         msData.CA_Name || '',
    'englishTitle':       msData.TitleEN || '',
    'Journal_Name':       settings.Journal_Name || '',
    'Resubmittion_Limit': settings.Resubmittion_Limit || '8 weeks',
    'manuscriptID':       msData.MsVer || '',
    'Editor_Name':        settings.Editor_Name || 'Editor-in-Chief',
    'dueDate':            dueDateStrA,
    'submissionLink':     resubmissionUrl,
    'formlink':           '<a href="' + resubmissionUrl + '">' + resubmissionUrl + '</a>'
  };
  var templateText = replaceDecisionPlaceholders(decisionTemplates.mailText, replacements);

  // 編集幹事・委員長コメント
  var commentsHtml =
    (meAuthorComment ? '<p><strong>編集幹事よりのコメント / Comments from Managing Editor:</strong><br>' + escHtml(meAuthorComment).replace(/\n/g, '<br>') + '</p>' : '') +
    (eicComment      ? '<p><strong>編集委員長よりのコメント / Comments from Editor-in-Chief:</strong><br>' + escHtml(eicComment).replace(/\n/g, '<br>') + '</p>' : '');

  // ファイルリンク（EICファイル・コメントPDFを 1 フォルダにまとめて送付）
  // 投稿原稿フォルダ・編集幹事ファイルはダッシュボードから参照
  var fileLinksHtml = eicFileUrl
    ? '<div style="margin-top:20px; padding:20px; background:#f0f9ff; border:1px solid #bae6fd; border-radius:8px; text-align:center;">' +
        '<p style="margin:0 0 12px 0; font-weight:bold; color:#0369a1; font-size:15px;">📁 判定資料・コメントPDF / Decision Materials &amp; Comments PDF</p>' +
        '<p style="margin:0 0 15px 0; font-size:13.5px; color:#0c4a6e; line-height:1.5;">編集委員長からの添付ファイル・オープンコメントPDFをご確認ください。<br>Please find the EIC\'s attached files and open-comments PDF in the shared folder below.</p>' +
        '<a href="' + eicFileUrl + '" style="display:inline-block; padding:10px 24px; background:#2563eb; color:#ffffff !important; text-decoration:none; border-radius:6px; font-weight:bold; font-size:14.5px;">閲覧用フォルダを開く / Open Shared Folder</a>' +
      '</div>'
    : '';

  var bodyHtml =
    '<div style="background:#f1f5f9; padding:20px; border-radius:8px; margin:20px 0; font-size:15px; line-height:1.6;">' +
      templateText +
    '</div>' +
    commentsHtml +
    fileLinksHtml +
    '<hr style="border:none; border-top:1px solid #e2e8f0; margin:20px 0;">' +
    '<p>原稿 <strong>' + escHtml(msData.MsVer || '') + '</strong> について、編集委員会よりご連絡いたします。</p>' +
    '<p>上記内容および添付ファイルをご確認ください。</p>';

  // 再投稿ボタンは Decisions シートの Resubmit 列が yes の場合のみ表示
  var showResubmitButton = !!decisionTemplates.allowsResubmit;
  var html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting:    'Dear Dr. ' + escHtml(msData.CA_Name || '') + ',',
    bodyHtml:    bodyHtml,
    buttonUrl:   showResubmitButton ? authorUrl : null,
    buttonLabel: showResubmitButton ? 'Submit Revised Manuscript / 修正版を投稿する' : null,
    footerHtml:  settings.mailFooter || ''
  });

  var subject = '[' + settings.Journal_Name + '] 原稿の審査結果について / Decision for your manuscript, ' + (msData.MsVer || '') +
    (decisionTemplates.shortExplanation ? ': ' + decisionTemplates.shortExplanation : '');

  // BCC: 編集幹事 + 担当編集者
  var bccList = [];
  if (settings.managingEditorEmail) bccList.push(settings.managingEditorEmail);
  var acceptedEditorA = _getAcceptedEditorEmail(ssId || getSpreadsheetId(), msData.MsVer || '');
  if (acceptedEditorA && acceptedEditorA.email) bccList.push(acceptedEditorA.email);

  var mailOptions = {
    to:       msData.CA_Email || '',
    cc:       msData.ccEmails || '',
    subject:  subject,
    htmlBody: html
  };
  if (bccList.length > 0) mailOptions.bcc = bccList.join(', ');

  sendEmailSafe(mailOptions, 'Final Route A (to Author): ' + (msData.MsVer || ''));
}

/**
 * ルートb: 印刷担当者への通知メール
 * 原稿基本情報・印刷関連情報・各種コメント＋受領票PDFを添付
 * 委員長アップロードファイルはメール添付ではなく eicFileUrl（Drive リンク）で共有する
 */
function _sendFinalRouteBToProductionEditor(msData, data, eicFileUrl, settings) {
  if (!settings.productionEditorEmail) {
    Logger.log('_sendFinalRouteBToProductionEditor: productionEditorEmail 未設定のためスキップ');
    return;
  }
  var meInternalComment = msData.managingEditorInternalComment || '';
  var eicProductionComment = data.eicProductionComment || '';
  var meFileUrl         = msData.managingEditorFileUrl || '';
  var paperTitle        = _buildPaperTitle(msData);

  // 印刷関連情報
  // Manuscripts シートの列名は 'Reprint request' と 'English_editing'。
  // findRecordByKey はヘッダ名そのままのキーで返すため、ブラケット記法で参照する。
  var reprintInfo  = msData['Reprint request'] || '';
  var editingInfo  = msData['English_editing'] || '';

  // 受領票 PDF の生成
  var receiptBlob = null;
  try {
    receiptBlob = generateReceiptPdf(_mapMsDataForReceipt(msData), settings);
  } catch (e) {
    Logger.log('_sendFinalRouteBToProductionEditor: receipt PDF error: ' + e.message);
  }

  // Drive フォルダリンク
  var receiptFolderUrl = msData.receiptFolderUrl || '';
  var submittedFolderUrl = msData.submittedFolderUrl || msData.folderUrl || '';
  var fileParts = [];
  if (submittedFolderUrl) fileParts.push('<li><a href="' + submittedFolderUrl + '" target="_blank">【閲覧専用】投稿原稿フォルダ / Submitted Files Folder</a></li>');
  if (receiptFolderUrl)   fileParts.push('<li><a href="' + receiptFolderUrl   + '" target="_blank">【閲覧専用】受領票フォルダ / Receipt Folder</a></li>');
  if (meFileUrl)  fileParts.push('<li><a href="' + meFileUrl  + '" target="_blank">【閲覧専用】編集幹事ファイル / Managing Editor\'s Files</a></li>');
  if (eicFileUrl) fileParts.push('<li><a href="' + eicFileUrl + '" target="_blank">【閲覧専用】委員長ファイル / Editor-in-Chief\'s Files</a></li>');
  var fileLinksHtml = fileParts.length > 0
    ? '<p><strong>参照フォルダ / Reference Folders:</strong></p><ul>' + fileParts.join('') + '</ul>'
    : '';

  var bodyHtml =
    '<p>The manuscript listed below has been accepted. Please proceed with the production process at your earliest convenience.</p>' +
    '<p>下記の原稿が受理されましたので、印刷・制作工程への移行をお願いいたします。</p>' +
    '<table style="width:100%; font-size:14px; border-collapse:collapse; margin:20px 0;">' +
      '<tr style="background:#f1f5f9;"><th colspan="2" style="text-align:left; padding:10px 8px; border-bottom:2px solid #cbd5e1; color:#1e40af; font-size:13px;">原稿情報 / Manuscript Information</th></tr>' +
      '<tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee; width:35%;">原稿番号 / Manuscript ID</th>' +
          '<td style="padding:8px; border-bottom:1px solid #eee; font-weight:bold;">' + escHtml(msData.MsVer || '') + '</td></tr>' +
      '<tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">原稿種別 / Type</th>' +
          '<td style="padding:8px; border-bottom:1px solid #eee;">' + escHtml(msData.MS_Type || '') + '</td></tr>' +
      '<tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">タイトル / Title</th>' +
          '<td style="padding:8px; border-bottom:1px solid #eee;">' + escHtml(paperTitle) + '</td></tr>' +
      '<tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">責任著者 / Corresponding Author</th>' +
          '<td style="padding:8px; border-bottom:1px solid #eee;">' + escHtml(msData.CA_Name || '') + '</td></tr>' +
      '<tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">責任著者メール / Email</th>' +
          '<td style="padding:8px; border-bottom:1px solid #eee;">' + escHtml(msData.CA_Email || '') + '</td></tr>' +
      '<tr style="background:#f1f5f9;"><th colspan="2" style="text-align:left; padding:10px 8px; border-bottom:2px solid #cbd5e1; color:#1e40af; font-size:13px;">印刷関連情報 / Production Information</th></tr>' +
      '<tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">別刷希望部数 / Reprints</th>' +
          '<td style="padding:8px; border-bottom:1px solid #eee;">' + escHtml(reprintInfo || '（未記入 / Not specified）') + '</td></tr>' +
      '<tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">英文校閲 / English Editing</th>' +
          '<td style="padding:8px; border-bottom:1px solid #eee;">' + escHtml(editingInfo || '（未記入 / Not specified）') + '</td></tr>' +
      (meInternalComment
        ? '<tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">編集幹事メモ<br><span style="font-size:0.85em;font-weight:normal;">Managing Editor\'s Note</span></th>' +
          '<td style="padding:8px; border-bottom:1px solid #eee; white-space:pre-wrap;">' + escHtml(meInternalComment) + '</td></tr>'
        : '') +
      (eicProductionComment
        ? '<tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">委員長メモ<br><span style="font-size:0.85em;font-weight:normal;">EIC\'s Note</span></th>' +
          '<td style="padding:8px; border-bottom:1px solid #eee; white-space:pre-wrap;">' + escHtml(eicProductionComment) + '</td></tr>'
        : '') +
    '</table>' +
    fileLinksHtml;

  var html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting: 'Dear Production Editor / 印刷担当者殿,',
    bodyHtml: bodyHtml,
    footerHtml: settings.mailFooter || ''
  });

  // 添付ファイル: 受領票PDF のみ（委員長ファイルは Drive リンクで共有済み）
  var attachments = [];
  if (receiptBlob) attachments.push(receiptBlob);

  // BCC: 編集幹事 + 担当編集者
  var bccListB = [];
  if (settings.managingEditorEmail) bccListB.push(settings.managingEditorEmail);
  var acceptedEditorB = _getAcceptedEditorEmail(getSpreadsheetId(), msData.MsVer || '');
  if (acceptedEditorB && acceptedEditorB.email) bccListB.push(acceptedEditorB.email);

  var mailOptions = {
    to:       settings.productionEditorEmail,
    subject:  '[' + settings.Journal_Name + '] 印刷工程依頼 / Production Request: ' + (msData.MsVer || ''),
    htmlBody: html
  };
  if (bccListB.length > 0) mailOptions.bcc = bccListB.join(', ');
  if (attachments.length > 0) mailOptions.attachments = attachments;

  sendEmailSafe(mailOptions, 'Final Route B (to Production Editor): ' + (msData.MsVer || ''));
}

/**
 * ルートb: 著者への受理通知メール（添付なし）
 * DecisionシートのMail text + 委員長コメントを送付
 * BCC: 編集幹事 + 担当編集者
 */
function _notifyAuthorOfAcceptance(msData, data, settings, ssId) {
  var webAppUrl    = ScriptApp.getService().getUrl();
  var authorUrl    = webAppUrl + '?key=' + (msData.key || '');
  var eicComment   = data.eicAuthorComment || '';
  var resolvedSsId = ssId || getSpreadsheetId();

  // Decision Mail テンプレートを取得
  var decisionTemplates = getDecisionTemplates(resolvedSsId, data.decision || '');
  var replacements = {
    'authorName':         msData.CA_Name              || '',
    'englishTitle':       msData.TitleEN               || '',
    'Journal_Name':       settings.Journal_Name        || '',
    'Resubmittion_Limit': settings.Resubmittion_Limit  || '8 weeks',
    'manuscriptID':       msData.MsVer                 || '',
    'Editor_Name':        settings.Editor_Name         || 'Editor-in-Chief',
    'dueDate':            '',
    'submissionLink':     authorUrl,
    'formlink':           authorUrl
  };
  var templateText = replaceDecisionPlaceholders(decisionTemplates.mailText, replacements);

  var commentsHtml = eicComment
    ? '<p style="margin-top:1.5rem;"><strong>編集委員長よりのコメント / Comments from Editor-in-Chief:</strong><br>' +
      escHtml(eicComment).replace(/\n/g, '<br>') + '</p>'
    : '';

  var bodyHtml =
    '<p>A decision has been reached regarding your manuscript <strong>' + escHtml(msData.MsVer || '') + '</strong>.</p>' +
    '<div style="background:#f1f5f9; padding:20px; border-radius:8px; margin:20px 0; font-size:15px; line-height:1.6;">' +
      templateText +
    '</div>' +
    commentsHtml +
    '<hr style="border:none; border-top:1px solid #e2e8f0; margin:20px 0;">' +
    '<p>ご投稿いただいた原稿 <strong>' + escHtml(msData.MsVer || '') + '</strong> に対する判定をお送りいたします。</p>';

  var html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting:    'Dear Dr. ' + escHtml(msData.CA_Name || '') + ',',
    bodyHtml:    bodyHtml,
    footerHtml:  settings.mailFooter || ''
  });

  var subject = '[' + settings.Journal_Name + '] 原稿の審査結果について / Decision for your manuscript, ' + (msData.MsVer || '') +
    (decisionTemplates.shortExplanation ? ': ' + decisionTemplates.shortExplanation : '');

  // BCC: 編集幹事 + 担当編集者
  var bccList = [];
  if (settings.managingEditorEmail) bccList.push(settings.managingEditorEmail);
  var acceptedEditor = _getAcceptedEditorEmail(resolvedSsId, msData.MsVer || '');
  if (acceptedEditor && acceptedEditor.email) bccList.push(acceptedEditor.email);

  var mailOptions = {
    to:       msData.CA_Email || '',
    cc:       msData.ccEmails || '',
    subject:  subject,
    htmlBody: html
  };
  if (bccList.length > 0) mailOptions.bcc = bccList.join(', ');

  sendEmailSafe(mailOptions, 'Acceptance Notification to Author (Route B): ' + (msData.MsVer || ''));
}

/**
 * ルートc: 担当編集者への再判定依頼メール
 */
function _sendFinalRouteCToEditor(msData, data, eicFileUrl, settings, ssId) {
  var editorLogs = findAllRecordsByKey(ssId, EDITOR_LOG_SHEET_NAME, 'MsVer', msData.MsVer || '');
  var acceptedEditor = editorLogs.find(function(log) { return String(log.edtOk || '').trim() === 'ok'; });
  if (!acceptedEditor || !acceptedEditor.Editor_Email) {
    Logger.log('_sendFinalRouteCToEditor: 承諾済み担当編集者が見つかりません (MsVer: ' + msData.MsVer + ')');
    return;
  }

  var editorLink = ScriptApp.getService().getUrl() + '?editorKey=' + (acceptedEditor.editorKey || '');
  var eicComment = data.eicAuthorComment || '';

  var commentBlock = eicComment
    ? '<div style="margin:16px 0; padding:12px 16px; background:#fef3c7; border:1px solid #fbbf24; border-radius:8px;">' +
        '<p style="margin:0 0 6px; font-weight:bold; color:#92400e;">委員長よりのコメント / EIC\'s Comment:</p>' +
        '<p style="margin:0; white-space:pre-wrap; color:#78350f;">' + escHtml(eicComment) + '</p>' +
      '</div>'
    : '';

  var fileLink = eicFileUrl
    ? '<p><a href="' + eicFileUrl + '" target="_blank">【閲覧専用】委員長添付ファイル / EIC\'s Attached Files</a></p>'
    : '';

  var bodyHtml =
    '<p>The manuscript <strong>' + escHtml(msData.MsVer || '') + '</strong> has been returned to you by the Editor-in-Chief for re-evaluation.</p>' +
    commentBlock +
    fileLink +
    '<p>Please re-examine the manuscript and resubmit your recommendation.</p>' +
    '<hr style="border:none; border-top:1px solid #e2e8f0; margin:20px 0;">' +
    '<p>原稿 <strong>' + escHtml(msData.MsVer || '') + '</strong> について、編集委員長より再判定の依頼がありました。</p>' +
    '<p>上記コメントをご確認のうえ、改めて推薦内容をご提出ください。</p>';

  var html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting: 'Dear Dr. ' + escHtml(acceptedEditor.Editor_Name || '') + ',',
    bodyHtml: bodyHtml,
    buttonUrl: editorLink,
    buttonLabel: 'Open Editor Dashboard / 担当編集者ダッシュボードを開く',
    footerHtml: settings.mailFooter || ''
  });

  sendEmailSafe({
    to: acceptedEditor.Editor_Email,
    subject: '[' + settings.Journal_Name + '] 再判定依頼 / Re-evaluation Request: ' + (msData.MsVer || ''),
    htmlBody: html
  }, 'Final Route C (to Editor): ' + (msData.MsVer || ''));
}

/**
 * 委員長がどのルートに回したかを編集幹事へ通知
 */
function _notifyManagingEditorOfEicRoute(msData, route, eicAuthorComment, eicProductionComment, decision, settings, ssId) {
  if (!settings.managingEditorEmail) {
    Logger.log('_notifyManagingEditorOfEicRoute: managingEditorEmail 未設定のためスキップ');
    return;
  }

  // 判定ラベル（decision が指定されている場合はその値を、なければルート説明を使用）
  var routeLabels = {
    a: 'Returned to Author (Route A)',
    b: 'Sent to Production (Route B)',
    c: 'Returned to Editor for Re-evaluation (Route C)'
  };
  var decisionLabel = decision || routeLabels[route] || route;

  // Decisions シートの Mail text を取得（decision がある場合のみ）
  var mailTextHtml = '';
  if (decision && ssId) {
    try {
      var dt = getDecisionTemplates(ssId, decision);
      if (dt && dt.mailText) {
        var replacements = {
          'authorName':         msData.CA_Name              || '',
          'englishTitle':       msData.TitleEN               || '',
          'Journal_Name':       settings.Journal_Name        || '',
          'Resubmittion_Limit': settings.Resubmittion_Limit  || '8 weeks',
          'manuscriptID':       msData.MsVer                 || '',
          'Editor_Name':        settings.Editor_Name         || 'Editor-in-Chief',
          'dueDate':            '',
          'submissionLink':     ScriptApp.getService().getUrl() + '?key=' + (msData.key || ''),
          'formlink':           ScriptApp.getService().getUrl() + '?key=' + (msData.key || '')
        };
        var mailText = replaceDecisionPlaceholders(dt.mailText, replacements);
        mailTextHtml =
          '<div style="margin:20px 0;">' +
            '<p style="margin:0 0 8px; font-weight:bold; color:#475569; font-size:13px;">▼ 著者に送付したメール本文 / Email text sent to author:</p>' +
            '<div style="background:#f1f5f9; padding:16px; border-radius:8px; font-size:14px; line-height:1.7; white-space:pre-wrap;">' +
              mailText +
            '</div>' +
          '</div>';
      }
    } catch (e) {
      Logger.log('_notifyManagingEditorOfEicRoute: Mail text 取得エラー: ' + e.message);
    }
  }

  var bodyHtml =
    '<p>The Editor-in-Chief has taken action on the following manuscript.</p>' +
    '<p>編集委員長が以下の原稿に対してアクションを取りました。</p>' +
    '<table style="width:100%; font-size:14px; border-collapse:collapse; margin:20px 0;">' +
      '<tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee; width:30%;">原稿番号 / MS ID</th>' +
          '<td style="padding:8px; border-bottom:1px solid #eee;">' + escHtml(msData.MsVer || '') + '</td></tr>' +
      '<tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">判定 / デシジョン</th>' +
          '<td style="padding:8px; border-bottom:1px solid #eee; font-weight:bold;">' + escHtml(decisionLabel) + '</td></tr>' +
      (eicAuthorComment ? '<tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">委員長コメント（著者宛） / EIC\'s Comment (for Author)</th>' +
          '<td style="padding:8px; border-bottom:1px solid #eee; white-space:pre-wrap;">' + escHtml(eicAuthorComment) + '</td></tr>' : '') +
      (eicProductionComment ? '<tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">委員長コメント（印刷担当者宛） / EIC\'s Comment (for Production)</th>' +
          '<td style="padding:8px; border-bottom:1px solid #eee; white-space:pre-wrap;">' + escHtml(eicProductionComment) + '</td></tr>' : '') +
    '</table>' +
    mailTextHtml;

  var html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting: '編集幹事 殿 / Dear Managing Editor,',
    bodyHtml: bodyHtml,
    footerHtml: settings.mailFooter || ''
  });

  sendEmailSafe({
    to: settings.managingEditorEmail,
    subject: '[' + settings.Journal_Name + '] 委員長アクション通知 / EIC Action: ' + escHtml(decisionLabel) + ' (' + (msData.MsVer || '') + ')',
    htmlBody: html
  }, 'EIC Route Notification to Managing Editor: ' + (msData.MsVer || ''));
}

/**
 * 委員長による投稿直後の審査停止 API
 * 著者への通知メールをシステムから送信する
 * 停止後は apiArchiveManuscript で手動削除できる
 * data: { eicKey, message }
 */
function apiStopManuscriptByEic(data) {
  var ssId = getSpreadsheetId();
  var msData = getManuscriptData('eic', data.eicKey);
  if (!msData) throw new Error('原稿が見つかりません。/ Manuscript not found.');

  var now = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss');
  updateLogCell(ssId, MANUSCRIPTS_SHEET_NAME, 'key', msData.key, {
    'stoppedByEicAt': now
  });

  // 著者への通知メール送信
  if (data.message) {
    var caEmail = String(msData.CA_Email || '').trim();
    if (!caEmail) {
      // CA_Email が未設定の場合はメールをスキップしてログに記録し、DB更新は維持する
      writeLog('EIC Early Rejection: CA_Email missing, email skipped for ' + (msData.MsVer || ''));
    } else {
      var settings = getSettings();
      var journalName = (settings && settings.Journal_Name) ? settings.Journal_Name : 'Journal';
      var msVer = msData.MsVer || '';
      var bodyHtml = String(data.message)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
        .replace(/\n/g, '<br>');
      var html = renderRichEmail({
        journalName: journalName,
        greeting:    '',
        bodyHtml:    bodyHtml,
        footerHtml:  settings.mailFooter || ''
      });
      sendEmailSafe({
        to:       caEmail,
        subject:  '[' + journalName + '] 原稿の掲載見合わせについて / Manuscript Not Accepted: ' + msVer,
        htmlBody: html
      }, 'EIC Early Rejection Notification: ' + msVer);
    }
  }

  writeLog('EIC Stopped Manuscript (Early Rejection): ' + (msData.MsVer || ''));
  return { success: true };
}

/**
 * 原稿タイトルを日英で結合するユーティリティ
 */
function _buildPaperTitle(msData) {
  var jp = msData.TitleJP || '';
  var en = msData.TitleEN || '';
  return (jp && en) ? jp + ' / ' + en : (jp || en || '');
}

/**
 * 承諾済み担当編集者のメールアドレスと氏名を返す
 * @returns {{ email: string, name: string } | null}
 */
function _getAcceptedEditorEmail(ssId, msVer) {
  if (!msVer) return null;
  var editorLogs = findAllRecordsByKey(ssId, EDITOR_LOG_SHEET_NAME, 'MsVer', msVer);
  var acceptedEditor = editorLogs.find(function(log) {
    return String(log.edtOk || '').trim() === 'ok';
  });
  if (!acceptedEditor) return null;
  return {
    email: String(acceptedEditor.Editor_Email || '').trim(),
    name:  String(acceptedEditor.Editor_Name  || '').trim()
  };
}

/**
 * Manuscripts シートのフィールド名を generateReceiptPdf が期待する形式にマッピング
 */
function _mapMsDataForReceipt(msData) {
  return {
    MsVer:             msData.MsVer            || '',
    authorName:        msData.CA_Name          || '',
    authorEmail:       msData.CA_Email         || '',
    ccEmails:          msData.ccEmails         || '',
    paperType:         msData.MS_Type          || '',
    titleJp:           msData.TitleJP          || '',
    titleEn:           msData.TitleEN          || '',
    runningTitle:      msData.RunningTitle     || '',
    submittedFiles:    msData.submittedFiles   || '',
    sendDateTime:      msData.Submitted_At     || '',
    reprintRequest:    msData['Reprint request'] || '',
    englishEditing:    msData['English_editing'] || '',
    authorAffiliation: msData.CA_Affiliation   || '',
    authorsJp:         msData.AllAuthors_JP    || msData.Authors_JP       || '',
    authorsEn:         msData.AllAuthors_EN    || msData.Authors_EN       || ''
  };
}
