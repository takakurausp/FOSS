/**
 * RecommendationModule.js - 担当編集者から委員長への推薦 (SPA)
 */

function apiSubmitRecommendation(data) {
  Logger.log('apiSubmitRecommendation: start. editorKey=' + data.editorKey + ' score=' + data.score + ' files=' + (data.files ? data.files.length : 0));
  const ssId = getSpreadsheetId();
  const settings = getSettings();

  // 1. 担当編集者のキーから原稿・編集者データを取得
  const msData = getManuscriptData('editor', data.editorKey);
  if (!msData) throw new Error("Editor record not found.");
  Logger.log('apiSubmitRecommendation: msData loaded. MsVer=' + msData.MsVer + ' MS_ID=' + msData.MS_ID);

  const hexId = msData.MsVerRevHex;
  const msVer = msData.MsVer;
  const msId = msData.MS_ID;
  const now = new Date();
  const todayNow = Utilities.formatDate(now, 'JST', 'yyyy/MM/dd HH:mm');

  // 2. 査読報告書（推薦書）の作成 — Drive 操作を先に完了させてから DB を更新する。
  // createRecommendationReport が例外を投げても DB は未変更のままなので不整合が生じない。
  Logger.log('apiSubmitRecommendation: calling createRecommendationReport...');
  const reportFiles = createRecommendationReport(msData, data, ssId, settings);
  Logger.log('apiSubmitRecommendation: createRecommendationReport done.');

  // 3. 全 Drive 操作が成功した後に Editor_log をまとめて更新
  updateLogCell(ssId, EDITOR_LOG_SHEET_NAME, 'editorKey', data.editorKey, {
    'Score':               data.score,
    'Received_At':         todayNow,
    'Message':             data.openComments         || '',
    'ConfidentialMessage': data.confidentialComments || '',
    'reportPdfUrl':              reportFiles.pdf.getUrl(),
    'reportWordUrl':             reportFiles.word.getUrl(),
    'reportFolderUrl':           reportFiles.folderUrl           || '',
    'reportAttachmentsFolderUrl': reportFiles.attachmentsFolderUrl || '',
    'reportGoogleDocId':         reportFiles.googleDocId         || ''
  });

  // 4. 受理スコアかどうかで通知先を分岐
  //    条件A（受理）: managingEditorKey を生成して編集幹事に通知
  //    条件B（その他）: 既存フロー通り委員長に通知
  const isAccepted = isScoreAccepted(ssId, data.score);
  writeLog(`ルーティング判定: score="${data.score}" isAccepted=${isAccepted} → ${isAccepted ? '編集幹事ルート (条件A)' : '委員長直通ルート (条件B)'}`);

  if (isAccepted) {
    // 条件A: 編集幹事ルート
    if (!settings.managingEditorEmail) {
      // 編集幹事メール未設定の場合はエラーを記録して処理を中断
      writeLog('[ERROR] Recommendation Routing: score="' + data.score + '" は受理スコアですが、Settings に managingEditorEmail が設定されていません。編集幹事への通知ができません。Settings シートに managingEditorEmail を設定してください。');
      throw new Error('managingEditorEmail が Settings に設定されていません。受理原稿の通知先として編集幹事のメールアドレスを設定してください。');
    }
    const managingEditorKey = Utilities.getUuid();
    // Manuscripts シートに managingEditorKey と finalStatus を記録
    updateLogCell(ssId, MANUSCRIPTS_SHEET_NAME, 'key', msData.key, {
      'managingEditorKey': managingEditorKey,
      'finalStatus':       'final_review'
    });
    sendRecommendationToManagingEditor(msData, data, reportFiles, settings, managingEditorKey);
  } else {
    // 条件B: 既存フロー（委員長へ直接通知）
    sendRecommendationToChiefEditor(msData, data, reportFiles, settings);
  }

  writeLog(`Recommendation Submitted: ${msVer} by ${msData.Editor_Name} - Score: ${data.score} (${isAccepted ? 'Accepted→ManagingEditor' : 'NotAccepted→EIC'})`);

  return { success: true };
}

/**
 * 査読報告書の作成 (PDF 全史レポート + Word オープンコメント集)
 */
function createRecommendationReport(msData, data, ssId, settings) {
  Logger.log('createRecommendationReport: getting verFolder for MS_ID=' + msData.MS_ID);
  const folder = getManuscriptVerFolder(msData, settings);
  Logger.log('createRecommendationReport: verFolder OK. getting workingFolder...');
  const workingFolder = driveFolderCache.getOrCreateFolder(folder, 'working');
  Logger.log('createRecommendationReport: workingFolder OK.');

  const journalName = (settings && settings.Journal_Name) ? settings.Journal_Name : 'Journal';
  const now = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm');
  const esc = s => String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  const nl2br = s => esc(s).replace(/\n/g, '<br>');

  // 1. 査読者コメントの収集
  const reviewLogLines = getFilteredReviewLog(ssId, msData.MsVer);

  // 2. 担当編集者がアップロードしたファイルを保存
  //    添付ファイルは attachments/ サブフォルダに分離し、
  //    コメント系ファイル（PDF/Word/Google Docs）が混在する workingFolder は直接見せない
  const uploadedFiles = [];
  let attachmentsFolderUrl = '';
  if (data.files && data.files.length > 0) {
    const attachFolder = driveFolderCache.getOrCreateFolder(workingFolder, 'attachments');
    data.files.forEach(file => {
      const blob = Utilities.newBlob(Utilities.base64Decode(file.content), file.mimeType, file.name);
      attachFolder.createFile(blob);
      uploadedFiles.push(file.name);
    });
    attachFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    attachmentsFolderUrl = attachFolder.getUrl();
  }

  // ─────────────────────────────────────────────────────────────
  // 3. Word ファイル作成 (オープンコメントのみ / Open Comments only)
  // ─────────────────────────────────────────────────────────────
  const wordDoc  = DocumentApp.create('Open-Comments-' + msData.MsVer);
  const wordBody = wordDoc.getBody();
  wordBody.setMarginTop(72).setMarginBottom(72).setMarginLeft(85).setMarginRight(85);

  // スタイル定義
  const styleTitle = {};
  styleTitle[DocumentApp.Attribute.FONT_SIZE]        = 12;
  styleTitle[DocumentApp.Attribute.BOLD]             = true;
  styleTitle[DocumentApp.Attribute.FONT_FAMILY]      = 'Arial';

  const styleBody = {};
  styleBody[DocumentApp.Attribute.FONT_SIZE]         = 10.5;
  styleBody[DocumentApp.Attribute.BOLD]              = false;
  styleBody[DocumentApp.Attribute.FONT_FAMILY]       = 'Arial';

  const styleLabel = {};
  styleLabel[DocumentApp.Attribute.FONT_SIZE]        = 10.5;
  styleLabel[DocumentApp.Attribute.BOLD]             = true;
  styleLabel[DocumentApp.Attribute.FONT_FAMILY]      = 'Arial';

  // ─── 原稿情報ブロック
  // 注: 雑誌名・発行日時・判定は PDF 化時に先頭へ動的付加するため、ここには含めない
  wordBody.appendParagraph('');
  const msInfoStyle = Object.assign({}, styleBody);

  const msNoLine = wordBody.appendParagraph('原稿番号 / Manuscript No.: ' + (msData.MsVer || ''));
  msNoLine.setAttributes(styleBody);

  const titleJP = msData.TitleJP || '';
  const titleEN = msData.TitleEN || '';
  const titleText = titleJP && titleEN ? titleJP + '\n' + titleEN : (titleJP || titleEN || '');
  const titleLine = wordBody.appendParagraph('論文タイトル / Title: ' + titleText);
  titleLine.setAttributes(styleBody);

  const authorsJP = msData.AuthorsJP || '';
  const authorsEN = msData.AuthorsEN || '';
  const hasFullAuthors = authorsJP || authorsEN;
  const authorsText = hasFullAuthors
    ? (authorsJP && authorsEN ? authorsJP + ' / ' + authorsEN : (authorsJP || authorsEN))
    : (msData.CA_Name || '');
  const authLabel = hasFullAuthors ? '著者 / Authors' : '責任著者 / Corresponding Author';
  const authLine = wordBody.appendParagraph(authLabel + ': ' + authorsText);
  authLine.setAttributes(styleBody);

  const msTypeLine = wordBody.appendParagraph('論文種別 / Manuscript Type: ' + (msData.MS_Type || ''));
  msTypeLine.setAttributes(styleBody);

  // ─── 冒頭文
  wordBody.appendParagraph('');
  const introText =
    'このたびはご投稿いただきありがとうございます。' +
    'ご投稿原稿に対して査読を行いました結果、編集委員および査読者より以下のコメントが寄せられましたので、お知らせいたします。' +
    '別途添付ファイルがございます場合は、あわせてご参照ください。' +
    '改訂稿をご提出いただく場合は、各コメントに対する回答書を作成のうえ、改訂稿とともにご送付くださいますようお願いいたします。\n\n' +
    'Thank you for submitting your manuscript to our journal. ' +
    'Following a peer review, we are pleased to share the comments provided by the editors and reviewers below. ' +
    'If any attachments are included separately, please refer to them as well. ' +
    'Should you wish to submit a revised manuscript, please prepare a point-by-point response to the comments and submit it together with your revised manuscript.';
  const introPara = wordBody.appendParagraph(introText);
  introPara.setAttributes(styleBody);
  introPara.setSpacingAfter(12);

  wordBody.appendHorizontalRule();

  // ─── Section 1: 担当編集者のオープンコメント
  wordBody.appendParagraph('');
  const edHeading = wordBody.appendParagraph('Section 1: Responsible Editor\'s Comments / 担当編集者のコメント');
  edHeading.setAttributes(styleLabel);
  edHeading.setSpacingAfter(4);

  const edBody = wordBody.appendParagraph(data.openComments || '(No comments)');
  edBody.setAttributes(styleBody);

  // ─── Section 2: 査読者のオープンコメント
  wordBody.appendParagraph('');
  wordBody.appendParagraph('');
  const revSectionHeading = wordBody.appendParagraph('Section 2: Reviewer Comments / 査読者コメント');
  revSectionHeading.setAttributes(styleLabel);
  revSectionHeading.setSpacingAfter(4);

  reviewLogLines.forEach((rev, idx) => {
    wordBody.appendParagraph('');
    const revHead = wordBody.appendParagraph('Reviewer #' + (idx + 1));
    revHead.setAttributes(styleLabel);
    revHead.setSpacingAfter(4);
    const revBody = wordBody.appendParagraph(rev.openCommentsText || '(No comments)');
    revBody.setAttributes(styleBody);
    if (idx < reviewLogLines.length - 1) {
      wordBody.appendParagraph('');
    }
  });

  Logger.log('createRecommendationReport: saving Word doc...');
  wordDoc.saveAndClose();
  Logger.log('createRecommendationReport: exporting as DOCX via getDocxBlob...');
  const wordBlob = getDocxBlob(wordDoc.getId());
  Logger.log('createRecommendationReport: DOCX export OK. saving to workingFolder...');
  const wordFile = workingFolder.createFile(wordBlob).setName('Editor-Comments-' + msData.MsVer + '.docx');
  wordFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  // Google Doc を削除せず保持し、EDIT 共有を設定して workingFolder へ移動する。
  // EIC がブラウザ上で直接編集し、著者への送付に利用できるようにする。
  const wordDocFile = DriveApp.getFileById(wordDoc.getId());
  wordDocFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
  wordDocFile.moveTo(workingFolder);

  // ─────────────────────────────────────────────────────────────
  // 4. PDF 全史レポート作成 (コンフィデンシャルコメント含む全情報)
  // ─────────────────────────────────────────────────────────────
  const fmtDate = v => v instanceof Date ? Utilities.formatDate(v, 'JST', 'yyyy/MM/dd HH:mm') : String(v || '');
  const sectionTitle = (title) =>
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

  let reviewerSections = '';
  reviewLogLines.forEach((rev, idx) => {
    reviewerSections += `
      <div style="margin-bottom:16px; padding:10px 14px; border:1px solid #e2e8f0; border-radius:8px; break-inside:avoid;">
        <p style="margin:0 0 6px; font-size:13px; font-weight:bold;">Reviewer #${idx + 1}: ${esc(rev.Rev_Name)}</p>
        <table style="width:100%; border-collapse:collapse; margin-bottom:8px;">
          ${infoRow('判定スコア', 'Score', rev.Score || '')}
          ${infoRow('査読結果提出日', 'Submitted', fmtDate(rev.Received_At))}
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
    <p style="color:#bfdbfe; margin:4px 0 0; font-size:12px;">Peer Review Report / 査読報告書</p>
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
    ${infoRow('担当編集者', 'Responsible Editor', (msData.Editor_Name || '') + (msData.Editor_Email ? ' (' + msData.Editor_Email + ')' : ''))}
    ${infoRowHtml('推薦スコア', 'Recommended Score', `<strong style="color:#1e40af; font-size:14px;">${esc(data.score || '')}</strong>`)}
    ${infoRow('推薦提出日時', 'Submitted at', now)}
  </table>
  ${commentBox('オープンコメント / Open Comments (for authors)', data.openComments || '', '#f8fafc', '#e2e8f0')}
  ${commentBox('🔒 コンフィデンシャルコメント / Confidential Comments (for EIC only)', data.confidentialComments || '', '#fffbeb', '#fcd34d')}

  ${uploadedFiles.length > 0 ? sectionTitle('添付ファイル / Attached Files (' + uploadedFiles.length + ' files)') : ''}
  ${uploadedFiles.length > 0 ? `
    <div style="margin-bottom:16px; padding:10px 14px; border:1px solid #e2e8f0; border-radius:8px; break-inside:avoid;">
      <p style="margin:0 0 6px; font-size:13px; font-weight:bold;">担当編集者が追加したファイル / Files uploaded by responsible editor:</p>
      <ul style="margin:0; padding-left:20px;">
        ${uploadedFiles.map(file => `<li style="margin-bottom:4px; font-size:12px;">${esc(file)}</li>`).join('')}
      </ul>
    </div>
  ` : ''}

  ${reviewLogLines.length > 0 ? sectionTitle('査読結果 / Peer Review Results (' + reviewLogLines.length + ' reviewers)') : ''}
  ${reviewerSections}
</body>
</html>`;

  Logger.log('createRecommendationReport: generating PDF...');
  const pdfBlob = HtmlService.createHtmlOutput(pdfHtml).getBlob().getAs(MimeType.PDF);
  Logger.log('createRecommendationReport: PDF generated. saving to workingFolder...');
  const pdfFile = workingFolder.createFile(pdfBlob).setName('Editor-Report-' + msData.MsVer + '.pdf');
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    pdf:                 pdfFile,
    word:                wordFile,
    uploadedFiles:       uploadedFiles,
    folderUrl:           workingFolder.getUrl(),         // 内部用（コメント含む）
    attachmentsFolderUrl: attachmentsFolderUrl,          // EIC・ME向け（添付のみ）
    googleDocId:         wordDoc.getId()
  };
}

/**
 * 査読ログから特定の原稿の「受諾済み且つ回答済み」のレコードを抽出
 */
function getFilteredReviewLog(ssId, msVer) {
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(REVIEW_LOG_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const msVerIdx = headers.indexOf('MsVer');
  const revOkIdx = headers.indexOf('revOk');
  const scoreIdx = headers.indexOf('Score');
  const openDocIdIdx = headers.indexOf('openCommentsId');
  const revNameIdx = headers.indexOf('Rev_Name');

  const confDocIdIdx = headers.indexOf('confidentialCommentsId');

  const results = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][msVerIdx] === msVer && data[i][revOkIdx] === 'ok' && data[i][scoreIdx] !== '') {
      const revData = {};
      headers.forEach((h, j) => revData[h] = data[i][j]);

      if (data[i][openDocIdIdx]) {
        try {
          revData.openCommentsText = DocumentApp.openById(data[i][openDocIdIdx]).getBody().getText().trim();
        } catch(e) {
          revData.openCommentsText = '(Error reading comments)';
        }
      }
      if (confDocIdIdx !== -1 && data[i][confDocIdIdx]) {
        try {
          revData.confidentialCommentsText = DocumentApp.openById(data[i][confDocIdIdx]).getBody().getText().trim();
        } catch(e) {
          revData.confidentialCommentsText = '';
        }
      }
      results.push(revData);
    }
  }
  return results;
}

/**
 * 条件A: 編集幹事への通知（受理スコアの場合）
 */
function sendRecommendationToManagingEditor(msData, data, reportFiles, settings, managingEditorKey) {
  // managingEditorEmail の存在チェックは呼び出し元 (apiSubmitRecommendation) で実施済み

  const webAppUrl = ScriptApp.getService().getUrl();
  const meLink = webAppUrl + '?managingEditorKey=' + managingEditorKey;
  const paperTitle = (msData.TitleJP && msData.TitleEN)
    ? msData.TitleJP + ' / ' + msData.TitleEN
    : (msData.TitleJP || msData.TitleEN || '');

  const bodyHtml = `
    <p>Responsible editor <strong>${msData.Editor_Name}</strong> has submitted an acceptance recommendation for the following manuscript. Please open the Managing Editor dashboard using the button below to complete your review.</p>
    <p>担当編集者 <strong>${msData.Editor_Name}</strong> より、以下の原稿の判定案（受理推薦）が提出されました。以下のボタンより編集幹事ダッシュボードを開き、最終確認作業をお願いいたします。</p>
    <table style="width:100%; font-size:14px; border-collapse:collapse; margin:20px 0;">
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee; width:30%;">原稿番号 / MS ID</th>
          <td style="padding:8px; border-bottom:1px solid #eee;">${msData.MsVer}</td></tr>
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">タイトル / Title</th>
          <td style="padding:8px; border-bottom:1px solid #eee;">${paperTitle}</td></tr>
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">推薦スコア / Recommended Score</th>
          <td style="padding:8px; border-bottom:1px solid #eee; font-weight:bold;">${data.score}</td></tr>
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">責任著者 / Corresponding Author</th>
          <td style="padding:8px; border-bottom:1px solid #eee;">${msData.CA_Name || ''}</td></tr>
    </table>
  `;

  const html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting: '編集幹事 殿 / Dear Managing Editor,',
    bodyHtml: bodyHtml,
    buttonUrl: meLink,
    buttonLabel: '編集幹事ダッシュボードを開く / Open Managing Editor Dashboard',
    footerHtml: settings.mailFooter || ''
  });

  sendEmailSafe({
    to: settings.managingEditorEmail,
    subject: `[${settings.Journal_Name}] 受理原稿の最終確認依頼 / Final Review Request: ${msData.MsVer}`,
    htmlBody: html,
    attachments: [reportFiles.pdf.getAs(MimeType.PDF), reportFiles.word.getBlob()]
  }, 'Recommendation (Accepted) to Managing Editor: ' + msData.MsVer);

  // ※ 編集委員長への通知は、編集幹事がチェックを完了して送信した後に
  //   _sendManagingEditorReviewToEIC() によって行われる。ここでは送らない。
}

/**
 * 委員長へ通知（条件B: 非受理スコアの場合の既存フロー）
 */
function sendRecommendationToChiefEditor(msData, data, reportFiles, settings) {
  const uploadedFiles = reportFiles.uploadedFiles || [];
  const folderUrl     = reportFiles.attachmentsFolderUrl || '';
  const webAppUrl = ScriptApp.getService().getUrl();
  const decisionLink = webAppUrl + '?eicKey=' + msData.eicKey;

  const paperTitle = (msData.TitleJP && msData.TitleEN) ? msData.TitleJP + ' / ' + msData.TitleEN : (msData.TitleJP || msData.TitleEN || '');

  const folderButtonHtml = (uploadedFiles.length > 0 && folderUrl) ? `
    <div style="margin:16px 0;">
      <a href="${folderUrl}" target="_blank"
         style="display:inline-block; background:#1d4ed8; color:#ffffff; text-decoration:none;
                padding:10px 20px; border-radius:6px; font-size:14px; font-weight:600;">
        📁 担当編集者の添付ファイルを開く / Open Editor's Uploaded Files
      </a>
      <p style="margin:6px 0 0; font-size:12px; color:#6b7280;">
        添付ファイル数 / Number of files: ${uploadedFiles.length}件
      </p>
    </div>` : '';

  const bodyHtml = `
    <p>Responsible editor <strong>${msData.Editor_Name}</strong> has sent a recommendation for the following manuscript. The peer review results and summary are attached. Please review these results and send the final decision to the authors by clicking the button below.</p>
    <p>担当編集者 <strong>${msData.Editor_Name}</strong> 殿より、以下の原稿の判定案（推薦）が提出されました。添付の資料を確認し、著者への最終通知（判定）を行ってください。</p>
    ${folderButtonHtml}
    <table style="width:100%; font-size: 14px; border-collapse: collapse; margin: 20px 0;">
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee; width: 30%;">ID</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${msData.MsVer}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Type</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${msData.MS_Type || ''}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee; width: 30%;">Title</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${paperTitle}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Corresponding Author</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${msData.CA_Name || ''}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Recommended Score</th><td style="padding: 8px; border-bottom: 1px solid #eee; font-weight:bold;">${data.score || ''}</td></tr>
    </table>
  `;

  const html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting: `Dear Editor-in-Chief,`,
    bodyHtml: bodyHtml,
    buttonUrl: decisionLink,
    buttonLabel: 'Send Decision to Author / 著者に判定を送る',
    footerHtml: settings.mailFooter || ''
  });

  if (!settings.chiefEditorEmail) {
    Logger.log('sendRecommendationToChiefEditor: chiefEditorEmail が設定されていないためスキップします (MsVer: ' + msData.MsVer + ')');
    return;
  }

  sendEmailSafe({
    to: settings.chiefEditorEmail,
    subject: `[${settings.Journal_Name}] 推薦レポート受領 / Recommendation Received: ${msData.MsVer}`,
    htmlBody: html,
    attachments: [reportFiles.pdf.getAs(MimeType.PDF), reportFiles.word.getBlob()]
  }, 'Recommendation to EIC: ' + msData.MsVer);
}

/**
 * 担当編集者へ受領確認
 */
function sendRecommendationConfirmationToEditor(msData, data, submittedAt, settings) {
  const uploadedFiles = data.files ? data.files.map(f => f.name) : [];
  const subject = `[${settings.Journal_Name}] 推薦レポート送信の確認 / Confirmation: Recommendation Submitted: ${msData.MsVer}`;

  const titleJP = msData.TitleJP || '';
  const titleEN = msData.TitleEN || '';
  const titleCell = titleJP && titleEN ? `${titleJP}<br>${titleEN}` : titleJP || titleEN || '';

  const dashboardUrl = msData.editorKey
    ? ScriptApp.getService().getUrl() + '?editorKey=' + msData.editorKey
    : null;

  const bodyHtml = `
    <p>Thank you for sending your recommendation for the following manuscript. We have successfully received your submission and notified the Editor-in-Chief. You can verify the submission date and time by opening your editor dashboard using the button below.</p>
    <p>原稿 <strong>${msData.MsVer}</strong> の判定案（推薦）をご送付いただき、ありがとうございます。内容を確かに受領し、編集委員長へ通知いたしました。以下のボタンから担当編集者ダッシュボードを開くと、推薦の送信日時をご確認いただけます。</p>
    <table style="width:100%; font-size: 14px; border-collapse: collapse; margin: 20px 0;">
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee; width: 30%;">ID</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${msData.MsVer}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Type</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${msData.MS_Type || ''}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee; width: 30%;">Title</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${titleCell}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Corresponding Author</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${msData.CA_Name || ''}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Recommended Score</th><td style="padding: 8px; border-bottom: 1px solid #eee; font-weight:bold;">${data.score || ''}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Submitted at</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${submittedAt}</td></tr>
    </table>
    ${uploadedFiles.length > 0 ? `
      <p><strong>Files you uploaded / アップロードしたファイル:</strong></p>
      <ul style="margin:0 0 16px 20px; padding:0;">
        ${uploadedFiles.map(file => `<li style="margin-bottom:4px;">${file}</li>`).join('')}
      </ul>
    ` : ''}
  `;

  const html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting: `Dear Dr. ${msData.Editor_Name},`,
    bodyHtml: bodyHtml,
    buttonUrl:   dashboardUrl || undefined,
    buttonLabel: dashboardUrl ? 'Open Editor Dashboard / 担当編集者ダッシュボードを開く' : undefined,
    footerHtml: settings.mailFooter || ''
  });

  sendEmailSafe({ to: msData.Editor_Email, subject, htmlBody: html },
    'Recommendation Confirmation to Editor: ' + msData.MsVer);
}

/**
 * 原稿のバージョンフォルダを取得
 */
function getManuscriptVerFolder(msData, settings) {
  const rootName = settings.SUBFOLDER || 'Journal Files';
  const root = driveFolderCache.getRootFolder(rootName);
  if (!root) throw new Error('Root folder not found: ' + rootName);
  const msFolder = driveFolderCache.getFolderByName(root, msData.MS_ID);
  if (!msFolder) throw new Error('Manuscript folder not found: ' + msData.MS_ID);
  const verNo = msData.Ver_No || 1;
  const verFolder = driveFolderCache.getFolderByName(msFolder, 'ver.' + verNo);
  if (!verFolder) throw new Error('Version folder not found: ver.' + verNo);
  return verFolder;
}

/**
 * Google Docs ID を Docx の Blob に変換
 * UrlFetchApp の失敗（ネットワークエラー・HTTP エラー）は呼び出し元に伝播させる。
 * muteHttpExceptions: true により非 200 レスポンスも例外なく受け取り、
 * レスポンスコードを確認して明確なエラーメッセージを throw する。
 */
function getDocxBlob(docId) {
  const url = 'https://docs.google.com/feeds/download/documents/export/Export?id=' + docId + '&exportFormat=docx';
  let response;
  try {
    response = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
  } catch (e) {
    throw new Error('Word ファイルのエクスポートに失敗しました（ネットワークエラー）: ' + e.message);
  }
  const code = response.getResponseCode();
  if (code !== 200) {
    const preview = response.getContentText().substring(0, 200);
    throw new Error('Word ファイルのエクスポートに失敗しました（HTTP ' + code + '）: ' + preview);
  }
  return response.getBlob();
}

/**
 * 受理後の修正版が再投稿された際の編集幹事への通知
 */
function sendResubmittedAcceptedNotificationToManagingEditor(msData, settings, managingEditorKey) {
  const webAppUrl = ScriptApp.getService().getUrl();
  const meLink = webAppUrl + '?managingEditorKey=' + managingEditorKey;
  const paperTitle = (msData.titleJp && msData.titleEn)
    ? msData.titleJp + ' / ' + msData.titleEn
    : (msData.titleJp || msData.titleEn || '');

  const bodyHtml = `
    <p>A revised version of the provisionally accepted manuscript <strong>${msData.MsVer}</strong> has been resubmitted by the author. Since the peer review steps are already complete, this manuscript has been routed directly to the Managing Editor's final review flow. Please open the Managing Editor dashboard using the button below to review the files and forward them to the Editor-in-Chief.</p>
    <p>受理内定済み（Provisional Accept）の原稿 <strong>${msData.MsVer}</strong> について、著者より修正版が再投稿されました。担当編集者による査読ステップは完了しているため、本原稿は直接編集幹事の最終確認フローへ回されました。以下のボタンより編集幹事ダッシュボードを開き、内容の確認と委員長への回送をお願いいたします。</p>
    <table style="width:100%; font-size:14px; border-collapse:collapse; margin:20px 0;">
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee; width:30%;">原稿番号 / MS ID</th>
          <td style="padding:8px; border-bottom:1px solid #eee;">${msData.MsVer}</td></tr>
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">タイトル / Title</th>
          <td style="padding:8px; border-bottom:1px solid #eee;">${paperTitle}</td></tr>
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">責任著者 / Corresponding Author</th>
          <td style="padding:8px; border-bottom:1px solid #eee;">${msData.authorName || ''} (${msData.authorEmail || ''})</td></tr>
    </table>
  `;

  const html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting: '編集幹事 殿 / Dear Managing Editor,',
    bodyHtml: bodyHtml,
    buttonUrl: meLink,
    buttonLabel: '編集幹事ダッシュボードを開く / Open Managing Editor Dashboard',
    footerHtml: settings.mailFooter || ''
  });

  sendEmailSafe({
    to: settings.managingEditorEmail,
    subject: `[${settings.Journal_Name}] 修正版の受領通知（受理済み原稿）/ Revised Accepted Manuscript: ${msData.MsVer}`,
    htmlBody: html
  }, 'Resubmitted Accepted Manuscript Notification to Managing Editor: ' + msData.MsVer);
}
