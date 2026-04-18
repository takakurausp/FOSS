/**
 * ReviewSubmissionModule.js - 査読結果受付バックエンド (SPA)
 */

function apiSubmitReview(data) {
  const ssId = getSpreadsheetId();
  const settings = getSettings();
  
  // 1. Get manuscript data using the reviewer key
  const msData = getManuscriptData('reviewer', data.reviewKey);
  if (!msData) throw new Error("Review record not found.");
  
  const hexId = msData.MsVerRevHex;
  const msVer = msData.MsVer;
  const msId = msData.MS_ID;
  const todayNow = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm');

  // 2 & 3. Drive/Docs 操作を先に完了させてから DB を更新する。
  // Drive/Docs で例外が発生しても DB は未変更のままなので不整合が生じない。

  // 2. Create a subfolder for this review
  const rootName = settings.SUBFOLDER || 'Journal Files';
  const root = driveFolderCache.getRootFolder(rootName) || DriveApp.createFolder(rootName);
  const msFolder = driveFolderCache.getOrCreateFolder(root, msId);
  const verNo = msData.Ver_No || 1;
  const verFolder = driveFolderCache.getOrCreateFolder(msFolder, 'ver.' + verNo);

  const reviewerName = msData.Rev_Name || 'Unknown_Reviewer';
  // 査読者フォルダ（コメントGDocsはここに保存 / 外部には共有しない）
  const revFolder = driveFolderCache.getOrCreateFolder(verFolder, reviewerName);

  // 3. Save attachments — 添付ファイルは attachments/ サブフォルダに分離
  //    EIC・編集幹事には attachments/ フォルダのリンクのみ表示し、
  //    コメントGDocsが混在する revFolder を直接見せないようにする
  let reviewFolderUrl = 'nofile';
  if (data.files && data.files.length > 0) {
    const attachFolder = driveFolderCache.getOrCreateFolder(revFolder, 'attachments');
    data.files.forEach(file => {
      const blob = Utilities.newBlob(Utilities.base64Decode(file.content), file.mimeType, file.name);
      attachFolder.createFile(blob);
    });
    attachFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    reviewFolderUrl = attachFolder.getUrl();
  }

  // 4. Save comments to Google Docs — コメントは revFolder（添付フォルダの親）に保存
  //    共有リンクは渡さないため EIC・編集幹事からは見えない
  const openDoc = DocumentApp.create('Open-Comments-' + hexId);
  openDoc.getBody().appendParagraph(data.openComments || '');
  const confDoc = DocumentApp.create('Confidential-Comments-' + hexId);
  confDoc.getBody().appendParagraph(data.confidentialComments || '');

  // Move docs to revFolder (not the attachments subfolder)
  DriveApp.getFileById(openDoc.getId()).moveTo(revFolder);
  DriveApp.getFileById(confDoc.getId()).moveTo(revFolder);

  // 5. Drive/Docs が全て成功した後に DB をまとめて更新
  // updateReviewLogCells で1回読み込み・1回書き込みにまとめる（旧: 5回個別呼び出し）
  updateReviewLogCells(ssId, hexId, {
    'Received_At':             todayNow,
    'Score':                   data.score,
    'reviewerUploadFolderUrl': reviewFolderUrl,
    'openCommentsId':          openDoc.getId(),
    'confidentialCommentsId':  confDoc.getId()
  });
  
  // 7. Send Email to Editor
  sendReviewResultToEditor(msData, data, reviewFolderUrl, settings, ssId);
  
  return { success: true };
}

/**
 * ReviewLog シートの対象行を一括更新する
 * @param {string} ssId   - スプレッドシート ID
 * @param {string} hexId  - 更新対象行の MsVerRevHex 値
 * @param {Object} updates - { 列名: 値, ... } の形式で更新内容を渡す
 *
 * シートを1回だけ読み込み、対象行の配列を更新してから行全体を1回の setValues で
 * 書き戻すことで、旧実装（列ごとに個別呼び出し）の複数回フルスキャンを解消する。
 */
function updateReviewLogCells(ssId, hexId, updates) {
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(REVIEW_LOG_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const hexIdx = headers.indexOf('MsVerRevHex');
  if (hexIdx === -1) return;

  // data[0] はヘッダー行なので i > 0 から探索する
  const rowIdx = data.findIndex((r, i) => i > 0 && r[hexIdx] === hexId);
  if (rowIdx <= 0) return;

  // 対象行を配列としてコピーし、更新フィールドだけ上書きしてから行全体を一括書き込み
  const rowData = data[rowIdx].slice();
  Object.entries(updates).forEach(function([colName, value]) {
    const colIdx = headers.indexOf(colName);
    if (colIdx !== -1) rowData[colIdx] = value;
  });
  sheet.getRange(rowIdx + 1, 1, 1, rowData.length).setValues([rowData]);
}

function sendReviewResultToEditor(msData, data, reviewFolderUrl, settings, ssId) {
  // 承諾済み査読者のうち何名が結果を提出したか集計
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(REVIEW_LOG_SHEET_NAME);
  const rows = sheet.getDataRange().getValues();

  // 大文字小文字を区別しないヘッダー検索（updateLogCell と同方式）
  const headers    = rows[0].map(h => String(h).toLowerCase().trim());
  const colIdx     = name => headers.indexOf(name.toLowerCase().trim());
  const msVerCol   = colIdx('MsVer');
  const edEmailCol = colIdx('Editor_Email');
  const rcvAtCol   = colIdx('Received_At');
  const revOkCol   = colIdx('revOk');
  const revNameCol = colIdx('Rev_Name');

  // totalNum = 辞退(revOk==='ng')・取消済(revOk==='cancelled')を除いた査読者数
  // endedNum = 承諾済み(revOk==='ok')かつ査読結果提出済み(Received_At が空でない)の人数
  // ※ダッシュボードの _allReviewsIn と同じ基準で判定する
  let totalNum = 0;
  let endedNum = 0;
  const submittedNames = [];
  const pendingNames   = [];
  const declinedNames  = [];
  const cancelledNames = [];

  for (let i = 1; i < rows.length; i++) {
    const rowMsVer   = String(rows[i][msVerCol]   || '').trim();
    const rowEdEmail = String(rows[i][edEmailCol] || '').trim();
    const rowRevOk   = String(rows[i][revOkCol]   || '').trim();
    const rowRcvAt   = String(rows[i][rcvAtCol]   || '').trim();
    const rowName    = String(rows[i][revNameCol] || '').trim();

    if (rowMsVer   !== String(msData.MsVer).trim())        continue;
    if (rowEdEmail !== String(msData.Editor_Email).trim()) continue;

    if (rowRevOk === 'ng')        { declinedNames.push(rowName);  continue; }
    if (rowRevOk === 'cancelled') { cancelledNames.push(rowName); continue; }

    totalNum++;
    if (rowRevOk === 'ok' && rowRcvAt !== '') {
      endedNum++;
      submittedNames.push(rowName);
    } else {
      pendingNames.push(rowName); // 未回答 または 受諾済み未提出
    }
  }

  const webAppUrl = ScriptApp.getService().getUrl();
  const recommendationLink = msData.editorKey
    ? webAppUrl + '?editorKey=' + msData.editorKey
    : webAppUrl;

  const paperTitle = [msData.TitleJP, msData.TitleEN].filter(Boolean).join(' / ');
  const allDone = (totalNum > 0 && endedNum === totalNum);

  const msInfoTable = `
    <table style="width:100%; font-size:14px; border-collapse:collapse; margin:20px 0;">
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee; width:30%;">Manuscript ID</th>
          <td style="padding:8px; border-bottom:1px solid #eee;">${msData.MsVer}</td></tr>
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">Type</th>
          <td style="padding:8px; border-bottom:1px solid #eee;">${msData.MS_Type || ''}</td></tr>
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">Title</th>
          <td style="padding:8px; border-bottom:1px solid #eee;">${paperTitle}</td></tr>
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">Corresponding Author</th>
          <td style="padding:8px; border-bottom:1px solid #eee;">${msData.CA_Name || ''}</td></tr>
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">Review Progress</th>
          <td style="padding:8px; border-bottom:1px solid #eee;"><strong>${endedNum} / ${totalNum}</strong> submitted</td></tr>
      ${submittedNames.length > 0 ? `
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">Submitted / 提出済</th>
          <td style="padding:8px; border-bottom:1px solid #eee; color:#059669;">${submittedNames.join(', ')}</td></tr>` : ''}
      ${pendingNames.length > 0 ? `
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">Pending / 未提出</th>
          <td style="padding:8px; border-bottom:1px solid #eee; color:#d97706;">${pendingNames.join(', ')}</td></tr>` : ''}
      ${declinedNames.length > 0 ? `
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">Declined / 辞退</th>
          <td style="padding:8px; border-bottom:1px solid #eee; color:#94a3b8;">${declinedNames.join(', ')}</td></tr>` : ''}
      ${cancelledNames.length > 0 ? `
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">Cancelled / 取消</th>
          <td style="padding:8px; border-bottom:1px solid #eee; color:#94a3b8;">${cancelledNames.join(', ')}</td></tr>` : ''}
    </table>
  `;

  let subject, bodyHtml, buttonUrl, buttonLabel;

  if (allDone) {
    subject = `[${settings.Journal_Name}] 全査読完了 / All Reviews Complete — ${msData.MsVer}`;
    bodyHtml = `
      <p>All peer review results are now available for manuscript <strong>${msData.MsVer}</strong>. Please refer to those results and recommend your decision to the Editor-in-Chief by clicking the button below.</p>
      <p>原稿 <strong>${msData.MsVer}</strong> の全ての査読結果が出そろいました。これらの結果を確認し、以下のボタンより編集委員長への推薦（判定案の作成）をお願いいたします。</p>
      ${msInfoTable}
    `;
    buttonUrl   = recommendationLink;
    buttonLabel = 'Send Recommendation / 推薦を送る';
  } else {
    subject = `[${settings.Journal_Name}] 査読結果提出 (${endedNum}/${totalNum}) / Review Submitted — ${msData.MsVer}`;
    bodyHtml = `
      <p>A peer review result has been submitted for manuscript <strong>${msData.MsVer}</strong> by <strong>${msData.Rev_Name || 'a reviewer'}</strong>. ${pendingNames.length} reviewer(s) have not yet submitted their results. You will receive another notification when all reviews are complete.</p>
      <p>原稿 <strong>${msData.MsVer}</strong> の査読結果が1件届きました（${endedNum}/${totalNum} 件完了）。残り ${pendingNames.length} 名の査読者からの結果待ちです。全員の査読が完了した際に改めてご連絡いたします。</p>
      ${msInfoTable}
    `;
    buttonUrl   = recommendationLink;
    buttonLabel = 'Open Editor Menu / 担当編集メニューを開く';
  }

  const html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting: `Dear Dr. ${msData.Editor_Name},`,
    bodyHtml: bodyHtml,
    buttonUrl: buttonUrl,
    buttonLabel: buttonLabel,
    footerHtml: settings.mailFooter || ''
  });

  sendEmailSafe({ to: msData.Editor_Email, subject, htmlBody: html },
    'Review Result Notification: ' + msData.MsVer + ' to editor ' + msData.Editor_Name);
}

function sendReviewerThankYou(email, name, msData, settings) {
  if (!email) return;
  const msVer = msData.MsVer || '';
  const subject = `[${settings.Journal_Name}] 査読結果送信の確認 / Confirmation: Review Submitted: ${msVer}`;
  
  const paperTitle = (msData.TitleJP && msData.TitleEN) ? msData.TitleJP + ' / ' + msData.TitleEN : (msData.TitleJP || msData.TitleEN || '');
  
  const bodyHtml = `
    <p>Thank you for sending us the peer review results for the following manuscript. We have successfully received your submission.</p>
    <p>以下の原稿の査読結果をご送付いただき、誠にありがとうございます。査読結果、およびコメントを確かに受領いたしましたことをご報告申し上げます。</p>
    <table style="width:100%; font-size: 14px; border-collapse: collapse; margin: 20px 0;">
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee; width: 30%;">ID</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${msVer}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Type</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${msData.MS_Type || ''}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee; width: 30%;">Title</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${paperTitle}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Abstract (JP)</th><td style="padding: 8px; border-bottom: 1px solid #eee; white-space: pre-wrap;">${msData.AbstractJP || 'N/A'}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Abstract (EN)</th><td style="padding: 8px; border-bottom: 1px solid #eee; white-space: pre-wrap;">${msData.AbstractEN || 'N/A'}</td></tr>
    </table>
  `;

  const html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting: `Dear Dr. ${name},`,
    bodyHtml: bodyHtml,
    footerHtml: settings.mailFooter || ''
  });
  
  sendEmailSafe({ to: email, subject, htmlBody: html },
    'Reviewer Thank You: ' + msVer + ' to ' + name);
}
