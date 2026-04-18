/**
 * EmailModule.gs - メール送信および受領票（PDF）生成ロジック（修正版）
 */

/**
 * 著者への受領確認メール送信（受領票PDFを添付）
 *
 * Settings シートの追加設定:
 *   submissionBccEmails  投稿通知をBCCで受け取る担当者のメールアドレス（カンマ区切り、空欄可）
 *                        例: secretary@example.jp, checker@example.jp
 */
function sendReceiptEmail(ssId, ms) {
  const settings = getSettings();

  // 1. 受領票の作成
  const pdfBlob = generateReceiptPdf(ms, settings);

  // 2. メールの内容構築
  const subject = `[${settings.Journal_Name}] 原稿受領のお知らせ / Manuscript Received: ${ms.MsVer}`;
  const webAppUrl = ScriptApp.getService().getUrl();
  const showProgressUrl = webAppUrl + '?key=' + ms.key;
  const paperTitle = (ms.titleJp && ms.titleEn)
    ? ms.titleJp + ' / ' + ms.titleEn
    : (ms.titleJp || ms.titleEn || '');

  // テーブル形式で原稿情報を表示（他のロール向けメールと統一した書式）
  const bodyHtml = `
    <p>We are pleased to confirm the receipt of your manuscript submitted to <strong>${escHtml(settings.Journal_Name)}</strong>.</p>
    <p>Please find the submission receipt attached for your reference. You can also check the current status of your submission by clicking the button below.</p>
    <table style="width:100%; font-size:14px; border-collapse:collapse; margin:20px 0;">
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee; width:30%;">Manuscript ID / 受付番号</th>
          <td style="padding:8px; border-bottom:1px solid #eee;">${escHtml(ms.MsVer || '')}</td></tr>
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">Type / 原稿種別</th>
          <td style="padding:8px; border-bottom:1px solid #eee;">${escHtml(ms.paperType || '')}</td></tr>
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">Title (JP) / タイトル（日）</th>
          <td style="padding:8px; border-bottom:1px solid #eee;">${escHtml(ms.titleJp || '')}</td></tr>
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">Title (EN) / タイトル（英）</th>
          <td style="padding:8px; border-bottom:1px solid #eee;">${escHtml(ms.titleEn || '')}</td></tr>
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">Corresponding Author / 責任著者</th>
          <td style="padding:8px; border-bottom:1px solid #eee;">${escHtml(ms.authorName || '')} (${escHtml(ms.authorEmail || '')})</td></tr>
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">Submitted / 投稿日時</th>
          <td style="padding:8px; border-bottom:1px solid #eee;">${escHtml(ms.sendDateTime || '')}</td></tr>
    </table>
    <hr style="border:none; border-top:1px solid #e2e8f0; margin:20px 0;">
    <p><strong>${escHtml(settings.Journal_Name)}</strong>への投稿を受領いたしました。</p>
    <p>受領票を添付いたしましたので、大切に保管してください。</p>
    <p>現在の投稿ステータスは、以下のボタンよりいつでもご確認いただけます。</p>
  `;

  const html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting: `Dear Dr. ${escHtml(ms.authorName)},`,
    bodyHtml: bodyHtml,
    buttonUrl: showProgressUrl,
    buttonLabel: 'Check Status / ステータスを確認する',
    footerHtml: settings.mailFooter || ''
  });

  const mailOptions = {
    to: ms.authorEmail,
    subject: subject,
    htmlBody: html
  };

  // CC設定: 共著者（example.com ドメインは除外）
  const filteredCc = _filterExampleDomains(ms.ccEmails || '');
  if (filteredCc) mailOptions.cc = filteredCc;

  // BCC設定: 事務・初期チェック担当者（Settings: submissionBccEmails）
  const bccRaw = String(settings.submissionBccEmails || '').trim();
  if (bccRaw) mailOptions.bcc = bccRaw;

  // PDF添付（生成可能な場合）
  if (pdfBlob) mailOptions.attachments = [pdfBlob];

  sendEmailSafe(mailOptions, 'Receipt: ' + ms.MsVer + ' to ' + ms.authorName);

  return pdfBlob;
}

/**
 * メールアドレスリスト（カンマ区切り）から @example.com ドメインを除外して返す
 * テストデータとして使われるドメインを誤送信しないための安全ガード
 * @param {string} emailList  カンマ区切りのメールアドレス文字列
 * @returns {string} フィルタ後のカンマ区切り文字列（該当なしは空文字）
 */
function _filterExampleDomains(emailList) {
  if (!emailList || !emailList.trim()) return '';
  const filtered = emailList
    .split(',')
    .map(e => e.trim())
    .filter(e => e && !e.toLowerCase().endsWith('@example.com'));
  return filtered.join(', ');
}

/**
 * 編集委員長（EIC）への担当編集者選定依頼メール（修正版）
 * EICには著者キー（key）を送る。EIC自身はkeyがManuscripts.keyで検索され
 * author ロールで表示されるが、将来的にはeditorKeyを別途生成・送付する想定。
 */
function sendEicNotification(ms, pdfBlob) {
  const settings = getSettings();
  const ssId = getSpreadsheetId();
  const webAppUrl = ScriptApp.getService().getUrl();
  // EIC用には編集委員長専用のeicKeyを使う（著者キーとは独立したキー）
  const msDetailLink = webAppUrl + '?eicKey=' + ms.eicKey;
  const paperTitle = (ms.titleJp && ms.titleEn) ? ms.titleJp + ' / ' + ms.titleEn : (ms.titleJp || ms.titleEn || '');

  // 再投稿の場合: 過去バージョンの履歴テーブルを組み立てる
  const currentVerNo = Number(ms.Ver_No || ms.verNo || 1);
  let versionHistoryHtml = '';
  if (currentVerNo > 1) {
    const msIdStr = String(ms.MS_ID || ms.msId || '').trim();
    if (msIdStr) {
      const allVersions = findAllRecordsByKey(ssId, MANUSCRIPTS_SHEET_NAME, 'MS_ID', msIdStr);
      const priorVersions = allVersions
        .filter(v => Number(v.Ver_No || 0) < currentVerNo)
        .sort((a, b) => Number(a.Ver_No || 0) - Number(b.Ver_No || 0));
      if (priorVersions.length > 0) {
        const tz = SpreadsheetApp.openById(ssId).getSpreadsheetTimeZone();
        const fmtD = val => {
          if (!val) return '-';
          if (val instanceof Date) return Utilities.formatDate(val, tz, 'yyyy/MM/dd HH:mm');
          return String(val).trim() || '-';
        };
        const rows = priorVersions.map(v => {
          const vMsVer = String(v.MsVer || '').trim();
          const vEditorLogs = findAllRecordsByKey(ssId, EDITOR_LOG_SHEET_NAME, 'MsVer', vMsVer);
          const vAcceptedEditor = vEditorLogs.find(log => String(log.edtOk || '').trim() === 'ok');
          const reportPdfUrl = vAcceptedEditor ? String(vAcceptedEditor.reportPdfUrl || '').trim() : '';
          const reportLink = reportPdfUrl
            ? `<a href="${reportPdfUrl}" target="_blank" style="color:#2563eb;">📊 Report</a>`
            : '-';
          return `<tr>
            <td style="padding:6px 8px; border-bottom:1px solid #eee;">${vMsVer}</td>
            <td style="padding:6px 8px; border-bottom:1px solid #eee;">${fmtD(v.Submitted_At)}</td>
            <td style="padding:6px 8px; border-bottom:1px solid #eee; font-weight:bold;">${String(v.score || '').trim() || '-'}</td>
            <td style="padding:6px 8px; border-bottom:1px solid #eee;">${fmtD(v.sentBackAt || '')}</td>
            <td style="padding:6px 8px; border-bottom:1px solid #eee;">${reportLink}</td>
          </tr>`;
        }).join('');
        versionHistoryHtml = `
    <hr style="border:none; border-top:1px solid #e2e8f0; margin: 20px 0;">
    <h4 style="margin:0 0 0.5rem; color:#1e40af; font-size:14px;">投稿履歴 / Previous Submission History</h4>
    <table style="width:100%; font-size:13px; border-collapse:collapse; margin-bottom:1rem;">
      <tr style="background:#f1f5f9;">
        <th style="padding:6px 8px; text-align:left; border-bottom:2px solid #cbd5e1;">Version</th>
        <th style="padding:6px 8px; text-align:left; border-bottom:2px solid #cbd5e1;">Submitted / 投稿日</th>
        <th style="padding:6px 8px; text-align:left; border-bottom:2px solid #cbd5e1;">Decision / 判定</th>
        <th style="padding:6px 8px; text-align:left; border-bottom:2px solid #cbd5e1;">Sent Back / 返送日</th>
        <th style="padding:6px 8px; text-align:left; border-bottom:2px solid #cbd5e1;">Report</th>
      </tr>
      ${rows}
    </table>`;
      }
    }
  }

  const isResubmission = currentVerNo > 1;
  const subject = isResubmission
    ? `[${settings.Journal_Name}] 修正版原稿受領のお知らせ / Revised Manuscript Received: ${ms.MsVer}`
    : `[${settings.Journal_Name}] 新規原稿受領のお知らせ / New Manuscript Received: ${ms.MsVer}`;
  const bodyHtml = `
    <p>${isResubmission ? 'A revised manuscript (resubmission) has been received.' : 'A new manuscript has been submitted.'} Please review the submission details and assign a responsible editor.</p>
    <p>${isResubmission ? '修正原稿（再投稿）を受領いたしました。' : '新規投稿がありました。'}内容を確認し、担当編集者の割り当てを行ってください。</p>
    <table style="width:100%; font-size: 14px; border-collapse: collapse; margin: 20px 0;">
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee; width: 30%;">ID</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${escHtml(ms.MsVer)}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Type</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${escHtml(ms.paperType)}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Title</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${escHtml(paperTitle)}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Author</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${escHtml(ms.authorName)} (${escHtml(ms.authorEmail)})</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Abstract (JP)</th><td style="padding: 8px; border-bottom: 1px solid #eee; white-space: pre-wrap;">${escHtml(ms.abstractJp || 'N/A')}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Abstract (EN)</th><td style="padding: 8px; border-bottom: 1px solid #eee; white-space: pre-wrap;">${escHtml(ms.abstractEn || 'N/A')}</td></tr>
    </table>
    ${versionHistoryHtml}
  `;

  const html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting: `Dear Editor-in-Chief / 編集委員長殿,`,
    bodyHtml: bodyHtml,
    buttonUrl: msDetailLink,
    buttonLabel: 'Assign Editor / 担当編集者を指名する',
    footerHtml: settings.mailFooter || ''
  });

  if (!settings.chiefEditorEmail) {
    Logger.log('sendEicNotification: chiefEditorEmail が設定されていないためスキップします (MsVer: ' + ms.MsVer + ')');
    return;
  }

  const mailOptions = {
    to: settings.chiefEditorEmail,
    subject: subject,
    htmlBody: html
  };

  if (pdfBlob) {
    mailOptions.attachments = [pdfBlob];
  }

  sendEmailSafe(mailOptions, 'EIC Notification: ' + ms.MsVer);
}

/**
 * 投稿データから受領票HTMLを生成してPDF Blobを返す
 * スプレッドシートテンプレートを使わず、投稿情報を直接レイアウトしてPDF化する
 */
function generateReceiptPdf(ms, settings) {
  try {
    const journalName = (settings && settings.Journal_Name) ? settings.Journal_Name : 'Journal';
    const date = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
    const fileName = `${ms.MsVer || 'Receipt'}_Receipt.pdf`;

    const esc = s => String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');

    // グリッド形式の行（2カラム対応）
    const gridItem = (labelJp, labelEn, value, isFullWidth = false) => `
      <div class="grid-item ${isFullWidth ? 'full' : ''}">
        <div class="label">
          ${esc(labelJp)} <span class="label-en">/ ${esc(labelEn)}</span>
        </div>
        <div class="value">${esc(value)}</div>
      </div>`;

    const authorEmail = ms.authorEmail + (ms.ccEmails ? '\n（CC）: ' + ms.ccEmails : '');

    const html = `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<style>
  @page { margin: 20pt; }
  * { box-sizing: border-box; }
  body {
    font-family: 'Hiragino Kaku Gothic Pro', 'Meiryo', 'Arial', sans-serif;
    font-size: 10.5pt;
    color: hsl(220, 15%, 15%);
    margin: 0;
    padding: 0;
    line-height: 1.4;
    background: #fff;
  }
  .header {
    display: flex;
    justify-content: space-between;
    align-items: flex-end;
    border-bottom: 2pt solid hsl(225, 71%, 45%);
    padding-bottom: 8pt;
    margin-bottom: 12pt;
  }
  .journal-name {
    font-size: 14pt;
    font-weight: bold;
    color: hsl(225, 71%, 45%);
    margin: 0;
  }
  .receipt-title-box {
    text-align: center;
    margin: 16pt 0 12pt;
  }
  .receipt-title {
    font-size: 18pt;
    font-weight: bold;
    margin: 0;
  }
  .receipt-subtitle {
    font-size: 9pt;
    color: hsl(220, 10%, 45%);
    margin: 4pt 0 0;
  }
  .meta-grid {
    display: table;
    width: 100%;
    margin-bottom: 16pt;
    font-size: 9pt;
    color: hsl(220, 10%, 40%);
  }
  .meta-item {
    display: table-cell;
    width: 50%;
  }
  .meta-item.right { text-align: right; }

  .card {
    background: hsl(210, 40%, 98%);
    border-radius: 6pt;
    padding: 10pt;
    margin-bottom: 12pt;
    border: 0.5pt solid hsl(210, 20%, 90%);
  }
  .section-heading {
    font-size: 10pt;
    font-weight: bold;
    color: hsl(225, 71%, 45%);
    margin: 0 0 8pt;
    display: flex;
    align-items: center;
  }
  .section-heading::before {
    content: "";
    display: inline-block;
    width: 3pt;
    height: 10pt;
    background: hsl(225, 71%, 45%);
    margin-right: 6pt;
    border-radius: 1pt;
  }

  .grid-container {
    display: block; /* PDF generation compatibility */
    width: 100%;
  }
  .grid-row {
    display: table;
    width: 100%;
    table-layout: fixed;
    margin-bottom: 6pt;
  }
  .grid-item {
    display: table-cell;
    vertical-align: top;
    padding-right: 10pt;
  }
  .grid-item.full {
    display: block;
    width: 100%;
    padding-right: 0;
  }
  .label {
    font-size: 8.5pt;
    font-weight: bold;
    color: hsl(220, 10%, 40%);
    margin-bottom: 2pt;
  }
  .label-en {
    font-weight: normal;
    font-size: 8pt;
    color: hsl(220, 10%, 55%);
  }
  .value {
    font-size: 10pt;
    white-space: pre-wrap;
    word-break: break-all;
  }
  .footer {
    margin-top: 20pt;
    border-top: 0.5pt solid hsl(220, 15%, 90%);
    padding-top: 8pt;
    font-size: 9pt;
    color: hsl(220, 10%, 60%);
    text-align: center;
  }
</style>
</head>
<body>
  <div class="header">
    <div class="journal-name">${esc(journalName)}</div>
  </div>

  <div class="receipt-title-box">
    <h1 class="receipt-title">原稿受領票 / Manuscript Receipt</h1>
    <p class="receipt-subtitle">本票は原稿の受領を確認するものです。 This confirms your manuscript receipt.</p>
  </div>

  <div class="meta-grid">
    <div class="meta-item">受付番号 / Receipt No: <strong>${esc(ms.MsVer || '')}</strong></div>
    <div class="meta-item right">発行日時 / Issued: <strong>${date}</strong></div>
  </div>

  <div class="card">
    <div class="section-heading">著者情報 / Author Information</div>
    <div class="grid-row">
      ${gridItem('責任著者', 'Corresponding Author', ms.authorName || '')}
      ${gridItem('メールアドレス', 'Email', authorEmail)}
    </div>
    ${gridItem('所属機関', 'Affiliation', ms.authorAffiliation || '', true)}
    <div class="grid-row">
      ${gridItem('著者全員（日）', 'All Authors (JP)', ms.authorsJp || '', true)}
    </div>
    <div class="grid-row">
      ${gridItem('著者全員（英）', 'All Authors (EN)', ms.authorsEn || '', true)}
    </div>
  </div>

  <div class="card">
    <div class="section-heading">原稿情報 / Manuscript Information</div>
    <div class="grid-row">
      ${gridItem('原稿種別', 'Type', ms.paperType || '')}
      ${gridItem('投稿日時', 'Submitted at', ms.sendDateTime || date)}
    </div>
    ${gridItem('タイトル（日）', 'Title (JP)', ms.titleJp || '', true)}
    ${gridItem('タイトル（英）', 'Title (EN)', ms.titleEn || '', true)}
    ${gridItem('ランニングタイトル', 'Running Title', ms.runningTitle || '', true)}
    ${gridItem('投稿ファイル', 'Files', ms.submittedFiles || '', true)}
  </div>

  ${(ms.reprintRequest || ms.englishEditing) ? `
  <div class="card">
    <div class="section-heading">その他 / Other</div>
    <div class="grid-row">
      ${ms.reprintRequest ? gridItem('別刷希望', 'Reprint', ms.reprintRequest) : ''}
      ${ms.englishEditing ? gridItem('英文校閲', 'English Editing', ms.englishEditing) : ''}
    </div>
  </div>
  ` : ''}

  <div class="footer">
    ${settings.mailFooter || ''}<br>
    Copyright &copy; ${new Date().getFullYear()} ${esc(journalName)}. All rights reserved.
  </div>
</body>
</html>`;

    const blob = HtmlService.createHtmlOutput(html)
      .getBlob()
      .getAs(MimeType.PDF)
      .setName(fileName);

    return blob;
  } catch (err) {
    Logger.log('generateReceiptPdf error: ' + err);
    return null;
  }
}

/**
 * ログ記録
 */
/**
 * クォータを確認してメール送信。上限超過時は Emails シートに一時保存する。
 * @param {Object} options MailApp.sendEmail と同じ形式 (to, cc, subject, htmlBody, attachments)
 * @param {string} logText ログ用の説明文
 * @returns {boolean} 即時送信に成功した場合のみ true。キューイング・失敗時は false。
 */
function sendEmailSafe(options, logText) {
  const quota = MailApp.getRemainingDailyQuota();
  if (quota > 0) {
    try {
      MailApp.sendEmail(options);
      Logger.log('Email sent: ' + logText);
      return true;
    } catch (e) {
      Logger.log('Email send failed (' + logText + '): ' + e.message + ' — saving to pending');
      savePendingEmail(options, logText);
      return false;
    }
  }
  Logger.log('Quota exhausted — saving to pending: ' + logText);
  savePendingEmail(options, logText);
  return false;
}

/**
 * 送信できなかったメールを Emails シートに保存する。
 * 添付ファイル (Blob) は Drive の一時フォルダに保存し、ファイルIDを記録する。
 */
function savePendingEmail(options, logText) {
  const ssId = getSpreadsheetId();
  const ss = SpreadsheetApp.openById(ssId);

  // Emails シートが存在しない場合は自動作成
  let sheet = ss.getSheetByName(PENDING_EMAILS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PENDING_EMAILS_SHEET_NAME);
    sheet.appendRow(['to', 'cc', 'bcc', 'subject', 'htmlBody', 'attachmentFileIds', 'savedAt', 'logText']);
  }

  // 添付ファイル (Blob) を Drive の一時フォルダに保存
  let attachmentIds = '';
  if (options.attachments && options.attachments.length > 0) {
    try {
      const settings = getSettings();
      const rootName = settings.SUBFOLDER || 'Journal Files';
      const root = driveFolderCache.getRootFolder(rootName) || DriveApp.createFolder(rootName);
      const tempFolder = driveFolderCache.getOrCreateFolder(root, '_pending_attachments');

      const ids = options.attachments.map(att => {
        try {
          // Drive File オブジェクトの場合はそのまま ID を使用
          if (typeof att.getId === 'function') return att.getId();
          // Blob の場合は Drive に保存して ID を取得
          return tempFolder.createFile(att).getId();
        } catch (e) {
          Logger.log('Failed to save attachment: ' + e.message);
          return null;
        }
      }).filter(Boolean);
      attachmentIds = ids.join(',');
    } catch (e) {
      Logger.log('Failed to save attachments to Drive: ' + e.message);
    }
  }

  sheet.appendRow([
    options.to    || '',
    options.cc    || '',
    options.bcc   || '',
    options.subject  || '',
    options.htmlBody || '',
    attachmentIds,
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'),
    logText || ''
  ]);
}

/**
 * 送信保留メールの再送処理。
 * 時間ベーストリガー（例: 1時間ごと）に設定して使用する。
 * GAS エディタ → トリガー → retrySendingEmails → 時間ベース で設定してください。
 */
function retrySendingEmails() {
  const ssId = getSpreadsheetId();
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(PENDING_EMAILS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return;

  const data = sheet.getDataRange().getValues();
  let quota = MailApp.getRemainingDailyQuota();
  const sentRows = []; // 送信成功した行番号（1始まり）を収集

  // 古い順（上から）に処理
  for (let i = 1; i < data.length; i++) {
    if (quota <= 0) break;

    const [to, cc, bcc, subject, htmlBody, attachmentIds, , logText] = data[i];
    if (!to) continue;

    // 添付ファイルを Drive から取得
    const attachments = [];
    if (attachmentIds) {
      attachmentIds.split(',').forEach(fileId => {
        const id = fileId.trim();
        if (!id) return;
        try {
          attachments.push(DriveApp.getFileById(id).getBlob());
        } catch (e) {
          Logger.log('Cannot retrieve attachment ' + id + ': ' + e.message);
        }
      });
    }

    const mailOptions = { to, subject, htmlBody };
    if (cc)  mailOptions.cc  = cc;
    if (bcc) mailOptions.bcc = bcc;
    if (attachments.length > 0) mailOptions.attachments = attachments;

    try {
      MailApp.sendEmail(mailOptions);
      quota--;
      sentRows.push(i + 1); // シート行番号（ヘッダー行 +1）
      Logger.log('Pending email resent to ' + to + ' | ' + logText);

      // 一時添付ファイルを Drive から削除
      if (attachmentIds) {
        attachmentIds.split(',').forEach(fileId => {
          const id = fileId.trim();
          if (!id) return;
          try { DriveApp.getFileById(id).setTrashed(true); } catch(e) {}
        });
      }
    } catch (e) {
      Logger.log('Failed to resend to ' + to + ': ' + e.message);
    }
  }

  // 送信成功した行を後ろから削除（行番号ズレ防止）
  sentRows.reverse().forEach(rowNum => sheet.deleteRow(rowNum));
}


/**
 * リッチなHTMLメール制作のためのテンプレートレンダラー
 * @param {Object} params { title, greeting, bodyHtml, buttonUrl, buttonLabel, footerHtml, journalName }
 */
function renderRichEmail(params) {
  const primaryColor = '#2563eb'; // Sleek blue
  const bgColor = '#f8fafc';
  const cardBg = '#ffffff';
  const textColor = '#1e293b';
  
  const html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    .btn:hover { background-color: #1d4ed8 !important; }
  </style>
</head>
<body style="margin:0; padding:0; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; background-color: ${bgColor}; color: ${textColor}; line-height: 1.6;">
  <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color: ${bgColor}; padding: 40px 10px;">
    <tr>
      <td align="center">
        <table width="100%" border="0" cellspacing="0" cellpadding="0" style="max-width: 600px; background-color: ${cardBg}; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);">
          <!-- Header -->
          <tr>
            <td style="background-color: ${primaryColor}; padding: 30px; text-align: center;">
              <h1 style="color: #ffffff; margin: 0; font-size: 20px; font-weight: 600; letter-spacing: 0.5px;">
                ${escHtml(params.journalName || 'Journal')}
              </h1>
            </td>
          </tr>

          <!-- Content -->
          <tr>
            <td style="padding: 40px 30px;">
              <p style="margin-top: 0; font-size: 16px; font-weight: 600;">${escHtml(params.greeting)}</p>
              <div style="font-size: 15px; color: ${textColor};">
                ${params.bodyHtml}
              </div>

              ${params.buttonUrl ? `
              <div style="margin-top: 40px; text-align: center;">
                <a href="${params.buttonUrl}" class="btn" style="display: inline-block; background-color: ${primaryColor}; color: #ffffff; padding: 14px 28px; border-radius: 8px; text-decoration: none; font-weight: 600; font-size: 15px;">
                  ${escHtml(params.buttonLabel || 'View Details / 内容を確認する')}
                </a>
              </div>
              ` : ''}
            </td>
          </tr>
          
          <!-- Footer -->
          <tr>
            <td style="padding: 30px; background-color: #f1f5f9; border-top: 1px solid #e2e8f0; font-size: 13px; color: #64748b; text-align: center;">
              ${params.footerHtml || ''}
              <p style="margin-top: 20px; font-size: 11px; color: #94a3b8;">
                This is an automated message from the ${params.journalName || 'Journal'}.<br>
                本メールはシステムによる自動送信です。
              </p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>
  `;
  return html;
}
