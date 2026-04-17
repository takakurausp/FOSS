/**
 * ReviewerModule.js - 査読者選定と査読依頼の統合バックエンド (SPA)
 */

function apiAssignReviewers(data) {
  const ssId = getSpreadsheetId();
  const settings = getSettings();
  
  // 1 & 2. 原稿情報全体を統合して取得 (フォルダURL等も自動取得される)
  const msData = getManuscriptData('editor', data.editorKey);
  if (!msData) throw new Error("Editor record not found.");
  
  const msVer = msData.MsVer || msData.msver || data.msVer;
  let currentHex = msData.MsVerRevHex; 
  if (!currentHex) throw new Error("MsVerRevHex not found for the editor.");
  
  // 3. メールの基本情報の準備
  const editorName = msData.Editor_Name || '';
  const editorEmail = msData.Editor_Email || '';
  const reviewDeadline = computeReviewDeadline(settings.Review_Period || '21'); 
  
  const results = [];
  
  // 4 & 5. 採番 → DB書き込みをスクリプトロック内でアトミックに実行する。
  // incrementHex の「空き確認」と logReviewerToDb の「行追加」の間に
  // 別リクエストが割り込んで同じ hexID を取得する競合を防ぐ。
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // 最大30秒待機（タイムアウト時は例外）
  let rev1Hex, rev1Key, rev2Hex, rev2Key;
  try {
    rev1Hex = incrementHex(ssId, currentHex);
    rev1Key = rev1Hex + Utilities.getUuid().replace(/-/g, '');  // CSPRNG (128 bit)
    logReviewerToDb(ssId, rev1Hex, msVer, rev1Key, editorName, editorEmail, data.nameCandidate1, data.emailCandidate1, reviewDeadline);
    results.push(rev1Key);

    if (data.numberReviewer === '2' && data.emailCandidate2) {
      rev2Hex = incrementHex(ssId, rev1Hex);
      rev2Key = rev2Hex + Utilities.getUuid().replace(/-/g, '');  // CSPRNG (128 bit)
      logReviewerToDb(ssId, rev2Hex, msVer, rev2Key, editorName, editorEmail, data.nameCandidate2, data.emailCandidate2, reviewDeadline);
      results.push(rev2Key);
    }
  } finally {
    // キャッシュを無効化してから解放することで、次のリクエストが新行を確実に読む
    spreadsheetCache.invalidate(ssId, REVIEW_LOG_SHEET_NAME);
    lock.releaseLock();
  }

  // メール送信はロック解放後に実行（時間のかかる処理をロック内に含めない）
  sendReviewerMail(data.emailCandidate1, data.nameCandidate1, data.letterToCandidate1, msData, settings, rev1Key, editorName, editorEmail, reviewDeadline);
  if (rev2Key) {
    sendReviewerMail(data.emailCandidate2, data.nameCandidate2, data.letterToCandidate2, msData, settings, rev2Key, editorName, editorEmail, reviewDeadline);
  }
  
  writeLog(`Reviewers Assigned: ${msVer} by ${editorName} (Target: ${data.nameCandidate1}${data.nameCandidate2 ? ', ' + data.nameCandidate2 : ''})`);
  
  return { success: true, targets: results };
}

function computeReviewDeadline(daysString) {
  const days = parseInt(daysString, 10) || 21;
  const d = new Date();
  d.setDate(d.getDate() + days);
  return Utilities.formatDate(d, 'JST', 'yyyy/MM/dd');
}

function sendReviewerMail(email, name, letterText, msData, settings, revKey, editorName, editorEmail, deadline) {
  const webAppUrl = ScriptApp.getService().getUrl();
  // Webアプリの単一エントリーポイント（doGet）にアクションを持たせる
  const reviewUrl = webAppUrl + '?reviewKey=' + revKey;
  
  const subject = `[${settings.Journal_Name}] 査読のご依頼 / Invitation for Reviewing Manuscript (${msData.MsVer})`;
  const paperTitle = escHtml((msData.TitleJP && msData.TitleEN) ? msData.TitleJP + ' / ' + msData.TitleEN : (msData.TitleJP || msData.TitleEN || ''));

  const bodyHtml = `
    <p>You have been invited to review the following manuscript for <strong>${escHtml(settings.Journal_Name)}</strong>:</p>
    <div style="background:#f1f5f9; padding:15px; border-radius:8px; margin:20px 0;">
      <p><strong>Message from Responsible Editor:</strong><br>${escHtml(letterText || '').replace(/\n/g, '<br>')}</p>
    </div>
    <table style="width:100%; font-size: 14px; border-collapse: collapse; margin: 20px 0;">
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee; width: 30%;">ID</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${escHtml(msData.MsVer)}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Type</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${escHtml(msData.MS_Type || '')}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee; width: 30%;">Title</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${paperTitle}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Abstract (JP)</th><td style="padding: 8px; border-bottom: 1px solid #eee; white-space: pre-wrap;">${escHtml(msData.AbstractJP || 'N/A')}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Abstract (EN)</th><td style="padding: 8px; border-bottom: 1px solid #eee; white-space: pre-wrap;">${escHtml(msData.AbstractEN || 'N/A')}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Deadline</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${escHtml(deadline)}</td></tr>
    </table>
    <p>Please respond to this invitation by clicking the button below.</p>
    <hr style="border:none; border-top:1px solid #e2e8f0; margin: 20px 0;">
    <p>以下の原稿について、査読のお願いを申し上げます。内容をご確認の上、以下のボタンより受諾または辞退のご回答をお願いいたします。</p>
  `;

  const html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting: `Dear Dr. ${name},`,
    bodyHtml: bodyHtml,
    buttonUrl: reviewUrl,
    buttonLabel: 'Respond to Invitation / 依頼に回答する',
    footerHtml: settings.mailFooter || ''
  });

  sendEmailSafe({ to: email, subject, htmlBody: html },
    'Reviewer Invitation: ' + msData.MsVer + ' to ' + name);
}

function sendReviewerConfirmationToEditor(editorEmail, editorName, data, settings, msVer, msData) {
  const subject = `[${settings.Journal_Name}] 査読依頼送信の確認 / Confirmation: Reviewer Invitation Sent (${msVer})`;

  const titleJP = msData.TitleJP || '';
  const titleEN = msData.TitleEN || '';
  const titleCell = titleJP && titleEN
    ? `${titleJP}<br>${titleEN}`
    : titleJP || titleEN || '';

  const reviewerRows = `
    <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee; width:30%;">Reviewer #1</th>
        <td style="padding:8px; border-bottom:1px solid #eee;">${data.nameCandidate1} (${data.emailCandidate1})</td></tr>
    ${data.numberReviewer === '2' && data.nameCandidate2 ? `
    <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">Reviewer #2</th>
        <td style="padding:8px; border-bottom:1px solid #eee;">${data.nameCandidate2} (${data.emailCandidate2})</td></tr>` : ''}
  `;

  const dashboardUrl = msData.editorKey
    ? ScriptApp.getService().getUrl() + '?editorKey=' + msData.editorKey
    : null;

  const bodyHtml = `
    <p>Thank you for sending us the peer reviewer candidate(s) for the following manuscript.</p>
    <table style="width:100%; font-size:14px; border-collapse:collapse; margin:20px 0;">
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee; width:30%;">Manuscript / 原稿</th>
          <td style="padding:8px; border-bottom:1px solid #eee;">${msData.MsVer || msData.MS_ID || ''}</td></tr>
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">Type / 種別</th>
          <td style="padding:8px; border-bottom:1px solid #eee;">${msData.MS_Type || ''}</td></tr>
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">Title / タイトル</th>
          <td style="padding:8px; border-bottom:1px solid #eee;">${titleCell}</td></tr>
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">Corresponding Author / 責任著者</th>
          <td style="padding:8px; border-bottom:1px solid #eee;">${msData.CA_Name || ''}</td></tr>
      ${reviewerRows}
    </table>
    <hr style="border:none; border-top:1px solid #e2e8f0; margin:20px 0;">
    <p>査読者候補のご登録を承りました。各候補者へ査読依頼メールを送信しました。</p>
    <p>以下のボタンから担当編集者ダッシュボードを開き、招待状況をご確認いただけます。</p>
    <p>Please use the button below to open your editor dashboard and check the invitation status.</p>
  `;

  const html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting: `Dear Dr. ${editorName},`,
    bodyHtml: bodyHtml,
    buttonUrl:   dashboardUrl || undefined,
    buttonLabel: dashboardUrl ? 'Open Editor Dashboard / 担当編集者ダッシュボードを開く' : undefined,
    footerHtml: settings.mailFooter || ''
  });

  sendEmailSafe({ to: editorEmail, subject, htmlBody: html },
    'Reviewer Assignment Confirmation: ' + msVer + ' to ' + editorName);
}

function logReviewerToDb(ssId, hexId, msVer, revKey, edName, edEmail, revName, revEmail, reviewDeadline) {
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(REVIEW_LOG_SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newRow = new Array(headers.length).fill('');

  const todayNow = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm');

  const mapping = {
    'MsVerRevHex': hexId,
    'MsVer': msVer,
    'reviewKey': revKey,
    'Editor_Email': edEmail,
    'Editor_Name': edName,
    'Rev_Name': revName,
    'Rev_Email': revEmail,
    'Ask_At': todayNow,
    'Review_Deadline': reviewDeadline || ''
  };

  const writtenKeys = new Set();
  headers.forEach((h, i) => {
    const headerLower = String(h).toLowerCase();
    for (const k of Object.keys(mapping)) {
      if (k.toLowerCase() === headerLower) {
        newRow[i] = mapping[k];
        writtenKeys.add(k.toLowerCase());
        break;
      }
    }
  });

  sheet.appendRow(newRow);

  // 対応する列が存在しなかったキーは、列ヘッダーを新規作成して値を書き込む
  const newRowNum = sheet.getLastRow();
  let nextCol = headers.length + 1;
  for (const k of Object.keys(mapping)) {
    if (!writtenKeys.has(k.toLowerCase()) && mapping[k] !== '') {
      sheet.getRange(1, nextCol).setValue(k);
      sheet.getRange(newRowNum, nextCol).setValue(mapping[k]);
      nextCol++;
    }
  }
}

/**
 * 査読者が承諾した際のウェルカムメールとフォルダ準備
 */
function sendReviewerWelcomeEmail(reviewLog, ms, settings) {
  const ssId = getSpreadsheetId();
  const webAppUrl = ScriptApp.getService().getUrl();
  const dashboardUrl = webAppUrl + '?reviewKey=' + reviewLog.reviewKey;
  const reviewerName = reviewLog.Rev_Name || 'Reviewer';

  // 1. 共通査読資料フォルダの参照
  //    Settings の reviewMaterialsFolder キーで指定したフォルダ名を
  //    Journal Files 直下から探し、そのURLをメールに埋め込む。
  const rootName      = settings.SUBFOLDER || 'Journal Files';
  const materialName  = settings.reviewMaterialsFolder || '';
  let   materialUrl   = '';

  if (materialName) {
    try {
      const root = driveFolderCache.getRootFolder(rootName);
      if (root) {
        const matFolder = driveFolderCache.getFolderByName(root, materialName);
        if (matFolder) {
          materialUrl = matFolder.getUrl();
        }
      }
    } catch (e) {
      Logger.log('reviewMaterialsFolder lookup failed: ' + e.message);
    }
  }

  // Review_log に審査票フォルダURLを記録（リマインドメールでも参照される）
  if (materialUrl) {
    updateLogCell(ssId, REVIEW_LOG_SHEET_NAME, 'reviewKey', reviewLog.reviewKey, {
      'reviewMaterialsFolderUrl': materialUrl
    });
  }

  // 2. 査読期限
  const deadline = reviewLog.Review_Deadline || reviewLog.review_deadline || '';

  // 4. メール送信
  const subject = `[${settings.Journal_Name}] 査読ご就任の確認 / Reviewer Assignment Confirmed: ${ms.MsVer}`;

  const paperTitle = (ms.TitleJP && ms.TitleEN)
    ? ms.TitleJP + ' / ' + ms.TitleEN
    : (ms.TitleJP || ms.TitleEN || '');

  const deadlineRow = deadline
    ? `<tr>
         <th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">Review Deadline / 査読期限</th>
         <td style="padding:8px; border-bottom:1px solid #eee;"><strong>${deadline}</strong></td>
       </tr>`
    : '';

  const bodyHtml = `
    <p>
      Thank you for accepting the invitation to review the following manuscript.<br>
      以下の原稿の査読依頼をお引き受けいただき、誠にありがとうございます。
    </p>
    <table style="width:100%; font-size:14px; border-collapse:collapse; margin:20px 0;">
      <tr>
        <th style="text-align:left; padding:8px; border-bottom:1px solid #eee; width:35%;">Manuscript / 原稿</th>
        <td style="padding:8px; border-bottom:1px solid #eee;"><strong>${ms.MsVer}</strong></td>
      </tr>
      ${paperTitle ? `<tr>
        <th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">Title / タイトル</th>
        <td style="padding:8px; border-bottom:1px solid #eee;">${paperTitle}</td>
      </tr>` : ''}
      ${deadlineRow}
    </table>
    <p>
      Please open your reviewer dashboard via the button below and submit your review results by the deadline.<br>
      期限までに、下記ボタンから査読者専用ダッシュボードを開き、査読結果をご提出ください。
    </p>
  `;

  const html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting: `Dear Dr. ${reviewerName},`,
    bodyHtml: bodyHtml,
    buttonUrl: dashboardUrl,
    buttonLabel: 'Open Reviewer Dashboard / 査読者用メニュー',
    footerHtml: settings.mailFooter || ''
  });

  sendEmailSafe({ to: reviewLog.Rev_Email, subject, htmlBody: html },
    'Reviewer Welcome: ' + (ms.MsVer || '') + ' to ' + reviewerName);
}

function incrementHex(ssId, baseHex) {
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(REVIEW_LOG_SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = headers.indexOf('MsVerRevHex');
  
  let existing = [];
  if (colIndex !== -1 && sheet.getLastRow() > 1) {
    const colData = sheet.getRange(2, colIndex + 1, sheet.getLastRow() - 1).getValues();
    existing = colData.map(r => r[0]);
  }
  
  let base = baseHex.slice(0, -1);
  let lastDigit = baseHex.slice(-1);
  let incremented = (parseInt(lastDigit, 16) + 1).toString(16);
  
  while (existing.includes(base + incremented)) {
    incremented = (parseInt(incremented, 16) + 1).toString(16);
    if (parseInt(incremented, 16) > 0xf) {
      throw new Error("Cannot increment beyond f.");
    }
  }
  return base + incremented;
}
