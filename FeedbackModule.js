/**
 * FeedbackModule.js - 編集委員長から著者への最終判定 (SPA)
 */

function apiSubmitFeedback(data) {
  const ssId = getSpreadsheetId();
  const settings = getSettings();
  
  // 1. 原稿データを取得（著者キーまたは編集委員長専用キーで検索）
  let msData = getManuscriptData('author', data.msKey);
  if (!msData) msData = getManuscriptData('eic', data.msKey);
  if (!msData) throw new Error("Manuscript record not found.");
  
  const msVer = msData.MsVer;
  const msId = msData.MS_ID;
  const now = new Date();
  const todayNow = Utilities.formatDate(now, 'JST', 'yyyy/MM/dd HH:mm');
  const decisionTemplates = getDecisionTemplates(ssId, data.score);
  const isAccepted = decisionTemplates.isAccepted ? 'yes' : 'no';

  // 2a. 最終承認（IsAccepted=yes かつ Resubmit=no）の場合は MEルートへ強制転送
  //     著者・印刷担当者への通知は ME のチェック完了後に行われる
  if (decisionTemplates.isAccepted && !decisionTemplates.allowsResubmit) {
    return _redirectFeedbackToManagingEditorRoute(ssId, msData, data, decisionTemplates, settings);
  }

  // 2. 判定用フォルダを先に作成し、コメントPDF・EICファイルをまとめて格納する
  //    著者には 1 つのフォルダリンクのみ送付することで混乱を防ぐ
  const decisionFolder = getAuthorDecisionFolder(msData, settings);

  // 既存ファイルをクリア（再審査・差し替えに対応）
  const existingFiles = decisionFolder.getFiles();
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }

  // 3. コメントPDFを decisionFolder へ保存
  let commentPdfUrl = '';
  if (data.commentDocId) {
    try {
      const journalName = (settings && settings.Journal_Name) ? settings.Journal_Name : '';
      const nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');

      // ① 原本の先頭にヘッダーを挿入
      try {
        const commentDoc = DocumentApp.openById(data.commentDocId);
        insertCommentDocHeader(commentDoc.getBody(), journalName, data.score || '', nowStr);
        commentDoc.saveAndClose();
      } catch(eInsert) {
        Logger.log('Header insertion failed: ' + eInsert.message);
      }

      // ② 原本から PDF を取得（Drive API 経由）
      const pdfBlob = DriveApp.getFileById(data.commentDocId).getAs(MimeType.PDF);
      pdfBlob.setName('Open-Comments-' + msVer + '.pdf');

      // ③ PDF を decisionFolder に保存（working フォルダではなく著者向けフォルダへ統合）
      const savedPdf = decisionFolder.createFile(pdfBlob);
      savedPdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      commentPdfUrl = savedPdf.getUrl();
      if (data.commentEditorKey) {
        updateLogCell(ssId, EDITOR_LOG_SHEET_NAME, 'editorKey', data.commentEditorKey,
          { 'reportCommentPdfUrl': commentPdfUrl });
      }
    } catch(e) {
      Logger.log('Comment PDF export failed: ' + e.message);
    }
  }

  // 4. EIC 添付ファイルを decisionFolder へ保存
  if (data.files && data.files.length > 0) {
    data.files.forEach(file => {
      const blob = Utilities.newBlob(Utilities.base64Decode(file.content), file.mimeType, file.name);
      decisionFolder.createFile(blob);
    });
  }

  // フォルダに何か入っていれば共有して URL を取得
  let resultFolderUrl = 'nofile';
  if (commentPdfUrl || (data.files && data.files.length > 0)) {
    decisionFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    resultFolderUrl = decisionFolder.getUrl();
  }

  // 4. Drive 操作が全て成功した後に Manuscripts シートを更新
  updateManuscriptCell(ssId, msData.key, {
    'score': data.score,
    'openComments': data.openComments || '',
    'resultFolderUrl': resultFolderUrl, // ★ 公開用判定フォルダのURLを保存
    'sentBackAt': todayNow,
    'accepted': isAccepted
  });
  
  // msData オブジェクトに最新の判定結果をセット（後の通知メール生成で使用されるため）
  msData.score = data.score;
  
  // 5. 著者への通知メール送信
  sendFeedbackToAuthor(msData, data, resultFolderUrl, settings, ssId, decisionTemplates);
  
  // 6. 委員長（および担当編集者）への確認通知
  sendFeedbackConfirmationToRequester(msData, settings);
  
  writeLog(`Final Decision Submitted: ${msVer} - Score: ${data.score} (Author: ${msData.CA_Email})`);
  
  return { success: true };
}

/**
 * Manuscriptsシートの更新
 */
function updateManuscriptCell(ssId, msKey, updates) {
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(MANUSCRIPTS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const keyIdx = headers.indexOf('key');
  if (keyIdx === -1) throw new Error("Key column not found in Manuscripts sheet.");
  
  const rowIdx = data.findIndex(r => String(r[keyIdx]).trim() === String(msKey).trim());
  if (rowIdx === -1) throw new Error("Manuscript with key " + msKey + " not found.");
  
  Object.keys(updates).forEach(colName => {
    const colIdx = headers.indexOf(colName);
    if (colIdx !== -1) {
      const cell = sheet.getRange(rowIdx + 1, colIdx + 1);
      cell.setValue(updates[colName]);
      // 受理された場合は行に色を付ける (Legacy logic)
      if (colName === 'accepted' && updates[colName] === 'yes') {
        sheet.getRange(rowIdx + 1, 1, 1, headers.length).setBackground('#CCFFCC');
      }
    }
  });
}

/**
 * 著者への判定通知
 */
function sendFeedbackToAuthor(msData, data, resultFolderUrl, settings, ssId, decisionTemplates) {
  if (!decisionTemplates) decisionTemplates = getDecisionTemplates(ssId, data.score);
  const resubmissionUrl = ScriptApp.getService().getUrl() + '?key=' + msData.key; // 著者がアクセスすると状態により再投稿画面が出る想定

  // 期限日の計算 (Resubmittion_Limit が "8 weeks" などの形式を想定、デフォルトは 56日)
  const limitStr = String(settings.Resubmittion_Limit || '8 weeks').toLowerCase();
  const weeksMatch = limitStr.match(/(\d+)\s*weeks?/);
  const weeks = weeksMatch ? parseInt(weeksMatch[1]) : 8;
  const dueDateObj = new Date();
  dueDateObj.setDate(dueDateObj.getDate() + (weeks * 7));
  const dueDateStr = Utilities.formatDate(dueDateObj, 'JST', 'yyyy/MM/dd');

  // プレースホルダの置換
  const replacements = {
    'authorName': msData.CA_Name,
    'englishTitle': msData.TitleEN,
    'Journal_Name': settings.Journal_Name,
    'Resubmittion_Limit': settings.Resubmittion_Limit || '8 weeks',
    'manuscriptID': msData.MsVer,
    'Editor_Name': settings.Editor_Name || 'Editor-in-Chief',
    'dueDate': dueDateStr,
    'submissionLink': '(Provided by the button below / 下記ボタンよりアクセスしてください)',
    'formlink': '(Provided by the button below / 下記ボタンよりアクセスしてください)'
  };
  
  let mainText = replaceDecisionPlaceholders(decisionTemplates.mailText, replacements);
  
  // 【デザイン改修】 委員長コメントのセクション
  if (data.openComments) {
    mainText += `
      <div style="margin-top: 25px; padding: 18px 20px; background-color: #f8faf9; border-left: 5px solid #059669; border-radius: 6px; box-shadow: 0 1px 2px rgba(0,0,0,0.05); text-align: left;">
        <p style="margin: 0 0 10px 0; font-weight: bold; color: #065f46; font-size: 15px; border-bottom: 1px solid #d1fae5; padding-bottom: 6px;">🖋 編集部からのコメント / Comments from Editorial Board</p>
        <div style="font-size: 14.5px; color: #1e293b; line-height: 1.6; white-space: pre-wrap;">${data.openComments}</div>
      </div>
    `;
  }
  
  // 判定資料フォルダへのリンク（コメントPDF・EIC添付ファイルをまとめた 1 フォルダ）
  if (resultFolderUrl !== 'nofile') {
    mainText += `
      <div style="margin-top: 20px; padding: 20px; background-color: #f0f9ff; border: 1px solid #bae6fd; border-radius: 8px; text-align: center;">
        <p style="margin: 0 0 12px 0; font-weight: bold; color: #0369a1; font-size: 15px;">📁 判定資料・コメントPDF / Decision Materials &amp; Comments PDF</p>
        <p style="margin: 0 0 15px 0; font-size: 13.5px; color: #0c4a6e; line-height: 1.5;">編集委員長からの添付ファイル・オープンコメントPDFをご確認ください。<br>Please find the EIC's attached files and open-comments PDF in the shared folder below.</p>
        <a href="${resultFolderUrl}" style="display: inline-block; padding: 10px 24px; background-color: #2563eb; color: #ffffff !important; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 14.5px; box-shadow: 0 4px 6px -1px rgba(37, 99, 235, 0.2);">
          閲覧用フォルダを開く / Open Shared Folder
        </a>
      </div>
    `;
  }

  const bodyHtml = `
    <p>A decision has been reached regarding your manuscript <strong>${msData.MsVer}</strong>, entitled "${msData.TitleEN}".</p>
    <div style="background:#f1f5f9; padding:20px; border-radius:8px; margin:20px 0; font-size: 15px; line-height: 1.6;">
      ${mainText}
    </div>
    <hr style="border:none; border-top:1px solid #e2e8f0; margin: 20px 0;">
    <p>ご投稿いただいた原稿 <strong>${msData.MsVer}</strong> に対する判定をお送りいたします。</p>
  `;

  // 再投稿ボタンは Decisions シートの Resubmit 列が yes の場合のみ表示
  const showResubmitButton = !!decisionTemplates.allowsResubmit;

  const html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting: `Dear Dr. ${msData.CA_Name},`,
    bodyHtml: bodyHtml,
    buttonUrl:   showResubmitButton ? resubmissionUrl : null,
    buttonLabel: showResubmitButton ? 'Submit Revised Manuscript / 修正版を投稿する' : null,
    footerHtml: settings.mailFooter || ''
  });

  const subject = `[${settings.Journal_Name}] 原稿の審査結果について / Decision for your manuscript, ${msData.MsVer}: ${decisionTemplates.shortExplanation}`;

  // BCC: 担当編集者（path1 は ME を経由しないため担当編集者のみ）
  const editorBcc = _getAcceptedEditorEmail(ssId || getSpreadsheetId(), msData.MsVer || '');
  const feedbackMailOptions = {
    to:       msData.CA_Email,
    cc:       msData.ccEmails || '',
    subject:  subject,
    htmlBody: html
  };
  if (editorBcc && editorBcc.email) feedbackMailOptions.bcc = editorBcc.email;

  sendEmailSafe(feedbackMailOptions, 'Decision Feedback to Author: ' + msData.MsVer);
}

/**
 * テンプレート置換ユーティリティ
 */
function replaceDecisionPlaceholders(template, values) {
  if (!template) return '';
  return template.replace(/\$\{([^\}]+)\}/g, (match, key) => values[key] || match);
}

/**
 * 著者向けの判定フォルダを取得（隔離用）
 */
function getAuthorDecisionFolder(msData, settings) {
  const verFolder = getManuscriptVerFolder(msData, settings);
  const parentDecisionFolder = driveFolderCache.getOrCreateFolder(verFolder, 'decision');
  
  // 'for-author' というサブフォルダに隔離することで、内部用資料の混入を防ぐ
  return driveFolderCache.getOrCreateFolder(parentDecisionFolder, 'for-author');
}

/**
 * 判定フォルダの取得（内部用・互換性維持のため残す）
 */
function getDecisionFolder(msData, settings) {
  const verFolder = getManuscriptVerFolder(msData, settings);
  return driveFolderCache.getOrCreateFolder(verFolder, 'decision');
}

/**
 * Decisions シートの選択肢一覧を返す（EICフォームのドロップダウン用）
 * ShortExplanation 列の値をキーとして返す
 * @returns {Array<{value: string}>}
 */
function getDecisionMailOptions(ssId) {
  const sheet = _getDecisionsSheet(ssId);
  if (!sheet) return [];
  const parsed = _findDecisionsSheetRows(sheet);
  if (!parsed) return [];
  const { headerRowIdx, headers, data } = parsed;
  const findIdx = kw => headers.findIndex(h => String(h).toLowerCase().trim() === kw.toLowerCase());
  const sIdx = findIdx('ShortExplanation');
  if (sIdx === -1) return [];
  const aIdx = findIdx('IsAccepted');
  const rIdx = findIdx('Resubmit');
  return data.slice(headerRowIdx + 1)
    .filter(r => String(r[sIdx]).trim() !== '')
    .map(r => ({
      value:          String(r[sIdx]).trim(),
      isAccepted:     aIdx !== -1 && String(r[aIdx] || '').trim().toLowerCase() === 'yes',
      allowsResubmit: rIdx !== -1 && String(r[rIdx] || '').trim().toLowerCase() === 'yes'
    }));
}

/**
 * Decisions シートから判定テンプレートを取得
 * ShortExplanation をキーとして検索し、
 * { shortExplanation, mailText, isAccepted, allowsResubmit } を返す
 */
function getDecisionTemplates(ssId, score) {
  const fallback = {
    shortExplanation: score || 'Decision Reached',
    mailText:         'Please check the system.',
    isAccepted:       false,
    allowsResubmit:   false
  };

  const sheet = _getDecisionsSheet(ssId);
  if (!sheet) {
    writeLog('[WARN] getDecisionTemplates: Decisions 関連シートが見つかりません。シート名や ShortExplanation 行を確認してください。');
    return fallback;
  }

  const parsed = _findDecisionsSheetRows(sheet);
  if (!parsed) {
    writeLog('[WARN] getDecisionTemplates: シート "' + DECISION_MAIL_SHEET_NAME + '" に ShortExplanation ヘッダー行が見つかりません。');
    return fallback;
  }
  const { headerRowIdx, headers, data } = parsed;

  const findHeaderIdx = (keyword) => {
    const kw = keyword.toLowerCase();
    return headers.findIndex(h => String(h).toLowerCase().trim() === kw);
  };

  const sIdx = findHeaderIdx('ShortExplanation');
  const aIdx = findHeaderIdx('IsAccepted');
  const rIdx = findHeaderIdx('Resubmit');
  const tIdx = findHeaderIdx('Mail text');

  if (sIdx === -1) return fallback;

  const row = data.slice(headerRowIdx + 1).find(r => String(r[sIdx]).trim() === String(score).trim());
  if (!row) {
    writeLog('[WARN] getDecisionTemplates: score="' + score + '" に一致する行が Decisions シートに見つかりません。');
    return fallback;
  }

  const acceptedValue = aIdx !== -1 ? String(row[aIdx] || '').trim().toLowerCase() : '';
  const isAcceptedBool = acceptedValue === 'yes' || acceptedValue === 'true' || acceptedValue === '1';

  const resubmitValue = rIdx !== -1 ? String(row[rIdx] || '').trim().toLowerCase() : '';
  const allowsResubmitBool = resubmitValue === 'yes' || resubmitValue === 'true' || resubmitValue === '1';

  return {
    shortExplanation: String(row[sIdx] || '').trim(),
    mailText:         tIdx !== -1 ? String(row[tIdx] || '').trim() : '',
    isAccepted:       isAcceptedBool,
    allowsResubmit:   allowsResubmitBool
  };
}

/**
 * Decisions シートを参照して、そのスコアが受理扱いかどうかを返す
 * （旧: Settings H5:I15 参照）
 */
function isScoreAccepted(ssId, score) {
  try {
    const templates = getDecisionTemplates(ssId, score);
    writeLog('isScoreAccepted: score="' + score + '" → isAccepted=' + templates.isAccepted);
    return templates.isAccepted;
  } catch(e) {
    writeLog('[ERROR] isScoreAccepted: ' + e.message);
    return false;
  }
}

/**
 * 送信完了後の確認通知
 */
function sendFeedbackConfirmationToRequester(msData, settings) {
  const paperTitle = (msData.TitleJP && msData.TitleEN) ? msData.TitleJP + ' / ' + msData.TitleEN : (msData.TitleJP || msData.TitleEN || 'Unknown Title');

  // 委員長宛（自分自身への確認）
  const subject = `[${settings.Journal_Name}] 判定通知送信の確認 / Confirmation: Decision sent to author(s): ${msData.MsVer}`;
  const bodyHtml = `
    <p>This is a confirmation email that the decision for manuscript <strong>${msData.MsVer}</strong> has been sent to the author(s).</p>
    <p>原稿 <strong>${msData.MsVer}</strong> に対する著者への判定通知が完了いたしました。</p>
    <table style="width:100%; font-size: 14px; border-collapse: collapse; margin: 20px 0;">
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee; width: 30%;">Manuscript / 原稿</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${msData.MsVer || ''}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Type / 種別</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${msData.MS_Type || ''}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Title / タイトル</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${paperTitle}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Corresponding Author / 責任著者</th><td style="padding: 8px; border-bottom: 1px solid #eee;">${msData.CA_Name || ''}</td></tr>
      <tr><th style="text-align:left; padding: 8px; border-bottom: 1px solid #eee;">Final Decision / 判定結果</th><td style="padding: 8px; border-bottom: 1px solid #eee; font-weight:bold;">${msData.score || ''}</td></tr>
    </table>
  `;

  const html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting: `Journal Editorial Board,`,
    bodyHtml: bodyHtml,
    footerHtml: settings.mailFooter || ''
  });

  if (!settings.chiefEditorEmail) {
    Logger.log('sendFeedbackConfirmationToRequester: chiefEditorEmail が設定されていないためスキップします (MsVer: ' + msData.MsVer + ')');
    return;
  }

  sendEmailSafe({ to: settings.chiefEditorEmail, subject, htmlBody: html },
    'Feedback Confirmation to EIC: ' + msData.MsVer);
}

/**
 * 著者への判定通知メールのプレビューを生成する (API)
 */
function apiGetFeedbackPreview(data) {
  const ssId = getSpreadsheetId();
  const settings = getSettings();
  
  // 1. 原稿データを取得
  let msData = getManuscriptData('author', data.msKey);
  if (!msData) msData = getManuscriptData('eic', data.msKey);
  if (!msData) throw new Error("Manuscript record not found.");
  
  // 2. テンプレートの取得
  const decisionTemplates = getDecisionTemplates(ssId, data.score);
  
  // 3. プレースホルダの置換
  const limitStr = String(settings.Resubmittion_Limit || '8 weeks').toLowerCase();
  const weeksMatch = limitStr.match(/(\d+)\s*weeks?/);
  const weeks = weeksMatch ? parseInt(weeksMatch[1]) : 8;
  const dueDateObj = new Date();
  dueDateObj.setDate(dueDateObj.getDate() + (weeks * 7));
  const dueDateStr = Utilities.formatDate(dueDateObj, 'JST', 'yyyy/MM/dd');

  const replacements = {
    'authorName': msData.CA_Name,
    'englishTitle': msData.TitleEN,
    'Journal_Name': settings.Journal_Name,
    'Resubmittion_Limit': settings.Resubmittion_Limit || '8 weeks',
    'manuscriptID': msData.MsVer,
    'Editor_Name': settings.Editor_Name || 'Editor-in-Chief',
    'dueDate': dueDateStr,
    'submissionLink': '(Provided by the button below / 下記ボタンよりアクセスしてください)',
    'formlink': '(Provided by the button below / 下記ボタンよりアクセスしてください)'
  };
  
  let mainText = replaceDecisionPlaceholders(decisionTemplates.mailText, replacements);
  
  // 【デザイン改修：プレビュー版】 委員長コメント
  if (data.openComments) {
    mainText += `
      <div style="margin-top: 25px; padding: 18px 20px; background-color: #f8faf9; border-left: 5px solid #059669; border-radius: 6px; box-shadow: 0 1px 2px rgba(0,0,0,0.05); text-align: left;">
        <p style="margin: 0 0 10px 0; font-weight: bold; color: #065f46; font-size: 15px; border-bottom: 1px solid #d1fae5; padding-bottom: 6px;">🖋 編集部からのコメント / Comments from Editorial Board</p>
        <div style="font-size: 14.5px; color: #1e293b; line-height: 1.6; white-space: pre-wrap;">${data.openComments}</div>
      </div>
    `;
  }
  
  // 【デザイン改修：プレビュー版】 判定資料フォルダへのリンク
  mainText += `
    <div style="margin-top: 20px; padding: 20px; background-color: #f0f9ff; border: 1px solid #bae6fd; border-radius: 8px; text-align: center;">
      <p style="margin: 0 0 12px 0; font-weight: bold; color: #0369a1; font-size: 15px;">📁 判定資料の確認 / Attached decision materials</p>
      <p style="margin: 0 0 15px 0; font-size: 13.5px; color: #0c4a6e; line-height: 1.5;">判定の根拠となる資料フォルダへアクセスしてください。<br>Access the shared folder for further details of the decision.</p>
      <div style="display: inline-block; padding: 10px 24px; background-color: #2563eb; color: #ffffff !important; border-radius: 6px; font-weight: bold; font-size: 14.5px; opacity: 0.8;">
        (Review folders URL will be inserted here / ここに資料フォルダのURLが挿入されます)
      </div>
    </div>
  `;

  const bodyHtml = `
    <p>A decision has been reached regarding your manuscript <strong>${msData.MsVer}</strong>, entitled "${msData.TitleEN}".</p>
    <div style="background:#f1f5f9; padding:20px; border-radius:8px; margin:20px 0; font-size: 15px; line-height: 1.6;">
      ${mainText}
    </div>
    <hr style="border:none; border-top:1px solid #e2e8f0; margin: 20px 0;">
    <p>ご投稿いただいた原稿 <strong>${msData.MsVer}</strong> に対する判定をお送りいたします。</p>
  `;

  const showResubmitButton = !!decisionTemplates.allowsResubmit;

  const html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting: `Dear Dr. ${msData.CA_Name},`,
    bodyHtml: bodyHtml,
    buttonUrl:   showResubmitButton ? '#' : null,
    buttonLabel: showResubmitButton ? 'Submit Revised Manuscript / 修正版を投稿する' : null,
    footerHtml: settings.mailFooter || ''
  });

  return {
    html: html,
    subject: `[${settings.Journal_Name}] 原稿の審査結果について / Decision for your manuscript, ${msData.MsVer}: ${decisionTemplates.shortExplanation}`
  };
}

/**
 * 通常ルートで「最終承認（IsAccepted=yes, Resubmit=no）」が選ばれた場合に
 * ME（編集幹事）ルートへ強制転送する処理。
 * 著者・印刷担当者への通知は MEのチェック完了後に行われるため、ここでは行わない。
 */
function _redirectFeedbackToManagingEditorRoute(ssId, msData, data, decisionTemplates, settings) {
  if (!settings.managingEditorEmail) {
    throw new Error(
      'この判定（最終承認・再投稿なし）を処理するには、Settings に managingEditorEmail が設定されている必要があります。 / ' +
      'managingEditorEmail must be configured in Settings to process a final acceptance decision.'
    );
  }

  // EIC がアップロードしたファイルを判定フォルダへ保存
  const decisionFolder = getAuthorDecisionFolder(msData, settings);
  const existingFiles = decisionFolder.getFiles();
  while (existingFiles.hasNext()) existingFiles.next().setTrashed(true);

  let folderUrl = '';
  const attachmentBlobs = [];
  if (data.files && data.files.length > 0) {
    data.files.forEach(file => {
      const blob = Utilities.newBlob(Utilities.base64Decode(file.content), file.mimeType, file.name);
      decisionFolder.createFile(blob);
      attachmentBlobs.push(blob);
    });
    decisionFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    folderUrl = decisionFolder.getUrl();
  }

  // managingEditorKey を生成してシートを更新（MEダッシュボードを有効化）
  const managingEditorKey = Utilities.getUuid();
  updateManuscriptCell(ssId, msData.key, {
    'managingEditorKey': managingEditorKey,
    'finalStatus':       'final_review',
    'resultFolderUrl':   folderUrl || 'nofile'
  });

  // ME（編集幹事）へ通知メールを送信
  _sendEicAcceptanceToManagingEditor(msData, data, decisionTemplates, settings, managingEditorKey, attachmentBlobs, folderUrl);

  writeLog('[ME Redirect] apiSubmitFeedback: ' + msData.MsVer + ' → 最終承認のため ME ルートへ転送。managingEditorKey 生成済み。');

  return { redirectedToME: true };
}

/**
 * 通常ルートからMEルートへ転送した旨を編集幹事（ME）へ通知する。
 * 担当編集者からの推薦とは異なり、EICが直接承認判定を選んだケースなので
 * PDF/Word レポート添付はなく、EIC のコメントと任意のファイルのみを送付する。
 */
function _sendEicAcceptanceToManagingEditor(msData, data, decisionTemplates, settings, managingEditorKey, attachmentBlobs, folderUrl) {
  const webAppUrl = ScriptApp.getService().getUrl();
  const meLink = webAppUrl + '?managingEditorKey=' + managingEditorKey;
  const paperTitle = (msData.TitleJP && msData.TitleEN)
    ? msData.TitleJP + ' / ' + msData.TitleEN
    : (msData.TitleJP || msData.TitleEN || '');

  const fileLinkRow = folderUrl
    ? `<tr>
         <th style="text-align:left; padding:8px; border-bottom:1px solid #eee; width:30%;">EIC 添付資料 / EIC Files</th>
         <td style="padding:8px; border-bottom:1px solid #eee;"><a href="${folderUrl}">フォルダを開く / Open Folder</a></td>
       </tr>`
    : '';

  const eicCommentHtml = data.openComments
    ? `<div style="margin-top:16px; padding:14px 16px; background:#f8faf9; border-left:4px solid #059669; border-radius:6px;">
         <p style="margin:0 0 6px; font-weight:bold; color:#065f46; font-size:13px;">EIC コメント / EIC Comments</p>
         <p style="margin:0; font-size:14px; color:#1e293b; white-space:pre-wrap;">${data.openComments}</p>
       </div>`
    : '';

  const bodyHtml = `
    <p>The Editor-in-Chief (EIC) has selected a <strong>final acceptance (no resubmission)</strong> decision for the manuscript below, and the process has been forwarded to the Managing Editor route. Please open the Managing Editor dashboard using the button below.</p>
    <p>編集委員長（EIC）が以下の原稿に対して <strong>最終承認（受理・再投稿なし）</strong> の判定を選択したため、編集幹事ルートへ転送されました。以下のボタンより編集幹事ダッシュボードを開き、確認作業をお願いいたします。</p>

    <table style="width:100%; font-size:14px; border-collapse:collapse; margin:20px 0;">
      <tr>
        <th style="text-align:left; padding:8px; border-bottom:1px solid #eee; width:30%;">原稿番号 / MS ID</th>
        <td style="padding:8px; border-bottom:1px solid #eee;">${msData.MsVer}</td>
      </tr>
      <tr>
        <th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">タイトル / Title</th>
        <td style="padding:8px; border-bottom:1px solid #eee;">${paperTitle}</td>
      </tr>
      <tr>
        <th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">EIC の判定 / EIC Decision</th>
        <td style="padding:8px; border-bottom:1px solid #eee; font-weight:bold;">${decisionTemplates.shortExplanation}</td>
      </tr>
      <tr>
        <th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">責任著者 / Corresponding Author</th>
        <td style="padding:8px; border-bottom:1px solid #eee;">${msData.CA_Name || ''}</td>
      </tr>
      ${fileLinkRow}
    </table>

    ${eicCommentHtml}

    <div style="margin-top:20px; padding:14px 16px; background:#fef3c7; border:1px solid #f59e0b; border-radius:8px;">
      <p style="margin:0 0 6px; font-weight:bold; color:#92400e; font-size:13px;">⚠️ 重要 / Important</p>
      <p style="margin:0; font-size:13.5px; color:#78350f; line-height:1.6;">
        The author and production editor have NOT yet been notified at this stage.<br>
        Please review the manuscript via the Managing Editor dashboard. After your submission, the Editor-in-Chief will complete the final step (Route B: Send to Production Editor).
      </p>
      <p style="margin:8px 0 0; font-size:13.5px; color:#78350f; line-height:1.6;">
        この時点では著者・印刷担当者にはまだ通知されていません。<br>
        編集幹事ダッシュボードでご確認のうえ、送信手続きを完了してください。その後、編集委員長が最終的な送付操作（Route B: 印刷担当者へ）を行います。
      </p>
    </div>
  `;

  const html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting: '編集幹事 殿 / Dear Managing Editor,',
    bodyHtml: bodyHtml,
    buttonUrl: meLink,
    buttonLabel: '編集幹事ダッシュボードを開く / Open Managing Editor Dashboard',
    footerHtml: settings.mailFooter || ''
  });

  const mailOptions = {
    to:       settings.managingEditorEmail,
    subject:  `[${settings.Journal_Name}] 受理原稿の最終確認依頼（EIC直接承認） / Final Review Request (EIC Direct Acceptance): ${msData.MsVer}`,
    htmlBody: html
  };
  if (attachmentBlobs && attachmentBlobs.length > 0) {
    mailOptions.attachments = attachmentBlobs;
  }

  sendEmailSafe(mailOptions, 'EIC Acceptance → ME Route: ' + msData.MsVer);
}
