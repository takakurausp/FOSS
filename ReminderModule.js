/**
 * ReminderModule.js - 査読者・担当編集者へのリマインドメール
 *
 * 【トリガー設定】
 * GASエディタ → トリガー → checkReminders → 時間主導型 → 毎日（午前9時など）
 *
 * 【Settingsシートに必要な設定】
 * firstReminderDays  : 依頼から何日後に1通目を送るか（例: 7）
 * secondReminderDays : 依頼から何日後に2通目を送るか（例: 14）
 * thirdReminderDays  : 依頼から何日後に最終通知を送るか（例: 21）
 *
 * 【ログシートに必要な列】
 * Editor_log  : Reminder1_At, Reminder2_At, Reminder3_At
 * Review_log  : Reminder1_At, Reminder2_At, Reminder3_At
 * （列が存在しない場合はリマインド送信済みフラグが記録されないため、毎日再送されます。
 *   必ず各シートに上記3列を追加してください。）
 */

/**
 * メインエントリ。毎日1回トリガーで呼び出す。
 */
function checkReminders() {
  const ssId = getSpreadsheetId();
  const settings = getSettings();

  const first  = parseInt(settings.firstReminderDays  || '7',  10);
  const second = parseInt(settings.secondReminderDays || '14', 10);
  const third  = parseInt(settings.thirdReminderDays  || '21', 10);
  const days = { first, second, third };

  Logger.log('checkReminders start — thresholds: ' + first + '/' + second + '/' + third + ' days');

  checkEditorReminders(ssId, settings, days);
  checkReviewerInvitationReminders(ssId, settings, days);
  checkReviewerSubmissionReminders(ssId, settings, days);

  Logger.log('checkReminders end');
}

// ─────────────────────────────────────────
// 担当編集者候補への未回答リマインド
// ─────────────────────────────────────────
function checkEditorReminders(ssId, settings, days) {
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(EDITOR_LOG_SHEET_NAME);
  if (!sheet) return;

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).toLowerCase().trim());
  const col     = name => headers.indexOf(name.toLowerCase().trim());

  const edtOkIdx       = col('edtok');
  const askAtIdx       = col('ask_at');
  const editorKeyIdx   = col('editorkey');
  const editorNameIdx  = col('editor_name');
  const editorEmailIdx = col('editor_email');
  const msVerIdx       = col('msver');
  const rem1Idx        = col('reminder1_at');
  const rem2Idx        = col('reminder2_at');
  const rem3Idx        = col('reminder3_at');

  if (edtOkIdx === -1 || askAtIdx === -1 || editorKeyIdx === -1) return;

  const today = new Date();

  for (let i = 1; i < data.length; i++) {
    const row    = data[i];
    const edtOk  = String(row[edtOkIdx] || '').trim();
    if (edtOk !== '') continue; // 承諾または辞退済みはスキップ

    const askAt = row[askAtIdx];
    if (!askAt) continue;
    const askDate = new Date(askAt);
    if (isNaN(askDate)) continue;
    const elapsed = Math.floor((today - askDate) / (1000 * 60 * 60 * 24));

    const editorKey   = editorKeyIdx   !== -1 ? String(row[editorKeyIdx]   || '').trim() : '';
    const editorName  = editorNameIdx  !== -1 ? String(row[editorNameIdx]  || '').trim() : '';
    const editorEmail = editorEmailIdx !== -1 ? String(row[editorEmailIdx] || '').trim() : '';
    const msVer       = msVerIdx       !== -1 ? String(row[msVerIdx]       || '').trim() : '';
    if (!editorKey || !editorEmail) continue;

    const rem1 = rem1Idx !== -1 ? String(row[rem1Idx] || '').trim() : '';
    const rem2 = rem2Idx !== -1 ? String(row[rem2Idx] || '').trim() : '';
    const rem3 = rem3Idx !== -1 ? String(row[rem3Idx] || '').trim() : '';

    const level = calcReminderLevel(elapsed, days, rem1, rem2, rem3);
    if (level === 0) continue;

    sendEditorReminderEmail(editorEmail, editorName, editorKey, msVer, elapsed, level, settings);

    const now     = Utilities.formatDate(today, 'JST', 'yyyy/MM/dd HH:mm');
    const colName = ['', 'Reminder1_At', 'Reminder2_At', 'Reminder3_At'][level];
    updateLogCell(ssId, EDITOR_LOG_SHEET_NAME, 'editorKey', editorKey, { [colName]: now });
    Logger.log('Editor reminder level ' + level + ' sent to ' + editorEmail + ' for ' + msVer);
  }
}

// ─────────────────────────────────────────
// 査読者候補への未回答リマインド
// ─────────────────────────────────────────
function checkReviewerInvitationReminders(ssId, settings, days) {
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(REVIEW_LOG_SHEET_NAME);
  if (!sheet) return;

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).toLowerCase().trim());
  const col     = name => headers.indexOf(name.toLowerCase().trim());

  const revOkIdx     = col('revok');
  const askAtIdx     = col('ask_at');
  const reviewKeyIdx = col('reviewkey');
  const revNameIdx   = col('rev_name');
  const revEmailIdx  = col('rev_email');
  const msVerIdx     = col('msver');
  const rem1Idx      = col('reminder1_at');
  const rem2Idx      = col('reminder2_at');
  const rem3Idx      = col('reminder3_at');

  if (revOkIdx === -1 || askAtIdx === -1 || reviewKeyIdx === -1) return;

  const today = new Date();

  for (let i = 1; i < data.length; i++) {
    const row   = data[i];
    const revOk = String(row[revOkIdx] || '').trim();
    if (revOk !== '') continue; // 承諾または辞退済みはスキップ

    const askAt = row[askAtIdx];
    if (!askAt) continue;
    const askDate = new Date(askAt);
    if (isNaN(askDate)) continue;
    const elapsed = Math.floor((today - askDate) / (1000 * 60 * 60 * 24));

    const reviewKey  = reviewKeyIdx !== -1 ? String(row[reviewKeyIdx] || '').trim() : '';
    const revName    = revNameIdx   !== -1 ? String(row[revNameIdx]   || '').trim() : '';
    const revEmail   = revEmailIdx  !== -1 ? String(row[revEmailIdx]  || '').trim() : '';
    const msVer      = msVerIdx     !== -1 ? String(row[msVerIdx]     || '').trim() : '';
    if (!reviewKey || !revEmail) continue;

    const rem1 = rem1Idx !== -1 ? String(row[rem1Idx] || '').trim() : '';
    const rem2 = rem2Idx !== -1 ? String(row[rem2Idx] || '').trim() : '';
    const rem3 = rem3Idx !== -1 ? String(row[rem3Idx] || '').trim() : '';

    const level = calcReminderLevel(elapsed, days, rem1, rem2, rem3);
    if (level === 0) continue;

    sendReviewerInvitationReminderEmail(revEmail, revName, reviewKey, msVer, elapsed, level, settings);

    const now     = Utilities.formatDate(today, 'JST', 'yyyy/MM/dd HH:mm');
    const colName = ['', 'Reminder1_At', 'Reminder2_At', 'Reminder3_At'][level];
    updateLogCell(ssId, REVIEW_LOG_SHEET_NAME, 'reviewKey', reviewKey, { [colName]: now });
    Logger.log('Reviewer invitation reminder level ' + level + ' sent to ' + revEmail + ' for ' + msVer);
  }
}

// ─────────────────────────────────────────
// 査読承諾済み・未提出の査読者へのリマインド
// ─────────────────────────────────────────
function checkReviewerSubmissionReminders(ssId, settings, days) {
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(REVIEW_LOG_SHEET_NAME);
  if (!sheet) return;

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).toLowerCase().trim());
  const col     = name => headers.indexOf(name.toLowerCase().trim());

  const revOkIdx      = col('revok');
  const rcvAtIdx      = col('received_at');
  const answerAtIdx   = col('answer_at');
  const deadlineIdx   = col('review_deadline');
  const reviewKeyIdx  = col('reviewkey');
  const revNameIdx    = col('rev_name');
  const revEmailIdx   = col('rev_email');
  const msVerIdx      = col('msver');
  const folderUrlIdx  = col('reviewmaterialsfolderurl');
  const rem1Idx       = col('reminder1_at');
  const rem2Idx       = col('reminder2_at');
  const rem3Idx       = col('reminder3_at');

  if (revOkIdx === -1 || reviewKeyIdx === -1) return;

  const today = new Date();

  for (let i = 1; i < data.length; i++) {
    const row   = data[i];
    const revOk = String(row[revOkIdx]  || '').trim();
    const rcvAt = rcvAtIdx !== -1 ? String(row[rcvAtIdx] || '').trim() : '';
    if (revOk !== 'ok') continue; // 未承諾・辞退はスキップ
    if (rcvAt !== '')   continue; // 提出済みはスキップ

    // 承諾日（Answer_At）を基準日とする。なければ Ask_At を使用。
    const baseRaw = (answerAtIdx !== -1 && row[answerAtIdx])
      ? row[answerAtIdx] : (col('ask_at') !== -1 ? row[col('ask_at')] : null);
    if (!baseRaw) continue;
    const baseDate = new Date(baseRaw);
    if (isNaN(baseDate)) continue;
    const elapsed = Math.floor((today - baseDate) / (1000 * 60 * 60 * 24));

    const reviewKey  = reviewKeyIdx !== -1 ? String(row[reviewKeyIdx] || '').trim() : '';
    const revName    = revNameIdx   !== -1 ? String(row[revNameIdx]   || '').trim() : '';
    const revEmail   = revEmailIdx  !== -1 ? String(row[revEmailIdx]  || '').trim() : '';
    const msVer      = msVerIdx     !== -1 ? String(row[msVerIdx]     || '').trim() : '';
    const deadline   = deadlineIdx  !== -1 ? String(row[deadlineIdx]  || '').trim() : '';
    const folderUrl  = folderUrlIdx !== -1 ? String(row[folderUrlIdx] || '').trim() : '';
    if (!reviewKey || !revEmail) continue;

    // 提出フェーズ用リマインドは専用列 SubReminder1_At 等があればそちらを使う。
    // なければ Reminder1_At 等を流用する（招待フェーズで既に送っていても再利用）。
    const subRem1Idx = col('subreminder1_at');
    const subRem2Idx = col('subreminder2_at');
    const subRem3Idx = col('subreminder3_at');
    const useSubCol  = subRem1Idx !== -1;

    const rem1 = useSubCol ? String(row[subRem1Idx] || '').trim()
                           : (rem1Idx !== -1 ? String(row[rem1Idx] || '').trim() : '');
    const rem2 = useSubCol ? String(row[subRem2Idx] || '').trim()
                           : (rem2Idx !== -1 ? String(row[rem2Idx] || '').trim() : '');
    const rem3 = useSubCol ? String(row[subRem3Idx] || '').trim()
                           : (rem3Idx !== -1 ? String(row[rem3Idx] || '').trim() : '');

    const level = calcReminderLevel(elapsed, days, rem1, rem2, rem3);
    if (level === 0) continue;

    sendReviewerSubmissionReminderEmail(
      revEmail, revName, reviewKey, msVer, deadline, folderUrl, elapsed, level, settings
    );

    const now     = Utilities.formatDate(today, 'JST', 'yyyy/MM/dd HH:mm');
    const colName = ['', 'Reminder1_At', 'Reminder2_At', 'Reminder3_At'][level];
    const subColName = ['', 'SubReminder1_At', 'SubReminder2_At', 'SubReminder3_At'][level];
    updateLogCell(ssId, REVIEW_LOG_SHEET_NAME, 'reviewKey', reviewKey,
      { [useSubCol ? subColName : colName]: now });
    Logger.log('Reviewer submission reminder level ' + level + ' sent to ' + revEmail + ' for ' + msVer);
  }
}

// ─────────────────────────────────────────
// ユーティリティ
// ─────────────────────────────────────────

/**
 * 経過日数と送信済みフラグから送信すべきレベル (1/2/3) を返す。
 * 送信不要なら 0 を返す。
 */
function calcReminderLevel(elapsed, days, rem1, rem2, rem3) {
  if (elapsed >= days.third  && !rem3) return 3;
  if (elapsed >= days.second && !rem2) return 2;
  if (elapsed >= days.first  && !rem1) return 1;
  return 0;
}

// ─────────────────────────────────────────
// メール送信関数
// ─────────────────────────────────────────

function sendEditorReminderEmail(toEmail, toName, editorKey, msVer, elapsed, level, settings) {
  const webAppUrl   = ScriptApp.getService().getUrl();
  const responseUrl = webAppUrl + '?editorKey=' + editorKey;
  const isFinal     = level >= 3;

  const subject = isFinal
    ? `[${settings.Journal_Name}] [最終リマインド / Final Reminder] 担当編集者依頼への回答をお願いします / Please respond to editor assignment invitation — ${msVer}`
    : `[${settings.Journal_Name}] [リマインド ${level} / Reminder ${level}] 担当編集者ご就任のご依頼 / Editor assignment invitation — ${msVer}`;

  const bodyHtml = `
    <p>${isFinal ? '<strong style="color:#dc2626;">[Final Reminder]</strong> ' : ''}You received an invitation to serve as a responsible editor for manuscript <strong>${msVer}</strong>, but we have not yet received your response.</p>
    <p>It has been <strong>${elapsed} days</strong> since the invitation was sent. Please respond at your earliest convenience by clicking the button below.</p>
    <hr style="border:none; border-top:1px solid #e2e8f0; margin:20px 0;">
    <p>${isFinal ? '<strong>【最終リマインド】</strong>' : ''}原稿 <strong>${msVer}</strong> の担当編集者ご就任のご依頼をお送りしてから <strong>${elapsed}日</strong> が経過しておりますが、まだご回答をいただいておりません。お手数ですが、以下のボタンよりご回答ください。</p>
  `;

  const html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting:    `Dear Dr. ${toName},`,
    bodyHtml:    bodyHtml,
    buttonUrl:   responseUrl,
    buttonLabel: 'Respond to Invitation / 依頼に回答する',
    footerHtml:  settings.mailFooter || ''
  });

  sendEmailSafe({ to: toEmail, subject, htmlBody: html },
    'Editor Reminder L' + level + ': ' + msVer + ' to ' + toName);
}

function sendReviewerInvitationReminderEmail(toEmail, toName, reviewKey, msVer, elapsed, level, settings) {
  const webAppUrl   = ScriptApp.getService().getUrl();
  const responseUrl = webAppUrl + '?reviewKey=' + reviewKey;
  const isFinal     = level >= 3;

  const subject = isFinal
    ? `[${settings.Journal_Name}] [最終リマインド / Final Reminder] 査読依頼への回答をお願いします / Please respond to reviewer invitation — ${msVer}`
    : `[${settings.Journal_Name}] [リマインド ${level} / Reminder ${level}] 査読のご依頼 / Reviewer invitation — ${msVer}`;

  const bodyHtml = `
    <p>${isFinal ? '<strong style="color:#dc2626;">[Final Reminder]</strong> ' : ''}You received an invitation to review manuscript <strong>${msVer}</strong>, but we have not yet received your response.</p>
    <p>It has been <strong>${elapsed} days</strong> since the invitation was sent. Please respond by clicking the button below.</p>
    <hr style="border:none; border-top:1px solid #e2e8f0; margin:20px 0;">
    <p>${isFinal ? '<strong>【最終リマインド】</strong>' : ''}原稿 <strong>${msVer}</strong> の査読依頼をお送りしてから <strong>${elapsed}日</strong> が経過しておりますが、まだご回答をいただいておりません。以下のボタンよりご回答ください。</p>
  `;

  const html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting:    `Dear Dr. ${toName},`,
    bodyHtml:    bodyHtml,
    buttonUrl:   responseUrl,
    buttonLabel: 'Respond to Invitation / 依頼に回答する',
    footerHtml:  settings.mailFooter || ''
  });

  sendEmailSafe({ to: toEmail, subject, htmlBody: html },
    'Reviewer Invitation Reminder L' + level + ': ' + msVer + ' to ' + toName);
}

function sendReviewerSubmissionReminderEmail(
  toEmail, toName, reviewKey, msVer, deadline, folderUrl, elapsed, level, settings
) {
  const webAppUrl      = ScriptApp.getService().getUrl();
  const reviewMenuUrl  = webAppUrl + '?reviewKey=' + reviewKey;
  const isFinal        = level >= 3;

  const subject = isFinal
    ? `[${settings.Journal_Name}] [最終リマインド / Final Reminder] 査読結果の提出をお願いします / Please submit your review — ${msVer}`
    : `[${settings.Journal_Name}] [リマインド ${level} / Reminder ${level}] 査読結果提出期限 / Review submission due — ${msVer}`;

  const deadlineRow = deadline
    ? `<tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee; width:30%;">Review Deadline / 査読期限</th>
       <td style="padding:8px; border-bottom:1px solid #eee;"><strong>${deadline}</strong></td></tr>`
    : '';

  const folderRow = folderUrl
    ? `<tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">Review Materials / 査読資料</th>
       <td style="padding:8px; border-bottom:1px solid #eee;"><a href="${folderUrl}" target="_blank">査読資料フォルダを開く / Open Review Materials</a></td></tr>`
    : '';

  const bodyHtml = `
    <p>${isFinal ? '<strong style="color:#dc2626;">[Final Reminder]</strong> ' : ''}This is a reminder that your review for manuscript <strong>${msVer}</strong> has not yet been submitted.</p>
    <table style="width:100%; font-size:14px; border-collapse:collapse; margin:20px 0;">
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee; width:30%;">Manuscript / 原稿</th>
          <td style="padding:8px; border-bottom:1px solid #eee;">${msVer}</td></tr>
      ${deadlineRow}
      ${folderRow}
    </table>
    <p>Please submit your review results via the button below.</p>
    <hr style="border:none; border-top:1px solid #e2e8f0; margin:20px 0;">
    <p>${isFinal ? '<strong>【最終リマインド】</strong>' : ''}原稿 <strong>${msVer}</strong> の査読結果がまだ提出されておりません。お手数ですが、以下のボタンより査読結果をご提出ください。</p>
  `;

  const html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting:    `Dear Dr. ${toName},`,
    bodyHtml:    bodyHtml,
    buttonUrl:   reviewMenuUrl,
    buttonLabel: 'Submit Review / 査読結果を提出する',
    footerHtml:  settings.mailFooter || ''
  });

  sendEmailSafe({ to: toEmail, subject, htmlBody: html },
    'Reviewer Submission Reminder L' + level + ': ' + msVer + ' to ' + toName);
}
