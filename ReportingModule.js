/**
 * ReportingModule.js - アクティビティログの週次レポートと月次アーカイブ
 */

/**
 * 週間アクティビティレポートを生成して編集委員長に送信する
 * 形式: PDF添付 + シンプルなメール本文
 */
function sendWeeklyActivityReport() {
  const ssId = getSpreadsheetId();
  const settings = getSettings();
  if (!settings.chiefEditorEmail) return;

  const now = new Date();
  const sevenDaysAgo = new Date(now.getTime() - (7 * 24 * 60 * 60 * 1000));
  const dateRangeStr = Utilities.formatDate(sevenDaysAgo, 'Asia/Tokyo', 'yyyy/MM/dd') + ' - ' + Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd');

  // 1. 最近のアクティビティを取得 (Logシート)
  const activityLogs = getRecentLogs(ssId, sevenDaysAgo);
  
  // 2. 現在の進捗状況を取得 (Manuscripts, Editor_log, Review_log)
  const statusSummary = getProgressSummary(ssId);

  // 3. HTMLレポートを生成
  const html = generateReportHtml(settings.Journal_Name, dateRangeStr, activityLogs, statusSummary);
  
  // 4. PDFに変換
  const pdfBlob = HtmlService.createHtmlOutput(html).getBlob().getAs(MimeType.PDF).setName(`Weekly_Report_${Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd')}.pdf`);

  // 5. メール送信
  const subject = `[${settings.Journal_Name}] 週次活動レポート / Weekly Activity Report (${dateRangeStr})`;
  const bodyHtml = `
    <p>Please find the weekly activity report for <strong>${settings.Journal_Name}</strong> attached to this email.</p>
    <p>This report includes a summary of activities from the last 7 days and the current status of all pending manuscripts.</p>
    <hr style="border:none; border-top:1px solid #e2e8f0; margin: 20px 0;">
    <p>今週の活動レポートをお送りいたします。添付のPDFにて詳細をご確認ください。</p>
  `;

  const emailHtml = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting: 'Dear Editor-in-Chief,',
    bodyHtml: bodyHtml,
    footerHtml: settings.mailFooter || ''
  });

  sendEmailSafe({
    to: settings.chiefEditorEmail,
    subject: subject,
    htmlBody: emailHtml,
    attachments: [pdfBlob]
  }, 'Weekly Activity Report');
}

/**
 * 月次ログアーカイブを実行する
 * 1. 前月分のログを抽出
 * 2. CSV/ZIP化してメール送信
 * 3. ログを Archive シートへ移動
 */
function archiveMonthlyLogs() {
  const ssId = getSpreadsheetId();
  const settings = getSettings();
  if (!settings.chiefEditorEmail) return;

  const now = new Date();
  // 前月の初日と最終日を特定
  const firstDayPrevMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const lastDayPrevMonth = new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59);
  
  const monthStr = Utilities.formatDate(firstDayPrevMonth, 'Asia/Tokyo', 'yyyy-MM');

  // 1. ログシートから対象データを抽出
  const logSheet = SpreadsheetApp.openById(ssId).getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) return;
  const data = logSheet.getDataRange().getValues();
  const headers = data[0];
  
  const rowsToArchive = [];
  const rowIndicesToRemove = [];

  for (let i = 1; i < data.length; i++) {
    const logDate = new Date(data[i][0]);
    if (logDate >= firstDayPrevMonth && logDate <= lastDayPrevMonth) {
      rowsToArchive.push(data[i]);
      rowIndicesToRemove.push(i + 1);
    }
  }

  if (rowsToArchive.length === 0) {
    Logger.log('No logs found for archiving: ' + monthStr);
    return;
  }

  // 2. CSV文字列の作成
  const csvContent = [headers, ...rowsToArchive].map(row => 
    row.map(val => {
      let str = String(val).replace(/"/g, '""');
      return str.includes(',') || str.includes('\n') || str.includes('"') ? `"${str}"` : str;
    }).join(',')
  ).join('\n');

  // 3. ZIP圧縮
  const csvBlob = Utilities.newBlob(csvContent, 'text/csv', `logs_${monthStr}.csv`);
  const zipBlob = Utilities.zip([csvBlob], `Activity_Logs_${monthStr}.zip`);

  // 4. メール送信
  const subject = `[${settings.Journal_Name}] 月次ログアーカイブ / Monthly Log Archive (${monthStr})`;
  const bodyHtml = `
    <p>Attached is the archived activity log for <strong>${monthStr}</strong>.</p>
    <p>The original entries have been moved from the <strong>Log</strong> sheet to the <strong>Log_archive</strong> sheet.</p>
    <hr style="border:none; border-top:1px solid #e2e8f0; margin: 20px 0;">
    <p>${monthStr} 分の活動ログをアーカイブし、ZIP形式で送付いたします。対象のデータは Log シートから Log_archive シートへ移動済みです。</p>
  `;

  const emailHtml = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting: 'Dear Editor-in-Chief,',
    bodyHtml: bodyHtml,
    footerHtml: settings.mailFooter || ''
  });

  sendEmailSafe({
    to: settings.chiefEditorEmail,
    subject: subject,
    htmlBody: emailHtml,
    attachments: [zipBlob]
  }, 'Monthly Log Archive');

  // 5. Log_archive シートへ移動（原稿アーカイブ用の Archive シートとは別管理）
  let archiveSheet = SpreadsheetApp.openById(ssId).getSheetByName(LOG_ARCHIVE_SHEET_NAME);
  if (!archiveSheet) {
    archiveSheet = SpreadsheetApp.openById(ssId).insertSheet(LOG_ARCHIVE_SHEET_NAME);
    archiveSheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#f1f5f9');
    archiveSheet.setFrozenRows(1);
  }

  archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsToArchive.length, headers.length).setValues(rowsToArchive);

  // 6. Logシートから削除（後ろから削除してインデックスずれ防止）
  rowIndicesToRemove.reverse().forEach(idx => logSheet.deleteRow(idx));
}

/**
 * 最近のログを取得
 */
function getRecentLogs(ssId, sinceDate) {
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(LOG_SHEET_NAME);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  return data.slice(1).filter(row => new Date(row[0]) >= sinceDate).map(row => ({
    date: Utilities.formatDate(new Date(row[0]), 'Asia/Tokyo', 'MM/dd HH:mm'),
    text: row[1]
  }));
}

/**
 * 現在の進捗状況をサマリーとして取得
 */
function getProgressSummary(ssId) {
  const msData = SpreadsheetApp.openById(ssId).getSheetByName(MANUSCRIPTS_SHEET_NAME).getDataRange().getValues();
  const edLog = SpreadsheetApp.openById(ssId).getSheetByName(EDITOR_LOG_SHEET_NAME).getDataRange().getValues();
  const revLog = SpreadsheetApp.openById(ssId).getSheetByName(REVIEW_LOG_SHEET_NAME).getDataRange().getValues();

  // Manuscripts シートのヘッダーインデックス（ループ外で一度だけ解決する）
  const msHeaders     = msData[0];
  const acceptedIdx   = msHeaders.indexOf('accepted');
  const scoreIdx      = msHeaders.indexOf('score');
  const stoppedIdx    = msHeaders.indexOf('stoppedByEicAt');
  const msVerIdx      = msHeaders.indexOf('MsVer');
  const sentBackAtIdx = msHeaders.indexOf('sentBackAt');

  // EditorLog シートのヘッダーインデックス（列順に依存しないよう名前で解決する）
  const edHeaders  = edLog[0];
  const edMsVerIdx = edHeaders.indexOf('MsVer');
  const edOkIdx    = edHeaders.indexOf('edtOk');
  const edLogRows  = edLog.slice(1); // ヘッダー行を除いたデータ行

  const pendingMss = msData.slice(1).filter(row =>
    (row[acceptedIdx] === '' || row[acceptedIdx] === null) &&
    !(stoppedIdx !== -1 && String(row[stoppedIdx] || '').trim())
  );

  return {
    totalPending: pendingMss.length,
    waitingEditor: pendingMss.filter(row => {
      const msVer = row[msVerIdx];
      const assigned = edLogRows.filter(e => e[edMsVerIdx] === msVer && e[edOkIdx] === 'ok');
      return assigned.length === 0;
    }).length,
    // score が空の原稿数。waitingEditor（担当未定）も内包するため「判定前合計」と表示する
    pendingDecision: pendingMss.filter(row => row[scoreIdx] === '').length,
    decidedThisWeek: msData.slice(1).filter(row => {
      const date = new Date(row[sentBackAtIdx]);
      return !isNaN(date.getTime()) && (new Date().getTime() - date.getTime()) < 7 * 24 * 60 * 60 * 1000;
    }).length
  };
}

/**
 * レポートHTMLの生成
 */
function generateReportHtml(journalName, dateRange, logs, summary) {
  const esc = s => String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  
  return `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  body { font-family: 'Helvetica Neue', Arial, sans-serif; color: #1e293b; line-height: 1.5; padding: 40px; }
  .header { border-bottom: 2px solid #2563eb; padding-bottom: 15px; margin-bottom: 30px; }
  .title { font-size: 24px; font-weight: bold; color: #2563eb; margin: 0; }
  .subtitle { font-size: 14px; color: #64748b; margin-top: 5px; }
  .section { margin-bottom: 40px; }
  .section-title { font-size: 18px; font-weight: 600; margin-bottom: 15px; padding-left: 10px; border-left: 4px solid #2563eb; }
  .stats-grid { display: table; width: 100%; border-spacing: 15px; margin: -15px; }
  .stat-card { display: table-cell; background: #f8fafc; padding: 20px; border-radius: 8px; border: 1px solid #e2e8f0; text-align: center; width: 25%; }
  .stat-value { font-size: 28px; font-weight: bold; color: #2563eb; }
  .stat-label { font-size: 12px; color: #64748b; text-transform: uppercase; letter-spacing: 0.5px; }
  table { width: 100%; border-collapse: collapse; margin-top: 10px; }
  th { background: #f1f5f9; text-align: left; padding: 12px; font-size: 13px; border-bottom: 1px solid #e2e8f0; }
  td { padding: 12px; font-size: 13px; border-bottom: 1px solid #f1f5f9; }
</style>
</head>
<body>
  <div class="header">
    <h1 class="title">${esc(journalName)}</h1>
    <p class="subtitle">Weekly Activity Report: ${dateRange}</p>
  </div>

  <div class="section">
    <h2 class="section-title">Current Status Overview / 進捗概計</h2>
    <div class="stats-grid">
      <div class="stat-card">
        <div class="stat-value">${summary.totalPending}</div>
        <div class="stat-label">Total Pending / 未判定合計</div>
      </div>
      <div class="stat-card">
        <div class="stat-value">${summary.waitingEditor}</div>
        <div class="stat-label">Waiting Editor / 担当未定</div>
      </div>
      <div class="stat-card">
        <div class="stat-value">${summary.pendingDecision}</div>
        <div class="stat-label">判定前合計 / Pending Decision</div>
      </div>
      <div class="stat-card">
        <div class="stat-value">${summary.decidedThisWeek}</div>
        <div class="stat-label">Decided (Week) / 今週の判定</div>
      </div>
    </div>
  </div>

  <div class="section">
    <h2 class="section-title">Recent Activity Logs / 最近の活動履歴</h2>
    <table>
      <thead>
        <tr>
          <th style="width: 120px;">Timestamp</th>
          <th>Action</th>
        </tr>
      </thead>
      <tbody>
        ${logs.length > 0 ? [...logs].reverse().map(l => `
          <tr>
            <td style="color: #64748b;">${l.date}</td>
            <td>${esc(l.text)}</td>
          </tr>
        `).join('') : '<tr><td colspan="2" style="text-align: center; color: #94a3b8; padding: 30px;">No activity recorded this week.</td></tr>'}
      </tbody>
    </table>
  </div>

  <footer style="margin-top: 50px; font-size: 11px; color: #94a3b8; text-align: center; border-top: 1px solid #f1f5f9; padding-top: 20px;">
    &copy; ${new Date().getFullYear()} ${esc(journalName)} - Automated Report
  </footer>
</body>
</html>`;
}

/**
 * 定期的な原稿アーカイブ処理
 * 判定日から ARCHIVE_AGE_MONTHS ヶ月以上経過した以下の原稿を各アーカイブシートへ移動：
 *   - 受理確定原稿 (accepted='yes')        → Accepted_archive
 *   - EIC早期却下原稿 (stoppedByEicAt あり) → Rejected_archive
 * Manuscripts シートから該当行を削除し、Drive フォルダ名に [ARCHIVED] を付加する。
 * Editor_log / Review_log は変更しない（本体の肥大化対策を優先）。
 */
function archiveAgedManuscripts() {
  const ssId = getSpreadsheetId();
  if (!ssId) return;
  const ss = SpreadsheetApp.openById(ssId);
  const msSheet = ss.getSheetByName(MANUSCRIPTS_SHEET_NAME);
  if (!msSheet) return;

  const now = new Date();
  const threshold = new Date(now.getFullYear(), now.getMonth() - ARCHIVE_AGE_MONTHS, now.getDate());

  const data = msSheet.getDataRange().getValues();
  if (data.length < 2) return;

  const headers = data[0];
  const acceptedIdx   = headers.indexOf('accepted');
  const stoppedAtIdx  = headers.indexOf('stoppedByEicAt');
  const sentBackAtIdx = headers.indexOf('sentBackAt');
  const meSentAtIdx   = headers.indexOf('managingEditorSentAt');
  const finalStatusIdx= headers.indexOf('finalStatus');
  const msVerIdx      = headers.indexOf('MsVer');
  const msIdIdx       = headers.indexOf('MS_ID');
  const folderUrlIdx  = headers.indexOf('folderUrl');
  const caNameIdx     = headers.indexOf('CA_Name');

  const acceptedRows = [];
  const rejectedRows = [];
  const rowsToDelete = []; // 1-indexed

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // EIC早期却下（優先判定）
    if (stoppedAtIdx !== -1) {
      const stoppedAt = row[stoppedAtIdx];
      if (stoppedAt && String(stoppedAt).trim()) {
        const d = new Date(stoppedAt);
        if (!isNaN(d.getTime()) && d < threshold) {
          rejectedRows.push(row);
          rowsToDelete.push(i + 1);
          continue;
        }
      }
    }

    // 受理確定（ME 進行中は除外: finalStatus='final_review' はまだ作業中）
    const accepted = String(row[acceptedIdx] || '').toLowerCase().trim();
    const finalStatus = finalStatusIdx !== -1 ? String(row[finalStatusIdx] || '').trim() : '';
    if (accepted === 'yes' && finalStatus !== 'final_review') {
      const ts = row[sentBackAtIdx] || (meSentAtIdx !== -1 ? row[meSentAtIdx] : '');
      if (ts) {
        const d = new Date(ts);
        if (!isNaN(d.getTime()) && d < threshold) {
          acceptedRows.push(row);
          rowsToDelete.push(i + 1);
        }
      }
    }
  }

  if (acceptedRows.length === 0 && rejectedRows.length === 0) {
    writeLog('Archive batch (manuscripts): no records older than ' + ARCHIVE_AGE_MONTHS + ' months.');
    return;
  }

  _appendRowsToManuscriptArchive(ss, ACCEPTED_ARCHIVE_SHEET_NAME, headers, acceptedRows);
  _appendRowsToManuscriptArchive(ss, REJECTED_ARCHIVE_SHEET_NAME, headers, rejectedRows);

  // Drive フォルダをリネーム（アーカイブ対象行すべて）
  [...acceptedRows, ...rejectedRows].forEach(row => {
    const folderUrl = folderUrlIdx !== -1 ? row[folderUrlIdx] : '';
    if (!folderUrl) return;
    try {
      const match = String(folderUrl).match(/[-\w]{25,}/);
      if (!match) return;
      const folder = DriveApp.getFolderById(match[0]);
      const oldName = folder.getName();
      if (!oldName.includes('[ARCHIVED]')) folder.setName('[ARCHIVED] ' + oldName);
    } catch(e) {
      Logger.log('Archive folder rename failed for MsVer=' + (msVerIdx !== -1 ? row[msVerIdx] : '?') + ': ' + e.message);
    }
  });

  // Manuscripts から削除（後ろから）
  rowsToDelete.sort((a, b) => b - a).forEach(idx => msSheet.deleteRow(idx));

  // キャッシュ無効化
  try { spreadsheetCache.invalidate(ssId, MANUSCRIPTS_SHEET_NAME); } catch(_) {}

  writeLog('Archive batch (manuscripts): ' + acceptedRows.length + ' accepted → Accepted_archive, ' + rejectedRows.length + ' EIC-rejected → Rejected_archive (≥ ' + ARCHIVE_AGE_MONTHS + ' months old).');
}

/**
 * 再投稿期限切れ原稿のアーカイブ処理
 *
 * Decisions シートで Resubmit='yes' に該当するスコア（Major/Minor Revision 等）の判定を受け、
 * Settings の Resubmission_Expire_Months を超過しても再投稿が行われなかった原稿を
 * Expired_archive シートへ移動する。
 *
 * 保護ルール：
 *   - 同 MS_ID に Ver_No が大きい行（再投稿済み）が存在する場合は対象外。
 *     → 前回ラウンドの記録は決してアーカイブされない。
 *   - sentBackAt が空／パース不可の行はスキップ。
 *   - EIC早期却下（stoppedByEicAt あり）は archiveAgedManuscripts() 側で処理済みのため除外。
 */
function archiveExpiredResubmissions() {
  const ssId = getSpreadsheetId();
  if (!ssId) return;

  const settings = getSettings();
  const expireMonths = parseInt(settings.Resubmission_Expire_Months || '6', 10) || 6;

  const ss = SpreadsheetApp.openById(ssId);
  const msSheet = ss.getSheetByName(MANUSCRIPTS_SHEET_NAME);
  if (!msSheet) return;

  const now = new Date();
  const threshold = new Date(now.getFullYear(), now.getMonth() - expireMonths, now.getDate());

  // ── Decisions シートから Resubmit=yes のスコア名を一括取得 ──────────────────
  const allowsResubmitSet = new Set();
  try {
    const decSheet = ss.getSheetByName(DECISION_MAIL_SHEET_NAME);
    if (decSheet) {
      const decData = decSheet.getDataRange().getValues();
      if (decData.length >= 2) {
        const decHeaders = decData[0].map(h => String(h).trim());
        const scoreNameIdx  = decHeaders.indexOf('ShortExplanation');
        const resubmitIdx   = decHeaders.indexOf('Resubmit');
        if (scoreNameIdx !== -1 && resubmitIdx !== -1) {
          for (let r = 1; r < decData.length; r++) {
            const v = String(decData[r][resubmitIdx] || '').trim().toLowerCase();
            if (v === 'yes' || v === 'true' || v === '1') {
              const scoreName = String(decData[r][scoreNameIdx] || '').trim();
              if (scoreName) allowsResubmitSet.add(scoreName);
            }
          }
        }
      }
    }
  } catch(e) {
    Logger.log('archiveExpiredResubmissions: Decisions sheet read failed: ' + e.message);
  }

  if (allowsResubmitSet.size === 0) {
    writeLog('archiveExpiredResubmissions: no Resubmit=yes decisions defined — skipping.');
    return;
  }

  // ── Manuscripts シート全行を取得 ──────────────────────────────────────────────
  const data = msSheet.getDataRange().getValues();
  if (data.length < 2) return;

  const headers      = data[0];
  const scoreIdx       = headers.indexOf('score');
  const sentBackAtIdx  = headers.indexOf('sentBackAt');
  const acceptedIdx    = headers.indexOf('accepted');
  const stoppedAtIdx   = headers.indexOf('stoppedByEicAt');
  const finalStatusIdx = headers.indexOf('finalStatus');
  const msIdIdx        = headers.indexOf('MS_ID');
  const verNoIdx       = headers.indexOf('Ver_No');
  const folderUrlIdx   = headers.indexOf('folderUrl');
  const msVerIdx       = headers.indexOf('MsVer');

  if (scoreIdx === -1 || sentBackAtIdx === -1 || acceptedIdx === -1 || msIdIdx === -1 || verNoIdx === -1) {
    Logger.log('archiveExpiredResubmissions: required header(s) missing — skipping.');
    return;
  }

  // ── MS_ID ごとの最大 Ver_No を先に計算（前回ラウンド保護のため）────────────
  const maxVerNoByMsId = {};
  for (let i = 1; i < data.length; i++) {
    const msId = String(data[i][msIdIdx] || '').trim();
    const verNo = Number(data[i][verNoIdx] || 0);
    if (msId) {
      maxVerNoByMsId[msId] = Math.max(maxVerNoByMsId[msId] || 0, verNo);
    }
  }

  // ── 各行を判定 ────────────────────────────────────────────────────────────────
  const expiredRows  = [];
  const rowsToDelete = []; // 1-indexed

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const score       = String(row[scoreIdx]      || '').trim();
    const sentBackAt  = row[sentBackAtIdx];
    const accepted    = String(row[acceptedIdx]   || '').trim().toLowerCase();
    const stoppedAt   = String(row[stoppedAtIdx !== -1 ? stoppedAtIdx : -1] || '').trim();
    const finalStatus = String(finalStatusIdx !== -1 ? (row[finalStatusIdx] || '') : '').trim();
    const msId        = String(row[msIdIdx]       || '').trim();
    const verNo       = Number(row[verNoIdx]      || 0);

    // (a) 改訂要求スコアであること
    if (!score || !allowsResubmitSet.has(score)) continue;

    // (b) sentBackAt が設定値より古いこと
    if (!sentBackAt || !String(sentBackAt).trim()) continue;
    const sentDate = new Date(sentBackAt);
    if (isNaN(sentDate.getTime()) || sentDate >= threshold) continue;

    // (c) 受理済みでないこと
    if (accepted === 'yes') continue;

    // (d) EIC早期却下でないこと（archiveAgedManuscripts で処理される）
    if (stoppedAt) continue;

    // (e) 最終確認フェーズ中でないこと
    if (finalStatus === 'final_review' || finalStatus === 'in_production') continue;

    // (f) 最新バージョンであること（再投稿済みの前回ラウンドを除外）
    if (!msId || verNo !== maxVerNoByMsId[msId]) continue;

    expiredRows.push(row);
    rowsToDelete.push(i + 1);
  }

  if (expiredRows.length === 0) {
    writeLog('archiveExpiredResubmissions: no expired records found (≥ ' + expireMonths + ' months).');
    return;
  }

  // ── Expired_archive へ追記 ────────────────────────────────────────────────────
  _appendRowsToManuscriptArchive(ss, EXPIRED_ARCHIVE_SHEET_NAME, headers, expiredRows);

  // ── Drive フォルダをリネーム ─────────────────────────────────────────────────
  expiredRows.forEach(row => {
    const folderUrl = folderUrlIdx !== -1 ? String(row[folderUrlIdx] || '') : '';
    if (!folderUrl) return;
    try {
      const match = folderUrl.match(/[-\w]{25,}/);
      if (!match) return;
      const folder = DriveApp.getFolderById(match[0]);
      const oldName = folder.getName();
      if (!oldName.includes('[ARCHIVED]')) folder.setName('[ARCHIVED] ' + oldName);
    } catch(e) {
      Logger.log('archiveExpiredResubmissions: folder rename failed for MsVer=' +
        (msVerIdx !== -1 ? String(row[msVerIdx] || '?') : '?') + ': ' + e.message);
    }
  });

  // ── Manuscripts から削除（後ろから順に）──────────────────────────────────────
  rowsToDelete.sort((a, b) => b - a).forEach(idx => msSheet.deleteRow(idx));

  // ── キャッシュ無効化 ──────────────────────────────────────────────────────────
  try { spreadsheetCache.invalidate(ssId, MANUSCRIPTS_SHEET_NAME); } catch(_) {}

  writeLog('archiveExpiredResubmissions: ' + expiredRows.length +
    ' manuscript(s) → Expired_archive (sentBackAt ≥ ' + expireMonths + ' months ago).');
}

/**
 * 原稿アーカイブシートへ行を追加するヘルパー。
 * ヘッダーが無ければ Manuscripts シートと同じヘッダーで初期化する。
 */
function _appendRowsToManuscriptArchive(ss, sheetName, headers, rows) {
  if (rows.length === 0) return;
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#f1f5f9');
    sheet.setFrozenRows(1);
  } else {
    // 既存シートのヘッダー列数が不足している場合は拡張
    const existingCols = sheet.getLastColumn();
    if (existingCols < headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers])
        .setFontWeight('bold').setBackground('#f1f5f9');
    }
  }
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}

/**
 * トリガー設定用ヘルパー
 * GASエディタからこの関数を一度実行すると、必要なトリガーが自動設定されます。
 */
function setupReportingTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 既存のリポート関連トリガーを削除して重複防止
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'sendWeeklyActivityReport' || fn === 'archiveMonthlyLogs' ||
        fn === 'archiveAgedManuscripts'  || fn === 'archiveExpiredResubmissions') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 1. 週間レポート: 毎週月曜日 午前9時
  ScriptApp.newTrigger('sendWeeklyActivityReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();

  // 2. 月次ログアーカイブ: 毎月1日 午前1時
  ScriptApp.newTrigger('archiveMonthlyLogs')
    .timeBased()
    .onMonthDay(1)
    .atHour(1)
    .create();

  // 3. 原稿アーカイブ（受理確定・EIC早期却下）: 毎月15日 午前2時
  ScriptApp.newTrigger('archiveAgedManuscripts')
    .timeBased()
    .onMonthDay(15)
    .atHour(2)
    .create();

  // 4. 再投稿期限切れ原稿アーカイブ: 毎月20日 午前3時
  ScriptApp.newTrigger('archiveExpiredResubmissions')
    .timeBased()
    .onMonthDay(20)
    .atHour(3)
    .create();

  Logger.log('Reporting triggers have been set up successfully.');
}
