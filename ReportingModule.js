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
 * トリガー設定用ヘルパー
 * GASエディタからこの関数を一度実行すると、必要なトリガーが自動設定されます。
 */
function setupReportingTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 既存のリポート関連トリガーを削除して重複防止
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'sendWeeklyActivityReport' || t.getHandlerFunction() === 'archiveMonthlyLogs') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 1. 週間レポート: 毎週月曜日 午前9時
  ScriptApp.newTrigger('sendWeeklyActivityReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();

  // 2. 月次アーカイブ: 毎月1日 午前1時
  ScriptApp.newTrigger('archiveMonthlyLogs')
    .timeBased()
    .onMonthDay(1)
    .atHour(1)
    .create();
    
  Logger.log('Reporting triggers have been set up successfully.');
}
