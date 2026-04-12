// 決定状態変更時に通知を送信するスクリプト
function sendDecisionNotification(decisionRow) {
  const TEMPLATE = {
    'Accept': '採択通知テンプレート',
    'Minor Revision': '軽微修正依頼テンプレート',
    'Reject': '不採択通知テンプレート'
  };
  
  const template = TEMPLATE[decisionRow.decision] || 'デフォルトテンプレート';
  const recipient = decisionRow.authorEmail;
  
  MailApp.sendEmail({
    to: recipient,
    subject: `論文審査結果 / Manuscript Decision: ${decisionRow.decision}`,
    htmlBody: template
      .replace('{{author}}', decisionRow.authorName)
      .replace('{{title}}', decisionRow.paperTitle)
  });
}