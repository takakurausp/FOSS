/**
 * 新しいPDFレイアウトをテストするための関数です。
 * GASエディタでこの関数を実行すると、マイドライブのルートに 「Test_Receipt_Combined.pdf」 が作成されます。
 */
function testNewPdfLayout() {
  const dummyMs = {
    MsVer: 'TEST-2024-001',
    authorName: '山田 太郎',
    authorEmail: 'yamada@example.com',
    authorAffiliation: '〇〇大学 大学院',
    authorsJp: '山田太郎、田中次郎',
    authorsEn: 'Taro Yamada, Jiro Tanaka',
    paperType: 'Original Paper (原著論文)',
    titleJp: '深層学習を用いた次世代受領票生成システムの提案',
    titleEn: 'A Proposal for Next-Generation Receipt Generation System Using Deep Learning',
    runningTitle: 'AI受領票生成',
    submittedFiles: 'Main.docx, Figure1.png, Table1.xlsx',
    sendDateTime: '2024/03/30 15:00',
    ccEmails: 'jiro@example.com',
    reprintRequest: 'Yes (50 copies)',
    englishEditing: 'No'
  };

  const dummySettings = {
    Journal_Name: 'Advanced Agentic Journal',
    mailFooter: 'Contact: support@example.com | https://example.com'
  };

  // generateReceiptPdf を呼び出し
  const pdfBlob = generateReceiptPdf(dummyMs, dummySettings);
  
  if (pdfBlob) {
    pdfBlob.setName('Test_Receipt_Combined.pdf');
    const file = DriveApp.createFile(pdfBlob);
    Logger.log('PDF generated successfully! URL: ' + file.getUrl());
  } else {
    Logger.log('Failed to generate PDF.');
  }
}
