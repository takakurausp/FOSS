/**
 * Bootstrap.js - 新規環境（または既存環境の再初期化）の一発セットアップ
 *
 * 【使い方】
 *   1. このファイルを含むプロジェクトを clasp push で対象 SS にバンドルする
 *   2. GAS エディタで bootstrap 関数を選択 → Run
 *   3. OAuth 認証ダイアログが出るので承認
 *   4. 実行ログ（Logger）に "=== bootstrap complete ===" が出れば成功
 *
 * 【冪等性】
 *   - すべてのステップが冪等。再実行しても既存データは破壊されない:
 *     - setupSheets       : 既存シートはそのまま、不足列のみ追加
 *     - setupSettings     : 既存値があるキーはスキップ、未設定のみ補充
 *     - setupDecisions    : 既存行があれば追加しない
 *     - setupManuscriptTypes : 同上
 *     - setupReportingTriggers / setupReminderTriggers
 *                          : 同名トリガを削除してから再登録（重複防止）
 *
 * 【setupAll() との関係】
 *   - setupAll() は SetupScript.js の既存関数で、シート骨格〜マスタデータの投入まで担当。
 *   - 本 bootstrap() は setupAll() を呼んだ上で、トリガ登録 2 種を追加実行する。
 *   - つまり「bootstrap = setupAll + トリガ登録」と覚えればよい。
 *
 * 【テスト環境用】
 *   - bootstrapTestEnv() は将来的にテスト fixture の自動投入を行うフック。
 *     現時点では bootstrap() と同じだが、テスト系コードを足す際の入り口として用意。
 */

/**
 * 本番・テスト共通の一発初期化関数。
 * 戻り値: 実行サマリ文字列（GAS UI alert にも表示される）。
 */
function bootstrap() {
  const messages = [];
  messages.push('=== bootstrap start ===');

  try {
    // 0. 本番初期化時は IS_TEST_ENV フラグを **明示的に削除**。
    //    過去にこの SS をテスト用途で使っていた場合の残留フラグを除去し、
    //    TestRunner.js の destructive 関数が誤って動かないことを保証する。
    PropertiesService.getScriptProperties().deleteProperty('IS_TEST_ENV');
    messages.push('[OK] IS_TEST_ENV property cleared (production-mode).');

    // 1. シート骨格 + Settings + Decisions + ManuscriptTypes + SS_ID 保存
    //    （setupAll は内部で alert 表示するが、戻り値も拾う）
    const setupAllResult = setupAll();
    messages.push(setupAllResult);

    // 2. レポート系トリガ（週次レポート、月次アーカイブ、原稿アーカイブ等）
    setupReportingTriggers();
    messages.push('[OK] setupReportingTriggers: レポート系トリガを登録しました。');

    // 3. リマインダ系トリガ（checkReminders 09:00 / retrySendingEmails 12:00）
    setupReminderTriggers();
    messages.push('[OK] setupReminderTriggers: リマインダ系トリガを登録しました。');

    // 4. 終了サマリ
    const triggers = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
    const ssUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    messages.push('');
    messages.push('=== bootstrap complete ===');
    messages.push('Active triggers: ' + JSON.stringify(triggers));
    messages.push('Spreadsheet URL: ' + ssUrl);
    messages.push('');
    messages.push('次のステップ / Next steps:');
    messages.push('  1. Settings シートで Journal_Name / chiefEditorEmail / managingEditorEmail 等を編集');
    messages.push('  2. Webアプリとしてデプロイ（公開: 全員 / 実行: 自分）');
    messages.push('  3. デプロイ URL を運用者・著者・編集者に共有');
  } catch (err) {
    messages.push('');
    messages.push('[ERROR] bootstrap failed: ' + err.message);
    messages.push(err.stack || '');
  }

  const result = messages.join('\n');
  Logger.log(result);
  try { SpreadsheetApp.getUi().alert(result); } catch (_) { /* UI 不可コンテキストでは無視 */ }
  return result;
}

/**
 * テスト環境向け初期化。bootstrap() を呼んだ後、テスト用フィクスチャ・設定を上書き。
 *
 * 現時点で行うこと:
 *   - bootstrap() 実行
 *   - Settings の値をテスト向けに調整（例: 全メール送信先を TEST_EMAIL に集約）は呼び出し側で
 *
 * 将来的には _seedTestFixture('xxx') 等の fixture 投入を追加予定。
 */
function bootstrapTestEnv() {
  Logger.log('=== bootstrapTestEnv start ===');

  // 安全装置 1/3: SS 名チェック（本番 SS で誤実行されないように）
  const ssName = SpreadsheetApp.getActiveSpreadsheet().getName();
  if (!/test/i.test(ssName)) {
    throw new Error(
      'bootstrapTestEnv refuses to run on a spreadsheet whose name does not contain "test". '
      + 'Current name: "' + ssName + '". '
      + 'Use scripts/deploy.ps1 -Name test to create a properly-named test environment.');
  }

  // 通常の初期化（bootstrap は IS_TEST_ENV を一旦削除する）
  bootstrap();

  // テスト環境であることを示す Script Property を **bootstrap 後に** セットする。
  // この順序により、本番初期化（bootstrap 単独）では IS_TEST_ENV は必ず削除されたまま、
  // テスト初期化（bootstrapTestEnv）では削除→セットの順で 'true' になる。
  PropertiesService.getScriptProperties().setProperty('IS_TEST_ENV', 'true');
  Logger.log('[OK] IS_TEST_ENV property set to "true" — TestRunner is now enabled.');

  // 将来の拡張ポイント: テスト用 Settings 上書き、fixture 自動投入 等
  Logger.log('--- test environment hook (currently no-op) ---');

  Logger.log('=== bootstrapTestEnv complete ===');
}
