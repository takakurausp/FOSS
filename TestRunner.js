/**
 * TestRunner.js — 13 バグ修正の検証用ユニット/統合テスト
 *
 * 【使用前の必須設定】
 *   このファイル冒頭の TEST_EMAIL を、自分が受信できるテスト用メールアドレスに変更すること。
 *   そのまま実行すると本番運用者・候補編集者の実アドレスに対して動作するリスクは無いが、
 *   実メール送信を確認したいテスト（XSS / deadline 表示）でメールが届かない。
 *
 * 【関数一覧】
 *   runAllUnitTests()         : 純関数のユニットテスト（数秒、毎回実行）
 *   runAllIntegrationTests()  : シート I/O + メール送信を含む統合テスト（1〜2分）
 *   runOneScenario(name)      : 個別シナリオを 1 つだけ実行
 *
 *   testScenario_xxx()        : 個別テストシナリオ。GAS エディタから直接 Run も可能
 *
 * 【設計方針】
 *   - 相対日付（_daysAgo / _daysFromNow）でフィクスチャ生成 → 実行日に依存しない
 *   - 各シナリオはクリーンスタート（_cleanTestData）→ シード → 実行 → 検証 の流れ
 *   - 検証は Logger に OK/FAIL を出力。実行ログを目視で確認する
 *   - bootstrap() で初期化済みの **テスト環境** での実行が前提
 */

// ============================================================
// 設定
// ============================================================

/** テスト実行で使用するメールアドレス（実在する受信箱を指定）。 */
const TEST_EMAIL = 'CHANGE_ME@example.com';

// ============================================================
// 安全装置 — 本番環境誤実行防止（多層）
// ============================================================

/**
 * テスト関数実行の事前チェック。3 つの条件をすべて満たさないと実行を拒否する:
 *
 *   1. Spreadsheet 名に "test" を含む（大文字小文字無視）
 *   2. Script Property の IS_TEST_ENV が 'true'
 *   3. TEST_EMAIL が既定値 'CHANGE_ME@example.com' から変更されている
 *
 * このうち 1 つでも欠けるとテストは実行されない。これにより:
 *   - 本番環境に誤って TestRunner.js が push されても、SS 名が違うので拒否
 *   - bootstrap()（本番初期化）では IS_TEST_ENV を立てないので拒否
 *   - TEST_EMAIL を実在アドレスに変えていない素の状態でも拒否
 *
 * すべての fixture 投入・データクリア・テスト実行関数は最初にこの関数を呼ぶ。
 */
function _assertTestEnvironment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssName = ss ? ss.getName() : '';

  // Layer 1: SS 名チェック
  if (!/test/i.test(ssName)) {
    throw new Error(
      'SAFETY ABORT: Spreadsheet name does not contain "test" (got: "' + ssName + '"). '
      + 'Tests are restricted to spreadsheets whose name contains "test" (case-insensitive). '
      + 'Use scripts/deploy.ps1 -Name test to create a properly-named test environment.');
  }

  // Layer 2: Script Property フラグ
  const isTestEnv = PropertiesService.getScriptProperties().getProperty('IS_TEST_ENV');
  if (isTestEnv !== 'true') {
    throw new Error(
      'SAFETY ABORT: Script Property IS_TEST_ENV is not set to "true" (got: '
      + JSON.stringify(isTestEnv) + '). '
      + 'Run bootstrapTestEnv() first to mark this as a test environment.');
  }

  // Layer 3: TEST_EMAIL 設定済チェック
  if (TEST_EMAIL === 'CHANGE_ME@example.com') {
    throw new Error(
      'SAFETY ABORT: TEST_EMAIL is still the default placeholder. '
      + 'Edit TestRunner.js and set TEST_EMAIL to your real test inbox address.');
  }
}

// ============================================================
// 共通ヘルパー
// ============================================================

function _daysAgo(n) {
  const d = new Date();
  d.setDate(d.getDate() - n);
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
}

function _daysFromNow(n) {
  const d = new Date();
  d.setDate(d.getDate() + n);
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
}

function _logTest(name, passed, details) {
  const tag = passed ? 'PASS' : 'FAIL';
  Logger.log('[' + tag + '] ' + name + (details ? ' — ' + details : ''));
  return passed;
}

function _logSection(title) {
  Logger.log('');
  Logger.log('==== ' + title + ' ====');
}

/** Editor_log / Review_log / Manuscripts / Log / Emails のデータ行をクリア（ヘッダ保持）。 */
function _cleanTestData() {
  _assertTestEnvironment();   // 本番誤実行を物理的に拒否
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  ['Editor_log', 'Review_log', 'Manuscripts', 'Log', 'Emails'].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s && s.getLastRow() > 1) {
      s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent();
    }
  });
  SpreadsheetApp.flush();
}

/** Manuscripts シートに 1 行追加。最小限の必須列のみ受け付ける。 */
function _seedManuscript(props) {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const sheet = ss.getSheetByName('Manuscripts');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => props[h] !== undefined ? props[h] : '');
  sheet.appendRow(row);
}

/** Editor_log に 1 行追加。 */
function _seedEditorLog(props) {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const sheet = ss.getSheetByName('Editor_log');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => props[h] !== undefined ? props[h] : '');
  sheet.appendRow(row);
}

/** Review_log に 1 行追加。 */
function _seedReviewLog(props) {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const sheet = ss.getSheetByName('Review_log');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => props[h] !== undefined ? props[h] : '');
  sheet.appendRow(row);
}

/** 指定 keyValue の行で Reminder1_At/2_At/3_At が埋まっているかを返す配列 [bool, bool, bool]。 */
function _getReminderState(sheetName, keyCol, keyValue) {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const data = ss.getSheetByName(sheetName).getDataRange().getValues();
  const headers = data[0].map(h => String(h).toLowerCase().trim());
  const keyIdx = headers.indexOf(keyCol.toLowerCase());
  const r1 = headers.indexOf('reminder1_at');
  const r2 = headers.indexOf('reminder2_at');
  const r3 = headers.indexOf('reminder3_at');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][keyIdx]).trim() === keyValue) {
      return [
        r1 !== -1 && Boolean(data[i][r1]),
        r2 !== -1 && Boolean(data[i][r2]),
        r3 !== -1 && Boolean(data[i][r3])
      ];
    }
  }
  return [false, false, false];
}

/** Log シートの本文に substring を含む行が 1 つでもあれば true。 */
function _logSheetContains(substring) {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const data = ss.getSheetByName('Log').getDataRange().getValues();
  return data.slice(1).some(row => String(row[1] || '').includes(substring));
}

// ============================================================
// ユニットテスト
// ============================================================

/** Bug 2: calcReminderLevel が「未送信のうち最低レベル」から発火することを確認。 */
function _testCalcReminderLevel() {
  _logSection('Unit: calcReminderLevel (Bug 2)');
  const days = { first: 7, second: 14, third: 21 };
  const T = (e, r1, r2, r3, expected, label) => {
    const got = calcReminderLevel(e, days, r1, r2, r3);
    return _logTest('calcReminderLevel ' + label,
      got === expected, 'expected=' + expected + ' got=' + got);
  };

  let pass = 0, fail = 0;
  const cases = [
    [0,  '', '', '', 0, 'before any threshold'],
    [7,  '', '', '', 1, 'L1 fires at 7 days'],
    [7,  'X','', '', 0, 'L1 already sent, no retry'],
    [14, 'X','', '', 2, 'L2 fires after L1 sent'],
    [21, 'X','X','', 3, 'L3 fires after L2 sent'],
    [21, 'X','X','X', 0, 'all sent, no fire'],
    // Bug 2 修正の本丸: トリガが数日停止して全レベル超過になっても順序保つ
    [30, '', '', '', 1, '30d elapsed but L1 first (Bug 2 fix)'],
    [30, 'X','', '', 2, 'next day L2'],
    [30, 'X','X','', 3, 'next day L3'],
    // 期限前リマインダ（Bug 4 + Bug 13）
    [-3, '', '', '', 0, 'pre-deadline with 7/14/21 (no fire)'],
  ];
  cases.forEach(c => { c[5] === 0 || true; (T.apply(null, c) ? pass++ : fail++); });

  // 期限前ロジック（負しきい値）の検証
  const subDays = { first: -3, second: 0, third: 7 };
  const T2 = (e, r1, r2, r3, expected, label) => {
    const got = calcReminderLevel(e, subDays, r1, r2, r3);
    return _logTest('calcReminderLevel(subDays) ' + label,
      got === expected, 'expected=' + expected + ' got=' + got);
  };
  [
    [-7, '', '', '', 0, 'too early (-7 < -3)'],
    [-3, '', '', '', 1, 'pre-deadline L1 fires'],
    [0,  'X','', '', 2, 'on deadline → L2'],
    [7,  'X','X','', 3, 'overdue 7 → L3'],
  ].forEach(c => { (T2.apply(null, c) ? pass++ : fail++); });

  Logger.log('  → ' + pass + ' pass, ' + fail + ' fail');
  return fail === 0;
}

/** Bug 7: parseReminderThresholds の検証ロジック。 */
function _testParseReminderThresholds() {
  _logSection('Unit: parseReminderThresholds (Bug 7)');
  const D = { first: 7, second: 14, third: 21 };
  let pass = 0, fail = 0;
  const C = (raw, expectedVals, expectedWarns, label) => {
    const r = parseReminderThresholds('test', raw, D);
    const valOk = JSON.stringify(r.values) === JSON.stringify(expectedVals);
    const warnOk = r.warnings.length === expectedWarns;
    const ok = valOk && warnOk;
    _logTest('parseReminderThresholds ' + label, ok,
      'values=' + JSON.stringify(r.values) + ' warns=' + r.warnings.length);
    return ok;
  };

  [
    [{first:'7', second:'14', third:'21'}, {first:7,second:14,third:21}, 0, 'normal'],
    [{first:'',  second:'',   third:''  }, D,                            0, 'all empty → defaults silently'],
    [{first:'abc',second:'14',third:'21'},{first:7,second:14,third:21},  1, 'NaN → warn + default'],
    [{first:'14',second:'7', third:'21'},{first:7,second:14,third:21},   1, 'reversed → sort + warn'],
    [{first:'-3',second:'0', third:'7' },{first:-3,second:0,third:7},    0, 'negatives accepted'],
  ].forEach(c => { (C.apply(null, c) ? pass++ : fail++); });

  Logger.log('  → ' + pass + ' pass, ' + fail + ' fail');
  return fail === 0;
}

/** Bug 3: isManuscriptStillActive の状態判定。 */
function _testIsManuscriptStillActive() {
  _logSection('Unit: isManuscriptStillActive (Bug 3)');
  let pass = 0, fail = 0;
  const T = (msEntry, maxVerByMsId, expected, label) => {
    const msState = {
      msByMsVer: { 'TEST-1': msEntry },
      maxVerByMsId: maxVerByMsId
    };
    const got = isManuscriptStillActive('TEST-1', msState);
    return _logTest('isManuscriptStillActive ' + label,
      got === expected, 'expected=' + expected + ' got=' + got)
      ? pass++ : fail++;
  };

  const base = { msId: 'TEST', verNo: 1, accepted:'', stoppedByEic:'', finalStatus:'', sentBackAt:'' };
  T(base,                                            { TEST: 1 }, true,  'normal active');
  T(Object.assign({}, base, {accepted:'yes'}),       { TEST: 1 }, false, 'accepted=yes → skip');
  T(Object.assign({}, base, {stoppedByEic:'2026/01/01'}), { TEST: 1 }, false, 'stoppedByEic → skip');
  T(Object.assign({}, base, {finalStatus:'final_review'}),{ TEST: 1 }, false, 'final_review → skip');
  T(Object.assign({}, base, {finalStatus:'in_production'}),{ TEST: 1 }, false, 'in_production → skip');
  T(Object.assign({}, base, {sentBackAt:'2026/01/01'}),  { TEST: 1 }, false, 'sentBackAt → skip');
  T(base,                                            { TEST: 2 }, false, 'newer version exists → skip');
  T(undefined,                                       { TEST: 1 }, false, 'no entry → skip');

  Logger.log('  → ' + pass + ' pass, ' + fail + ' fail');
  return fail === 0;
}

/** ユニットテスト全体ランナー。 */
function runAllUnitTests() {
  Logger.log('############################');
  Logger.log('# Unit Tests');
  Logger.log('############################');
  const r1 = _testCalcReminderLevel();
  const r2 = _testParseReminderThresholds();
  const r3 = _testIsManuscriptStillActive();
  Logger.log('');
  Logger.log('==== UNIT SUMMARY ====');
  Logger.log('calcReminderLevel:        ' + (r1 ? 'PASS' : 'FAIL'));
  Logger.log('parseReminderThresholds:  ' + (r2 ? 'PASS' : 'FAIL'));
  Logger.log('isManuscriptStillActive:  ' + (r3 ? 'PASS' : 'FAIL'));
  return r1 && r2 && r3;
}

// ============================================================
// 統合テスト（フィクスチャ駆動）
// ============================================================

/** Bug 1, 2 の通常フロー: 7 日経過の編集者候補に L1 が発火する。 */
function testScenario_normalEditorL1() {
  _logSection('Integration: normalEditorL1 (Bug 1, 2)');
  _cleanTestData();
  _seedManuscript({ MS_ID: 'TST001', MsVer: 'TST001-1', Ver_No: 1, accepted: '', stoppedByEicAt: '' });
  _seedEditorLog({
    MsVer: 'TST001-1', editorKey: 'edkey001',
    Editor_Name: 'Test Editor', Editor_Email: TEST_EMAIL,
    Ask_At: _daysAgo(7), edtOk: ''
  });

  checkReminders();

  const [r1, r2, r3] = _getReminderState('Editor_log', 'editorKey', 'edkey001');
  const ok = r1 && !r2 && !r3;
  return _logTest('Editor L1 sent only', ok,
    'r1=' + r1 + ' r2=' + r2 + ' r3=' + r3);
}

/** Bug 5: 同一原稿で他編集者が承諾済の場合、未回答候補にリマインダしない。 */
function testScenario_editorAlreadyAssigned() {
  _logSection('Integration: editorAlreadyAssigned (Bug 5)');
  _cleanTestData();
  _seedManuscript({ MS_ID: 'TST002', MsVer: 'TST002-1', Ver_No: 1 });
  // 編集者 A: 承諾済
  _seedEditorLog({
    MsVer: 'TST002-1', editorKey: 'edkeyA',
    Editor_Name: 'Editor A', Editor_Email: TEST_EMAIL,
    Ask_At: _daysAgo(10), edtOk: 'ok', Answer_At: _daysAgo(8)
  });
  // 編集者 B: 未回答（10日経過）
  _seedEditorLog({
    MsVer: 'TST002-1', editorKey: 'edkeyB',
    Editor_Name: 'Editor B', Editor_Email: TEST_EMAIL,
    Ask_At: _daysAgo(10), edtOk: ''
  });

  checkReminders();

  const [r1] = _getReminderState('Editor_log', 'editorKey', 'edkeyB');
  return _logTest('Editor B not reminded (already assigned)', !r1,
    'B Reminder1_At=' + r1);
}

/** Bug 3: manuscript が accepted=yes のとき未回答編集者にリマインダしない。 */
function testScenario_manuscriptAccepted() {
  _logSection('Integration: manuscriptAccepted (Bug 3)');
  _cleanTestData();
  _seedManuscript({
    MS_ID: 'TST003', MsVer: 'TST003-1', Ver_No: 1,
    accepted: 'yes', finalStatus: 'in_production'
  });
  _seedEditorLog({
    MsVer: 'TST003-1', editorKey: 'edkey003',
    Editor_Name: 'Editor', Editor_Email: TEST_EMAIL,
    Ask_At: _daysAgo(10), edtOk: ''
  });

  checkReminders();

  const [r1] = _getReminderState('Editor_log', 'editorKey', 'edkey003');
  return _logTest('Reminder skipped (manuscript accepted)', !r1, 'r1=' + r1);
}

/** Bug 3: 同 MS_ID で新版がある場合、旧版にはリマインダしない。 */
function testScenario_newerVersionExists() {
  _logSection('Integration: newerVersionExists (Bug 3)');
  _cleanTestData();
  _seedManuscript({ MS_ID: 'TST004', MsVer: 'TST004-1', Ver_No: 1 });
  _seedManuscript({ MS_ID: 'TST004', MsVer: 'TST004-2', Ver_No: 2 });
  // 旧版 v1 への未回答行
  _seedEditorLog({
    MsVer: 'TST004-1', editorKey: 'edkey004v1',
    Editor_Name: 'Editor', Editor_Email: TEST_EMAIL,
    Ask_At: _daysAgo(10), edtOk: ''
  });

  checkReminders();

  const [r1] = _getReminderState('Editor_log', 'editorKey', 'edkey004v1');
  return _logTest('No reminder for old version', !r1, 'r1=' + r1);
}

/** Bug 4 + 11 + 13: deadline 当日に L1（subDays.first=0）が発火し、本文が "due today"。 */
function testScenario_reviewSubmissionAtDeadline() {
  _logSection('Integration: reviewSubmissionAtDeadline (Bug 4, 11, 13)');
  _cleanTestData();
  _seedManuscript({ MS_ID: 'TST005', MsVer: 'TST005-1', Ver_No: 1 });
  _seedReviewLog({
    MsVer: 'TST005-1', reviewKey: 'rvkey005',
    Rev_Name: 'Test Reviewer', Rev_Email: TEST_EMAIL,
    Ask_At: _daysAgo(20), Answer_At: _daysAgo(15),
    revOk: 'ok',
    Review_Deadline: _daysFromNow(0)  // 期限は今日
  });

  // submission の subDays デフォルト: 0/7/14 → elapsed=0 で L1 発火
  checkReminders();

  // SubReminder1_At が埋まっているはず
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const data = ss.getSheetByName('Review_log').getDataRange().getValues();
  const headers = data[0].map(h => String(h).toLowerCase().trim());
  const keyIdx  = headers.indexOf('reviewkey');
  const subR1Idx = headers.indexOf('subreminder1_at');
  let subR1 = null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][keyIdx]).trim() === 'rvkey005') {
      subR1 = subR1Idx !== -1 ? data[i][subR1Idx] : null;
      break;
    }
  }
  return _logTest('Submission L1 (SubReminder1_At) sent at deadline',
    Boolean(subR1), 'SubReminder1_At=' + subR1);
}

/** Bug 9: Editor_log から必須列を削除した状態で実行 → Log シートに警告が記録される。 */
function testScenario_missingColumn() {
  _logSection('Integration: missingColumn (Bug 9)');
  _cleanTestData();
  // 警告は「Ask_At が無い」状態を意図的に作る。
  // 1 行目を読み Ask_At 列を一時退避し、その内容を消して再書き込みする。
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const sheet = ss.getSheetByName('Editor_log');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const askAtCol = headers.findIndex(h => String(h).toLowerCase().trim() === 'ask_at');
  if (askAtCol === -1) {
    Logger.log('Skipping: Editor_log has no Ask_At column to test against');
    return false;
  }
  const original = sheet.getRange(1, askAtCol + 1).getValue();
  sheet.getRange(1, askAtCol + 1).setValue('REMOVED_FOR_TEST');
  SpreadsheetApp.flush();

  try {
    checkReminders();
    const ok = _logSheetContains('required column(s) missing');
    _logTest('Log sheet recorded missing-column warning', ok);
    return ok;
  } finally {
    // 元のヘッダ名を復元
    sheet.getRange(1, askAtCol + 1).setValue(original);
    SpreadsheetApp.flush();
  }
}

/** Bug 10: HTML エスケープ — Rev_Name に HTML/JS を入れても本文が破綻しない。 */
function testScenario_xssInName() {
  _logSection('Integration: xssInName (Bug 10)');
  _cleanTestData();
  _seedManuscript({ MS_ID: 'TST007', MsVer: 'TST007-1', Ver_No: 1 });
  _seedReviewLog({
    MsVer: 'TST007-1', reviewKey: 'rvkey007',
    Rev_Name: "<script>alert(1)</script>O'Brien & Co",
    Rev_Email: TEST_EMAIL,
    Ask_At: _daysAgo(7), revOk: ''
  });

  checkReminders();

  const [r1] = _getReminderState('Review_log', 'reviewKey', 'rvkey007');
  // 受信箱を確認: メール本文ソースを開いて &lt;script&gt; のようにエスケープされていれば OK
  Logger.log('  Manual check: open the email at ' + TEST_EMAIL
    + ' and view source. Look for &lt;script&gt; (escaped) NOT <script> (raw).');
  return _logTest('Reminder sent (manual content check needed)', r1);
}

// ============================================================
// 統合テスト全体ランナー
// ============================================================

const _ALL_SCENARIOS = [
  testScenario_normalEditorL1,
  testScenario_editorAlreadyAssigned,
  testScenario_manuscriptAccepted,
  testScenario_newerVersionExists,
  testScenario_reviewSubmissionAtDeadline,
  testScenario_missingColumn,
  testScenario_xssInName
];

function runAllIntegrationTests() {
  _assertTestEnvironment();   // 3 層の安全装置を入口で実行

  Logger.log('############################');
  Logger.log('# Integration Tests');
  Logger.log('############################');
  const results = [];
  _ALL_SCENARIOS.forEach(fn => {
    try {
      const ok = fn();
      results.push({ name: fn.name, status: ok ? 'PASS' : 'FAIL' });
    } catch (e) {
      results.push({ name: fn.name, status: 'ERROR: ' + e.message });
      Logger.log('[ERROR] ' + fn.name + ': ' + e.message);
      Logger.log(e.stack || '');
    }
  });

  Logger.log('');
  Logger.log('==== INTEGRATION SUMMARY ====');
  results.forEach(r => Logger.log('  ' + r.status.padEnd(10) + ' ' + r.name));
  const failed = results.filter(r => r.status !== 'PASS').length;
  Logger.log('Total: ' + results.length + ', Failed: ' + failed);
  return failed === 0;
}

/** 個別シナリオを名前で実行（GAS エディタから直接実行する場合用）。 */
function runOneScenario(name) {
  const fn = _ALL_SCENARIOS.find(f => f.name === name || f.name === 'testScenario_' + name);
  if (!fn) throw new Error('Scenario not found: ' + name);
  return fn();
}

// ============================================================
// 全テスト一括実行（リリース前確認用）
// ============================================================

function runAllTests() {
  const u = runAllUnitTests();
  const i = runAllIntegrationTests();
  Logger.log('');
  Logger.log('############ FINAL ############');
  Logger.log('Unit:        ' + (u ? 'PASS' : 'FAIL'));
  Logger.log('Integration: ' + (i ? 'PASS' : 'FAIL'));
  return u && i;
}
