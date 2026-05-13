/**
 * ReminderModule.js - 査読者・担当編集者へのリマインドメール
 *
 * 【トリガー設定】
 * GAS エディタからこのファイル内の setupReminderTriggers() を一度実行すれば、
 * 必要な時間ベーストリガ（毎日 9:00 の checkReminders、12:00 の retrySendingEmails）
 * が自動登録されます。手動で UI から設定する必要はありません。
 *
 * 【Settingsシートに必要な設定】
 * 招待フェーズ（編集者・査読者の依頼への未回答リマインド）:
 *   firstReminderDays  : 依頼から何日後に1通目を送るか（例: 7）
 *   secondReminderDays : 依頼から何日後に2通目を送るか（例: 14）
 *   thirdReminderDays  : 依頼から何日後に最終通知を送るか（例: 21）
 *
 * 提出フェーズ（査読承諾済みの未提出査読者へのリマインド）:
 *   submissionReminderL1Days : Review_Deadline からの超過日数（例: 0  = 期限到来時）
 *   submissionReminderL2Days :   〃                          （例: 7  = 7日超過）
 *   submissionReminderL3Days :   〃                          （例: 14 = 14日超過 = Final）
 *   ※負の値（例: -3）にすれば期限前リマインダにできる。
 *   ※Review_Deadline が空・解析不能な場合は招待フェーズと同じしきい値で
 *     Answer_At からの経過日数を見る既存挙動にフォールバックする。
 *
 * 【ログシートに必要な列】
 * Editor_log  : Reminder1_At, Reminder2_At, Reminder3_At
 * Review_log  : Reminder1_At, Reminder2_At, Reminder3_At
 *               SubReminder1_At, SubReminder2_At, SubReminder3_At
 * （Reminder*_At 列が存在しない場合はリマインド送信済みフラグが記録されないため、
 *   毎日再送されます。必ず各シートに上記列を追加してください。
 *   Review_log の SubReminder*_At 列は、存在しなければ自動追加されます。）
 */

/**
 * メインエントリ。毎日1回トリガーで呼び出す。
 */
function checkReminders() {
  const ssId = getSpreadsheetId();
  const settings = getSettings();

  // Bug 7: Settings のしきい値を厳格にバリデート（非数値・逆順を検出して救済）。
  // 異常があった場合は writeLog で Log シートに記録し、運用者が気付けるようにする。
  const inviteParse = parseReminderThresholds('invitation reminder', {
    first:  settings.firstReminderDays,
    second: settings.secondReminderDays,
    third:  settings.thirdReminderDays
  }, { first: 7, second: 14, third: 21 });
  const days = inviteParse.values;

  // 提出フェーズ用しきい値（Review_Deadline からの超過日数）。Bug 4 対応。
  // 負の値は期限前リマインドとして許容するためそのままバリデート。
  const subParse = parseReminderThresholds('submission reminder (overdue days)', {
    first:  settings.submissionReminderL1Days,
    second: settings.submissionReminderL2Days,
    third:  settings.submissionReminderL3Days
  }, { first: 0, second: 7, third: 14 });
  const subDays = subParse.values;

  // バリデーションで検出された警告を Log シートと実行ログに残す
  inviteParse.warnings.concat(subParse.warnings).forEach(w => {
    writeLog('[checkReminders] ' + w);
    Logger.log('[checkReminders WARN] ' + w);
  });

  // 親 manuscript の状態を一括ロード（決着済 / 新版あり / アーカイブ済の原稿に
  // 不要なリマインダを送らないため）。Manuscripts シートを 1 回だけ読み、
  // MsVer → 状態 と MS_ID → 最大Ver_No のルックアップを構築する。
  const msState = buildManuscriptStateLookup(ssId);

  Logger.log('checkReminders start — invite thresholds: '
    + days.first + '/' + days.second + '/' + days.third + ' days; '
    + 'submission overdue thresholds: '
    + subDays.first + '/' + subDays.second + '/' + subDays.third + ' days');

  checkEditorReminders(ssId, settings, days, msState);
  checkReviewerInvitationReminders(ssId, settings, days, msState);
  checkReviewerSubmissionReminders(ssId, settings, days, subDays, msState);

  Logger.log('checkReminders end');
}

/**
 * Manuscripts シートを 1 回読み、リマインダ判定に必要な親状態のルックアップを返す。
 * - msByMsVer  : { MsVer → { msId, verNo, accepted, stoppedByEic, finalStatus, sentBackAt } }
 * - maxVerByMsId : { MS_ID → 最大 Ver_No }（旧版へのリマインダ抑止に使用）
 */
function buildManuscriptStateLookup(ssId) {
  const result = { msByMsVer: {}, maxVerByMsId: {} };
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(MANUSCRIPTS_SHEET_NAME);
  if (!sheet) {
    // Bug 12: Manuscripts シートが見つからないと isManuscriptStillActive が
    // 全行で false を返し、リマインダが無音で完全停止する。Log シートに残す。
    const msg = 'buildManuscriptStateLookup: sheet "' + MANUSCRIPTS_SHEET_NAME
      + '" not found — all reminders will be skipped (parent state unverifiable)';
    writeLog('[checkReminders] ' + msg);
    Logger.log('[checkReminders WARN] ' + msg);
    return result;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return result;

  const headers = data[0];
  const msVerIdx       = headers.indexOf('MsVer');
  const msIdIdx        = headers.indexOf('MS_ID');
  const verNoIdx       = headers.indexOf('Ver_No');
  const acceptedIdx    = headers.indexOf('accepted');
  const stoppedIdx     = headers.indexOf('stoppedByEicAt');
  const finalStatusIdx = headers.indexOf('finalStatus');
  const sentBackIdx    = headers.indexOf('sentBackAt');

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const msVer = msVerIdx !== -1 ? String(row[msVerIdx] || '').trim() : '';
    const msId  = msIdIdx  !== -1 ? String(row[msIdIdx]  || '').trim() : '';
    const verNo = verNoIdx !== -1 ? Number(row[verNoIdx] || 0) : 0;

    if (msVer) {
      result.msByMsVer[msVer] = {
        msId:         msId,
        verNo:        verNo,
        accepted:     acceptedIdx    !== -1 ? String(row[acceptedIdx]    || '').trim().toLowerCase() : '',
        stoppedByEic: stoppedIdx     !== -1 ? String(row[stoppedIdx]     || '').trim() : '',
        finalStatus:  finalStatusIdx !== -1 ? String(row[finalStatusIdx] || '').trim() : '',
        sentBackAt:   sentBackIdx    !== -1 ? String(row[sentBackIdx]    || '').trim() : ''
      };
    }
    if (msId && verNo > (result.maxVerByMsId[msId] || 0)) {
      result.maxVerByMsId[msId] = verNo;
    }
  }
  return result;
}

/**
 * MsVer から親 manuscript がリマインダ送信を継続すべき状態かを判定。
 * 戻り値が false の場合、その行はスキップ（メールを送らない）。
 *
 * 判定で false にする条件:
 *   - Manuscripts に行が無い（アーカイブ済 等）
 *   - accepted === 'yes'（受理済）
 *   - stoppedByEicAt あり（EIC 早期却下）
 *   - finalStatus === 'final_review' / 'in_production'（最終確認・印刷工程）
 *   - sentBackAt あり（このラウンドの判定が著者に返送済）
 *   - 同 MS_ID で新しい Ver_No が投稿済（この版は旧ラウンド）
 */
function isManuscriptStillActive(msVer, msState) {
  if (!msVer) return false;
  const ms = msState.msByMsVer[msVer];
  if (!ms) return false;
  if (ms.accepted === 'yes') return false;
  if (ms.stoppedByEic) return false;
  if (ms.finalStatus === 'final_review' || ms.finalStatus === 'in_production') return false;
  if (ms.sentBackAt) return false;
  if (ms.msId && (msState.maxVerByMsId[ms.msId] || 0) > ms.verNo) return false;
  return true;
}

// ─────────────────────────────────────────
// 担当編集者候補への未回答リマインド
// ─────────────────────────────────────────
function checkEditorReminders(ssId, settings, days, msState) {
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(EDITOR_LOG_SHEET_NAME);
  if (!sheet) {
    // Bug 12: シート未存在の silent return を可視化（Log シート + 実行ログ）
    const msg = 'checkEditorReminders: sheet "' + EDITOR_LOG_SHEET_NAME
      + '" not found — skipping editor reminders';
    writeLog('[checkReminders] ' + msg);
    Logger.log('[checkReminders WARN] ' + msg);
    return;
  }

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

  // Bug 9: 必須列が無い場合は silent return ではなく Log シートに警告を残す。
  if (edtOkIdx === -1 || askAtIdx === -1 || editorKeyIdx === -1) {
    const missing = [];
    if (edtOkIdx     === -1) missing.push('edtOk');
    if (askAtIdx     === -1) missing.push('Ask_At');
    if (editorKeyIdx === -1) missing.push('editorKey');
    const msg = 'checkEditorReminders: required column(s) missing in '
      + EDITOR_LOG_SHEET_NAME + ': ' + missing.join(', ') + ' — skipping editor reminders';
    writeLog('[checkReminders] ' + msg);
    Logger.log('[checkReminders WARN] ' + msg);
    return;
  }

  // Bug 5 対応: 同一 MsVer ですでに 'ok' 回答した編集者がいる行集合を先に集める。
  // 一人でも承諾している manuscript は他候補へのリマインダ送信を打ち切る
  // （複数編集者を立てる運用は無いため、未回答候補は事実上「辞退と同等」）。
  const assignedMsVers = new Set();
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (String(r[edtOkIdx] || '').trim() === 'ok' && msVerIdx !== -1) {
      const v = String(r[msVerIdx] || '').trim();
      if (v) assignedMsVers.add(v);
    }
  }

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

    // Bug 5: 同 manuscript で他編集者がすでに承諾済みならスキップ
    if (msVer && assignedMsVers.has(msVer)) continue;

    // 親 manuscript が決着済 / 新版あり / アーカイブ済ならスキップ（Bug 3）
    if (!isManuscriptStillActive(msVer, msState)) continue;

    const rem1 = rem1Idx !== -1 ? String(row[rem1Idx] || '').trim() : '';
    const rem2 = rem2Idx !== -1 ? String(row[rem2Idx] || '').trim() : '';
    const rem3 = rem3Idx !== -1 ? String(row[rem3Idx] || '').trim() : '';

    const level = calcReminderLevel(elapsed, days, rem1, rem2, rem3);
    if (level === 0) continue;

    const sent = sendEditorReminderEmail(editorEmail, editorName, editorKey, msVer, elapsed, level, settings);

    // Bug 6: sent=false はクオータ枯渇等でキュー投入された状態。retrySendingEmails が
    // 後で配信するため、ここでタイムスタンプを記録しないと checkReminders の翌日実行で
    // 同じレベルが再発火し、キュー処理と合わせて二重送信が起きる。配信コミット済とみなす。
    const now     = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
    const colName = ['', 'Reminder1_At', 'Reminder2_At', 'Reminder3_At'][level];
    updateLogCell(ssId, EDITOR_LOG_SHEET_NAME, 'editorKey', editorKey, { [colName]: now });
    Logger.log('Editor reminder level ' + level + ' ' + (sent ? 'sent' : 'QUEUED') + ' to ' + editorEmail + ' for ' + msVer);
  }
}

// ─────────────────────────────────────────
// 査読者候補への未回答リマインド
// ─────────────────────────────────────────
function checkReviewerInvitationReminders(ssId, settings, days, msState) {
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(REVIEW_LOG_SHEET_NAME);
  if (!sheet) {
    // Bug 12: シート未存在の silent return を可視化（Log シート + 実行ログ）
    const msg = 'checkReviewerInvitationReminders: sheet "' + REVIEW_LOG_SHEET_NAME
      + '" not found — skipping reviewer invitation reminders';
    writeLog('[checkReminders] ' + msg);
    Logger.log('[checkReminders WARN] ' + msg);
    return;
  }

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

  // Bug 9: 必須列が無い場合は silent return ではなく Log シートに警告を残す。
  if (revOkIdx === -1 || askAtIdx === -1 || reviewKeyIdx === -1) {
    const missing = [];
    if (revOkIdx     === -1) missing.push('revOk');
    if (askAtIdx     === -1) missing.push('Ask_At');
    if (reviewKeyIdx === -1) missing.push('reviewKey');
    const msg = 'checkReviewerInvitationReminders: required column(s) missing in '
      + REVIEW_LOG_SHEET_NAME + ': ' + missing.join(', ') + ' — skipping reviewer invitation reminders';
    writeLog('[checkReminders] ' + msg);
    Logger.log('[checkReminders WARN] ' + msg);
    return;
  }

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

    // 親 manuscript が決着済 / 新版あり / アーカイブ済ならスキップ（Bug 3）
    if (!isManuscriptStillActive(msVer, msState)) continue;

    const rem1 = rem1Idx !== -1 ? String(row[rem1Idx] || '').trim() : '';
    const rem2 = rem2Idx !== -1 ? String(row[rem2Idx] || '').trim() : '';
    const rem3 = rem3Idx !== -1 ? String(row[rem3Idx] || '').trim() : '';

    const level = calcReminderLevel(elapsed, days, rem1, rem2, rem3);
    if (level === 0) continue;

    const sent = sendReviewerInvitationReminderEmail(revEmail, revName, reviewKey, msVer, elapsed, level, settings);

    // Bug 6: キュー投入(false)でもタイムスタンプを記録する（二重送信防止）
    const now     = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
    const colName = ['', 'Reminder1_At', 'Reminder2_At', 'Reminder3_At'][level];
    updateLogCell(ssId, REVIEW_LOG_SHEET_NAME, 'reviewKey', reviewKey, { [colName]: now });
    Logger.log('Reviewer invitation reminder level ' + level + ' ' + (sent ? 'sent' : 'QUEUED') + ' to ' + revEmail + ' for ' + msVer);
  }
}

// ─────────────────────────────────────────
// 査読承諾済み・未提出の査読者へのリマインド
// ─────────────────────────────────────────

/**
 * Review_log に SubReminder1_At / SubReminder2_At / SubReminder3_At 列が
 * 無ければ末尾に追加する。これらの列が無いと招待フェーズの Reminder*_At 列が
 * 流用され、招待リマインダを既に受けた査読者に対し提出リマインダが
 * 送信できなくなる不具合の温床となるため、実行ごとに存在を保証する。
 */
function ensureSubReminderColumns(sheet) {
  const required = ['SubReminder1_At', 'SubReminder2_At', 'SubReminder3_At'];
  const lastCol  = sheet.getLastColumn();
  const headers  = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(h => String(h).toLowerCase().trim());

  const missing = required.filter(name => headers.indexOf(name.toLowerCase()) === -1);
  if (missing.length === 0) return;

  let nextCol = lastCol + 1;
  missing.forEach(name => {
    sheet.getRange(1, nextCol).setValue(name);
    nextCol++;
  });
  SpreadsheetApp.flush();
  Logger.log('ensureSubReminderColumns: added to ' + sheet.getName() + ': ' + missing.join(', '));
}

function checkReviewerSubmissionReminders(ssId, settings, days, subDays, msState) {
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(REVIEW_LOG_SHEET_NAME);
  if (!sheet) {
    // Bug 12: シート未存在の silent return を可視化（Log シート + 実行ログ）
    const msg = 'checkReviewerSubmissionReminders: sheet "' + REVIEW_LOG_SHEET_NAME
      + '" not found — skipping reviewer submission reminders';
    writeLog('[checkReminders] ' + msg);
    Logger.log('[checkReminders WARN] ' + msg);
    return;
  }

  ensureSubReminderColumns(sheet);

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).toLowerCase().trim());
  const col     = name => headers.indexOf(name.toLowerCase().trim());

  const revOkIdx      = col('revok');
  const rcvAtIdx      = col('received_at');
  const answerAtIdx   = col('answer_at');
  const askAtIdx      = col('ask_at');
  const deadlineIdx   = col('review_deadline');
  const reviewKeyIdx  = col('reviewkey');
  const revNameIdx    = col('rev_name');
  const revEmailIdx   = col('rev_email');
  const msVerIdx      = col('msver');
  const folderUrlIdx  = col('reviewmaterialsfolderurl');
  const rem1Idx       = col('reminder1_at');
  const rem2Idx       = col('reminder2_at');
  const rem3Idx       = col('reminder3_at');

  // Bug 9: 必須列が無い場合は silent return ではなく Log シートに警告を残す。
  if (revOkIdx === -1 || reviewKeyIdx === -1) {
    const missing = [];
    if (revOkIdx     === -1) missing.push('revOk');
    if (reviewKeyIdx === -1) missing.push('reviewKey');
    const msg = 'checkReviewerSubmissionReminders: required column(s) missing in '
      + REVIEW_LOG_SHEET_NAME + ': ' + missing.join(', ') + ' — skipping reviewer submission reminders';
    writeLog('[checkReminders] ' + msg);
    Logger.log('[checkReminders WARN] ' + msg);
    return;
  }

  const today = new Date();

  for (let i = 1; i < data.length; i++) {
    const row   = data[i];
    const revOk = String(row[revOkIdx]  || '').trim();
    const rcvAt = rcvAtIdx !== -1 ? String(row[rcvAtIdx] || '').trim() : '';
    if (revOk !== 'ok') continue; // 未承諾・辞退はスキップ
    if (rcvAt !== '')   continue; // 提出済みはスキップ

    const reviewKey  = reviewKeyIdx !== -1 ? String(row[reviewKeyIdx] || '').trim() : '';
    const revName    = revNameIdx   !== -1 ? String(row[revNameIdx]   || '').trim() : '';
    const revEmail   = revEmailIdx  !== -1 ? String(row[revEmailIdx]  || '').trim() : '';
    const msVer      = msVerIdx     !== -1 ? String(row[msVerIdx]     || '').trim() : '';
    const folderUrl  = folderUrlIdx !== -1 ? String(row[folderUrlIdx] || '').trim() : '';
    if (!reviewKey || !revEmail) continue;

    // Bug 11: Review_Deadline は Sheets セルが「日付」フォーマットだと Date オブジェクトで
    // 返ってくる。String(dateObj) すると "Tue May 26 2026 00:00:00 GMT+0900" のような
    // 汚い表現になり、メール本文の表示が崩れ、parse のフォーマット依存も生じる。
    // Date / 文字列のどちらの型でも一貫した yyyy/MM/dd 表示と Date オブジェクトに正規化する。
    let deadline = '';
    let deadlineDate = null;
    if (deadlineIdx !== -1) {
      const raw = row[deadlineIdx];
      if (raw instanceof Date) {
        deadlineDate = raw;
        deadline = Utilities.formatDate(raw, 'Asia/Tokyo', 'yyyy/MM/dd');
      } else if (raw !== null && raw !== undefined && raw !== '') {
        deadline = String(raw).trim();
        if (deadline) {
          const d = new Date(deadline);
          if (!isNaN(d)) deadlineDate = d;
        }
      }
    }

    // 親 manuscript が決着済 / 新版あり / アーカイブ済ならスキップ（Bug 3）
    if (!isManuscriptStillActive(msVer, msState)) continue;

    // Bug 4: Review_Deadline がある場合は「期限超過日数」基準で判定。
    // 無効/未設定なら従来の Answer_At 基準にフォールバック（後方互換）。
    let elapsed;
    let activeDays;
    if (deadlineDate) {
      elapsed   = Math.floor((today - deadlineDate) / (1000 * 60 * 60 * 24));
      activeDays = subDays;
    } else {
      // フォールバック: 承諾日（Answer_At）または Ask_At からの経過日数で判定
      const baseRaw = (answerAtIdx !== -1 && row[answerAtIdx])
        ? row[answerAtIdx] : (askAtIdx !== -1 ? row[askAtIdx] : null);
      if (!baseRaw) continue;
      const baseDate = new Date(baseRaw);
      if (isNaN(baseDate)) continue;
      elapsed   = Math.floor((today - baseDate) / (1000 * 60 * 60 * 24));
      activeDays = days;
    }

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

    const level = calcReminderLevel(elapsed, activeDays, rem1, rem2, rem3);
    if (level === 0) continue;

    // Bug 13: deadline ベースで判定した場合のみ、elapsed の符号で本文の緊急度文言を
    // 書き分ける。フォールバック（Answer_At ベース）では deadline 不明のため
    // 汎用文言を使う。
    const sent = sendReviewerSubmissionReminderEmail(
      revEmail, revName, reviewKey, msVer, deadline, folderUrl, elapsed, level,
      Boolean(deadlineDate), settings
    );

    // Bug 6: キュー投入(false)でもタイムスタンプを記録する（二重送信防止）
    const now     = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
    const colName = ['', 'Reminder1_At', 'Reminder2_At', 'Reminder3_At'][level];
    const subColName = ['', 'SubReminder1_At', 'SubReminder2_At', 'SubReminder3_At'][level];
    updateLogCell(ssId, REVIEW_LOG_SHEET_NAME, 'reviewKey', reviewKey,
      { [useSubCol ? subColName : colName]: now });
    Logger.log('Reviewer submission reminder level ' + level + ' ' + (sent ? 'sent' : 'QUEUED') + ' to ' + revEmail + ' for ' + msVer);
  }
}

// ─────────────────────────────────────────
// ユーティリティ
// ─────────────────────────────────────────

/**
 * Settings からリマインダしきい値を取り出してバリデートする（Bug 7 対応）。
 *
 * @param {string} label  - エラーメッセージ用ラベル（例: 'invitation reminder'）
 * @param {{first,second,third}} raw - Settings から取った生値（任意の型）
 * @param {{first,second,third}} defaults - 異常時のフォールバック値（数値）
 * @returns {{ values: {first,second,third}, warnings: string[] }}
 *
 * 検出する異常:
 *   1. parseInt 結果が NaN（非数値・空白文字列等）→ デフォルトに置換
 *   2. first <= second <= third になっていない（逆順） → 昇順ソートして救済
 *      ※ submission 系は負の値（期限前警告）も許容するため、絶対値ではなく
 *        単純な数値比較で並び替える。
 */
function parseReminderThresholds(label, raw, defaults) {
  const warnings = [];
  const values = {};
  ['first', 'second', 'third'].forEach(k => {
    const rawVal = raw[k];
    const trimmed = (rawVal === null || rawVal === undefined) ? '' : String(rawVal).trim();
    const n = parseInt(trimmed, 10);
    if (trimmed === '' || isNaN(n)) {
      if (trimmed !== '') {
        // 空欄は「未設定なのでデフォルト」で正常。値があるのに NaN は警告対象。
        warnings.push(label + ': ' + k + 'ReminderDays value "' + trimmed
          + '" is not a valid integer; using default ' + defaults[k]);
      }
      values[k] = defaults[k];
    } else {
      values[k] = n;
    }
  });
  if (!(values.first <= values.second && values.second <= values.third)) {
    const sorted = [values.first, values.second, values.third].sort((a, b) => a - b);
    warnings.push(label + ': thresholds (' + values.first + '/' + values.second + '/' + values.third
      + ') are not monotonically non-decreasing; auto-sorting to (' + sorted.join('/') + ')');
    values.first  = sorted[0];
    values.second = sorted[1];
    values.third  = sorted[2];
  }
  return { values: values, warnings: warnings };
}

/**
 * 経過日数と送信済みフラグから送信すべきレベル (1/2/3) を返す。
 * 送信不要なら 0 を返す。
 *
 * 【設計方針】
 * 「未送信のうち最低レベルから順に」発火させる。これは、トリガが数日停止して
 * 一気に複数レベルが期限超過になった場合でも、受信者には L1 → L2 → L3 の順で
 * 段階的に届けるため。一度に最終リマインダだけが飛ぶといった事態を防ぐ。
 *
 * 例: elapsed=21、rem1/rem2/rem3 すべて未送信 →
 *     今日 L1、翌日 L2、翌々日 L3（毎日トリガで段階的送信）。
 *
 * 注意: days.first <= days.second <= days.third を前提とする。
 *       Settings の数値が逆転しているとリマインダが永久に飛ばないことがある。
 */
function calcReminderLevel(elapsed, days, rem1, rem2, rem3) {
  if (elapsed >= days.first  && !rem1) return 1;
  if (elapsed >= days.second && !rem2) return 2;
  if (elapsed >= days.third  && !rem3) return 3;
  return 0;
}

// ─────────────────────────────────────────
// メール送信関数
// ─────────────────────────────────────────

function sendEditorReminderEmail(toEmail, toName, editorKey, msVer, elapsed, level, settings) {
  const webAppUrl   = ScriptApp.getService().getUrl();
  const responseUrl = webAppUrl + '?editorKey=' + editorKey;
  const isFinal     = level >= 3;

  // Bug 10: HTML 補間する動的値はすべて escHtml でエスケープする
  // （subject はプレーンテキスト扱いなので未エスケープのまま）
  const msVerEsc  = escHtml(msVer);
  const toNameEsc = escHtml(toName);

  const subject = isFinal
    ? `[${settings.Journal_Name}] [最終リマインド / Final Reminder] 担当編集者依頼への回答をお願いします / Please respond to editor assignment invitation — ${msVer}`
    : `[${settings.Journal_Name}] [リマインド ${level} / Reminder ${level}] 担当編集者ご就任のご依頼 / Editor assignment invitation — ${msVer}`;

  const bodyHtml = `
    <p>${isFinal ? '<strong style="color:#dc2626;">[Final Reminder]</strong> ' : ''}You received an invitation to serve as a responsible editor for manuscript <strong>${msVerEsc}</strong>, but we have not yet received your response.</p>
    <p>It has been <strong>${elapsed} days</strong> since the invitation was sent. Please respond at your earliest convenience by clicking the button below.</p>
    <hr style="border:none; border-top:1px solid #e2e8f0; margin:20px 0;">
    <p>${isFinal ? '<strong>【最終リマインド】</strong>' : ''}原稿 <strong>${msVerEsc}</strong> の担当編集者ご就任のご依頼をお送りしてから <strong>${elapsed}日</strong> が経過しておりますが、まだご回答をいただいておりません。お手数ですが、以下のボタンよりご回答ください。</p>
  `;

  const html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting:    `Dear Dr. ${toNameEsc},`,
    bodyHtml:    bodyHtml,
    buttonUrl:   responseUrl,
    buttonLabel: 'Respond to Invitation / 依頼に回答する',
    footerHtml:  settings.mailFooter || ''
  });

  return sendEmailSafe({ to: toEmail, subject, htmlBody: html },
    'Editor Reminder L' + level + ': ' + msVer + ' to ' + toName);
}

function sendReviewerInvitationReminderEmail(toEmail, toName, reviewKey, msVer, elapsed, level, settings) {
  const webAppUrl   = ScriptApp.getService().getUrl();
  const responseUrl = webAppUrl + '?reviewKey=' + reviewKey;
  const isFinal     = level >= 3;

  // Bug 10: HTML 補間する動的値はすべて escHtml でエスケープする
  const msVerEsc  = escHtml(msVer);
  const toNameEsc = escHtml(toName);

  const subject = isFinal
    ? `[${settings.Journal_Name}] [最終リマインド / Final Reminder] 査読依頼への回答をお願いします / Please respond to reviewer invitation — ${msVer}`
    : `[${settings.Journal_Name}] [リマインド ${level} / Reminder ${level}] 査読のご依頼 / Reviewer invitation — ${msVer}`;

  const bodyHtml = `
    <p>${isFinal ? '<strong style="color:#dc2626;">[Final Reminder]</strong> ' : ''}You received an invitation to review manuscript <strong>${msVerEsc}</strong>, but we have not yet received your response.</p>
    <p>It has been <strong>${elapsed} days</strong> since the invitation was sent. Please respond by clicking the button below.</p>
    <hr style="border:none; border-top:1px solid #e2e8f0; margin:20px 0;">
    <p>${isFinal ? '<strong>【最終リマインド】</strong>' : ''}原稿 <strong>${msVerEsc}</strong> の査読依頼をお送りしてから <strong>${elapsed}日</strong> が経過しておりますが、まだご回答をいただいておりません。以下のボタンよりご回答ください。</p>
  `;

  const html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting:    `Dear Dr. ${toNameEsc},`,
    bodyHtml:    bodyHtml,
    buttonUrl:   responseUrl,
    buttonLabel: 'Respond to Invitation / 依頼に回答する',
    footerHtml:  settings.mailFooter || ''
  });

  return sendEmailSafe({ to: toEmail, subject, htmlBody: html },
    'Reviewer Invitation Reminder L' + level + ': ' + msVer + ' to ' + toName);
}

function sendReviewerSubmissionReminderEmail(
  toEmail, toName, reviewKey, msVer, deadline, folderUrl, elapsed, level, hasDeadline, settings
) {
  const webAppUrl      = ScriptApp.getService().getUrl();
  const reviewMenuUrl  = webAppUrl + '?reviewKey=' + reviewKey;
  const isFinal        = level >= 3;

  // Bug 10: HTML 補間する動的値はすべて escHtml でエスケープする
  // folderUrl は href 属性内に入るため、" や ' が含まれた場合に属性が破綻する
  // のを防ぐためにもエスケープが必要。
  const msVerEsc    = escHtml(msVer);
  const toNameEsc   = escHtml(toName);
  const deadlineEsc = escHtml(deadline);
  const folderUrlEsc = escHtml(folderUrl);

  const subject = isFinal
    ? `[${settings.Journal_Name}] [最終リマインド / Final Reminder] 査読結果の提出をお願いします / Please submit your review — ${msVer}`
    : `[${settings.Journal_Name}] [リマインド ${level} / Reminder ${level}] 査読結果提出期限 / Review submission due — ${msVer}`;

  const deadlineRow = deadline
    ? `<tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee; width:30%;">Review Deadline / 査読期限</th>
       <td style="padding:8px; border-bottom:1px solid #eee;"><strong>${deadlineEsc}</strong></td></tr>`
    : '';

  const folderRow = folderUrl
    ? `<tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">Review Materials / 査読資料</th>
       <td style="padding:8px; border-bottom:1px solid #eee;"><a href="${folderUrlEsc}" target="_blank">査読資料フォルダを開く / Open Review Materials</a></td></tr>`
    : '';

  // Bug 13: deadline 基準で判定した場合は elapsed の符号で本文の緊急度文言を
  // 書き分ける（期限前 / 期限当日 / 期限超過の 3 段階）。フォールバック判定
  // （hasDeadline=false）では deadline 不明のため、従来の汎用文言を使う。
  let urgencyEn, urgencyJa;
  if (hasDeadline) {
    if (elapsed < 0) {
      const daysRemaining = -elapsed;
      const dayWord = daysRemaining === 1 ? 'day' : 'days';
      urgencyEn = `This is a friendly reminder that your review for manuscript <strong>${msVerEsc}</strong> is due in <strong>${daysRemaining} ${dayWord}</strong>.`;
      urgencyJa = `原稿 <strong>${msVerEsc}</strong> の査読期限まで <strong>あと ${daysRemaining}日</strong> となりました。`;
    } else if (elapsed === 0) {
      urgencyEn = `Your review for manuscript <strong>${msVerEsc}</strong> is due <strong>today</strong>.`;
      urgencyJa = `原稿 <strong>${msVerEsc}</strong> の査読期限は <strong>本日</strong> です。`;
    } else {
      const dayWord = elapsed === 1 ? 'day' : 'days';
      urgencyEn = `Your review for manuscript <strong>${msVerEsc}</strong> is now <strong>${elapsed} ${dayWord} overdue</strong>.`;
      urgencyJa = `原稿 <strong>${msVerEsc}</strong> の査読期限を <strong>${elapsed}日</strong> 過ぎており、まだご提出いただけておりません。`;
    }
  } else {
    // フォールバック（deadline 未設定）— 期限が分からないので汎用文言
    urgencyEn = `This is a reminder that your review for manuscript <strong>${msVerEsc}</strong> has not yet been submitted.`;
    urgencyJa = `原稿 <strong>${msVerEsc}</strong> の査読結果がまだ提出されておりません。`;
  }

  const bodyHtml = `
    <p>${isFinal ? '<strong style="color:#dc2626;">[Final Reminder]</strong> ' : ''}${urgencyEn} Please submit your review results via the button below.</p>
    <p>${isFinal ? '<strong>【最終リマインド】</strong>' : ''}${urgencyJa} お手数ですが、以下のボタンより査読結果をご提出ください。</p>
    <table style="width:100%; font-size:14px; border-collapse:collapse; margin:20px 0;">
      <tr><th style="text-align:left; padding:8px; border-bottom:1px solid #eee; width:30%;">Manuscript / 原稿</th>
          <td style="padding:8px; border-bottom:1px solid #eee;">${msVerEsc}</td></tr>
      ${deadlineRow}
      ${folderRow}
    </table>
  `;

  const html = renderRichEmail({
    journalName: settings.Journal_Name,
    greeting:    `Dear Dr. ${toNameEsc},`,
    bodyHtml:    bodyHtml,
    buttonUrl:   reviewMenuUrl,
    buttonLabel: 'Submit Review / 査読結果を提出する',
    footerHtml:  settings.mailFooter || ''
  });

  return sendEmailSafe({ to: toEmail, subject, htmlBody: html },
    'Reviewer Submission Reminder L' + level + ': ' + msVer + ' to ' + toName);
}

// ─────────────────────────────────────────
// トリガー設定（Bug 8 対応）
// ─────────────────────────────────────────

/**
 * リマインダ系の時間ベーストリガを一括登録する（Bug 8 対応）。
 *
 * GAS エディタからこの関数を一度実行すれば、以下のトリガが自動設定される:
 *   - checkReminders         : 毎日 09:00 JST（未回答編集者・査読者へのリマインド）
 *   - retrySendingEmails     : 毎日 12:00 JST（クオータ枯渇等で滞留した送信キューを再処理）
 *
 * 既存の同名トリガは削除してから再登録するため、複数回実行しても重複しない。
 *
 * 【なぜ retrySendingEmails も含めるか】
 * Bug 6 修正以降、checkReminders は sendEmailSafe がキューに投入した場合でも
 * タイムスタンプを記録して二重送信を防ぐ。そのため retrySendingEmails トリガが
 * 動かないとリマインダが「永久にキューから出てこない」障害になる。
 * リマインダ機能の信頼性を保つために、両者をワンセットで登録する。
 */
function setupReminderTriggers() {
  // 既存のリマインダ関連トリガーを削除（重複防止）
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'checkReminders' || fn === 'retrySendingEmails') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 1. リマインダ判定: 毎日 09:00 JST
  ScriptApp.newTrigger('checkReminders')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();

  // 2. 滞留メール再送: 毎日 12:00 JST
  //    （checkReminders がキュー投入してから数時間置いて、クオータ復活を待つ）
  ScriptApp.newTrigger('retrySendingEmails')
    .timeBased()
    .atHour(12)
    .everyDays(1)
    .create();

  Logger.log('Reminder triggers have been set up: checkReminders @09:00, retrySendingEmails @12:00');
}
