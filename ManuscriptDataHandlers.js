/**
 * ManuscriptDataHandlers.js - getManuscriptData() 関数の分割リファクタリング
 * 各ロール別のデータ取得ハンドラーを定義
 */

/**
 * 共通ユーティリティ関数
 */

/**
 * Dateオブジェクトを文字列に変換（JSONシリアライズエラー防止）
 */
function convertDatesToStrings(data) {
  if (!data) return data;
  
  Object.keys(data).forEach(k => {
    if (data[k] instanceof Date) {
      data[k] = Utilities.formatDate(data[k], 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    }
  });
  
  return data;
}

/**
 * タイムゾーン対応の日付フォーマット関数を生成
 */
function createDateFormatter(ssId) {
  // まず Settings キャッシュに乗せたタイムゾーンを使う（API 呼び出し不要）
  const cachedSettings = getSettingsOptimized(ssId);
  const timezone = (cachedSettings && cachedSettings._timezone)
    ? cachedSettings._timezone
    : SpreadsheetApp.openById(ssId).getSpreadsheetTimeZone();
  return function(val) {
    if (!val) return '';
    if (val instanceof Date) return Utilities.formatDate(val, timezone, 'yyyy/MM/dd HH:mm');
    return String(val).trim();
  };
}

/**
 * フォルダ情報を取得してデータに追加
 *
 * _reviewMaterialsFolderUrl はグローバル設定値であり原稿状態に依存しないため、
 * folderUrl / submittedFiles の有無に関わらず常に解決する。
 * 原稿固有のフォルダ探索（submittedFolderUrl, submittedFiles）は未取得の場合のみ実行する。
 */
function enrichWithFolderInfo(msData, ssId) {
  try {
    const settings = getSettingsOptimized(ssId);
    const rootName = settings.SUBFOLDER || 'Journal Files';
    const root = driveFolderCache.getRootFolder(rootName);

    // ① 審査票・作業フォルダURL は常に解決する（ダッシュボード全ロールで必要）
    const materialName = settings.reviewMaterialsFolder || '';
    let materialUrl = '';
    if (materialName && root) {
      try {
        const matFolder = driveFolderCache.getFolderByName(root, materialName);
        if (matFolder) materialUrl = matFolder.getUrl();
      } catch (e) {
        Logger.log('reviewMaterialsFolder lookup failed in enrichWithFolderInfo: ' + e.message);
      }
    }
    msData._reviewMaterialsFolderUrl = materialUrl;

    // ② 投稿フォルダ・ファイル一覧は未取得の場合のみ解決する
    const needsFolderLookup = !msData.folderUrl
      || !msData.submittedFiles
      || msData.submittedFiles === 'なし'
      || msData.submittedFiles === '';

    if (needsFolderLookup && root) {
      const msFolder = driveFolderCache.getFolderByName(root, msData.MS_ID);
      if (msFolder) {
        const verNo = msData.Ver_No || 1;
        const verFolder = driveFolderCache.getFolderByName(msFolder, 'ver.' + verNo);
        if (verFolder) {
          const subFolder = driveFolderCache.getFolderByName(verFolder, 'submitted');
          if (subFolder) {
            msData.submittedFolderUrl = subFolder.getUrl();
            // folderUrl が未設定（初回アクセス）の場合のみ共有設定を行う
            // 既に設定済みなら setSharing() の Drive API 呼び出しを省略して高速化
            if (!msData.folderUrl) {
              subFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            }

            // 互換性のため、従来通り folderUrl が空ならセットする
            if (!msData.folderUrl) {
              msData.folderUrl = msData.submittedFolderUrl;
            }

            if (!msData.submittedFiles || msData.submittedFiles === 'なし' || msData.submittedFiles === '') {
              const files = subFolder.getFiles();
              const fileNames = [];
              while (files.hasNext()) {
                fileNames.push(files.next().getName());
              }
              msData.submittedFiles = fileNames.length > 0 ? fileNames.join(', ') : 'なし';
            }
          }
        }
      }
    }
  } catch(e) {
    Logger.log('Error fetching folder info from Drive: ' + e);
  }

  return msData;
}

/**
 * 査読者リストを構築する共通関数
 */
function buildReviewerList(reviewLogs, dateFormatter) {
  return reviewLogs.map(log => {
    const revOk = String(log.revOk || '').trim();
    const receivedAt = dateFormatter(log.Received_At);
    const submitted = revOk === 'ok' && receivedAt !== '';

    return {
      reviewKey: String(log.reviewKey || '').trim(),
      Rev_Name: String(log.Rev_Name || '').trim(),
      Rev_Email: String(log.Rev_Email || '').trim(),
      revOk,
      Ask_At: dateFormatter(log.Ask_At),
      Answer_At: dateFormatter(log.Answer_At),
      Received_At: receivedAt,
      Review_Deadline: String(log.Review_Deadline || log.review_deadline || '').trim(),
      Score: String(log.Score || '').trim(),
      // 査読者本人のアップロードフォルダURL（Review_log.reviewerUploadFolderUrl 列）
      folderUrl: String(log.reviewerUploadFolderUrl || '').trim(),
      // コメント本文は遅延取得（apiGetReviewComments で on-demand に読み込む）
      openCommentsId: submitted ? String(log.openCommentsId || '').trim() : '',
      confidentialCommentsId: submitted ? String(log.confidentialCommentsId || '').trim() : ''
    };
  });
}

/**
 * 進捗ステータスを決定する関数
 */
function determineProgressStatus(msData, editorLogs, reviewLogs) {
  // 投稿直後却下: 委員長が審査を停止した原稿は最優先で返す
  if (String(msData.stoppedByEicAt || '').trim()) {
    return 'eic_stopped';
  }

  const acceptedEditor = editorLogs.find(log => String(log.edtOk || '').trim() === 'ok');
  const acceptedReviewers = reviewLogs.filter(log => String(log.revOk || '').trim() === 'ok');
  const completedReviews = reviewLogs.filter(log => String(log.Received_At || '').trim() !== '');

  // 全査読が完了したかどうか（承諾済み査読者全員が結果を提出）
  const allReviewsCompleted = acceptedReviewers.length > 0 &&
                             acceptedReviewers.length === completedReviews.length;

  // --- 拡張ステータス判定: 受理後の最終確認フェーズ ---
  const finalStatus = String(msData.finalStatus || '').trim();
  const managingEditorKey = String(msData.managingEditorKey || '').trim();

  // in_production / eicFinalDecision は最優先で判定（managingEditorKey チェックより前に置く）
  if (finalStatus === 'in_production') {
    return 'in_production';
  }
  if (msData.eicFinalDecision) {
    return 'decision';
  }

  // 編集幹事(ME)用キーが発行されており、まだ判定が著者に送られていない場合は「最終確認中」
  if (managingEditorKey && !msData.sentBackAt && !msData.score) {
    return 'final_review';
  }

  if (finalStatus === 'final_review') {
    return 'final_review';
  }

  if (msData.sentBackAt || msData.score) {
    return 'decision';
  } else if (acceptedEditor && acceptedEditor.Score) {
    return 'reviewed';
  } else if (allReviewsCompleted) {
    return 'reviewed'; // 全査読完了 → 査読完了ステータス
  } else if (acceptedReviewers.length > 0) {
    return 'reviewing';
  } else if (acceptedEditor) {
    return 'editor_assigned';
  } else if (editorLogs.length > 0) {
    return 'editor_requested';
  } else {
    return 'submitted';
  }
}

/**
 * ロール別ハンドラー関数
 */

/**
 * 著者用データ取得ハンドラー
 */
function getAuthorManuscriptData(ssId, key) {
  const msData = findRecordByKey(ssId, MANUSCRIPTS_SHEET_NAME, 'key', key);

  if (!msData) return null;

  const msVer = msData.MsVer || '';
  const editorLogs = findAllRecordsByKey(ssId, EDITOR_LOG_SHEET_NAME, 'MsVer', msVer);
  const reviewLogs = findAllRecordsByKey(ssId, REVIEW_LOG_SHEET_NAME, 'MsVer', msVer);

  msData._progressStatus = determineProgressStatus(msData, editorLogs, reviewLogs);

  // 判定スコアに基づいて「再提出が可能か」を判定してフラグを付記
  if (msData.score) {
    const templates = getDecisionTemplates(ssId, msData.score);
    msData._allowsResubmit = !!templates.allowsResubmit;
  } else {
    msData._allowsResubmit = false;
  }

  // 再投稿済みかどうかを確認（同じ MS_ID でより新しいバージョンが存在するか）
  const currentVerNo = Number(msData.Ver_No || 1);
  const msId = msData.MS_ID || '';
  if (msId && currentVerNo) {
    const allVersions = findAllRecordsByKey(ssId, MANUSCRIPTS_SHEET_NAME, 'MS_ID', msId);
    msData._hasNewerVersion = allVersions.some(v => Number(v.Ver_No || 0) > currentVerNo);
  } else {
    msData._hasNewerVersion = false;
  }

  return msData;
}

/**
 * 編集委員長用データ取得ハンドラー
 */
function getEicManuscriptData(ssId, key) {
  const msData = findRecordByKey(ssId, MANUSCRIPTS_SHEET_NAME, 'eicKey', key);
  
  if (!msData) return null;
  
  // EIC判定フォーム用スコア選択肢
  msData._scoreOptions = getScoreOptions();

  // 「最終承認・再投稿不可」となり ME ルートへ自動リダイレクトされるスコア一覧
  // （クライアント側で確認ダイアログの文言を切り替えるために使用）
  msData._meRedirectScores = (msData._scoreOptions || []).filter(function(sc) {
    try {
      var t = getDecisionTemplates(ssId, sc);
      return !!(t && t.isAccepted && !t.allowsResubmit);
    } catch (e) {
      return false;
    }
  });

  // EIC最終アクション（ルートa）用：DecisionMailシートの判定選択肢
  const ssIdForOptions = getSpreadsheetId();
  msData._decisionMailOptions = ssIdForOptions ? getDecisionMailOptions(ssIdForOptions) : [];
  
  // 担当編集者の承諾状況を確認
  const editorLogs = findAllRecordsByKey(ssId, EDITOR_LOG_SHEET_NAME, 'MsVer', msData.MsVer || '');
  const acceptedEditor = editorLogs.find(log => String(log.edtOk || '').trim() === 'ok');
  const pendingEditor = editorLogs.find(log => String(log.edtOk || '').trim() === '');
  
  msData._editorAccepted = !!acceptedEditor;
  msData._acceptedEditorName = acceptedEditor ? (acceptedEditor.Editor_Name || '') : '';
  msData._editorPending = !!pendingEditor;
  msData._pendingEditorName = pendingEditor ? (pendingEditor.Editor_Name || '') : '';
  msData._editorScore = acceptedEditor
    ? String(acceptedEditor.Score || '').trim()
    : '';
  
  // EIC用: タイムゾーン対応の日付フォーマット
  const eicFmt = createDateFormatter(ssId);
  
  // EIC用: 担当編集者一覧
  msData._editorList = editorLogs.map(log => ({
    editorKey: String(log.editorKey || '').trim(),
    Editor_Name: String(log.Editor_Name || '').trim(),
    Editor_Email: String(log.Editor_Email || '').trim(),
    edtOk: String(log.edtOk || '').trim(),
    Ask_At: eicFmt(log.Ask_At),
    Answer_At: eicFmt(log.Answer_At),
    Score: String(log.Score || '').trim(),
    Message: String(log.Message || '').trim(),
    ConfidentialMessage: String(log.ConfidentialMessage || '').trim(),
    reportPdfUrl:              String(log.reportPdfUrl              || '').trim(),
    reportWordUrl:             String(log.reportWordUrl             || '').trim(),
    reportFolderUrl:           String(log.reportFolderUrl           || '').trim(),
    reportAttachmentsFolderUrl: String(log.reportAttachmentsFolderUrl || '').trim(),
    reportGoogleDocId:         String(log.reportGoogleDocId         || '').trim(),
    reportCommentPdfUrl:       String(log.reportCommentPdfUrl       || '').trim()
  }));

  // EIC用: 査読者一覧
  const eicReviewLogs = findAllRecordsByKey(ssId, REVIEW_LOG_SHEET_NAME, 'MsVer', msData.MsVer || '');
  const eicCompleted = eicReviewLogs.filter(log => String(log.Received_At || '').trim() !== '');
  
  msData._reviewCompletedCount = eicCompleted.length;
  msData._reviewerAcceptedCount = eicReviewLogs.filter(log => String(log.revOk || '').trim() === 'ok').length;
  msData._reviewerList = buildReviewerList(eicReviewLogs, eicFmt);
  
  const eicHasIncomplete = msData._reviewerList.some(r =>
    r.revOk === '' || (r.revOk === 'ok' && r.Received_At === '')
  );
  msData._allReviewsIn = (msData._reviewCompletedCount > 0) && !eicHasIncomplete;
  
  // EIC用: 進捗ステータスの設定
  msData._progressStatus = determineProgressStatus(msData, editorLogs, eicReviewLogs);
  
  // EIC用: 全バージョン履歴
  const msIdForHistory = String(msData.MS_ID || '').trim();
  if (msIdForHistory) {
    const allVersions = findAllRecordsByKey(ssId, MANUSCRIPTS_SHEET_NAME, 'MS_ID', msIdForHistory);
    allVersions.sort((a, b) => Number(a.Ver_No || 0) - Number(b.Ver_No || 0));
    
    msData._versionHistory = allVersions.map(v => {
      const vMsVer = String(v.MsVer || '').trim();
      const vEditorLogs = findAllRecordsByKey(ssId, EDITOR_LOG_SHEET_NAME, 'MsVer', vMsVer);
      const vAcceptedEditor = vEditorLogs.find(log => String(log.edtOk || '').trim() === 'ok');

      return {
        MsVer: vMsVer,
        Ver_No: String(v.Ver_No || '').trim(),
        Submitted_At: eicFmt(v.Submitted_At),
        score: String(v.score || v.Score || '').trim(),
        sentBackAt: eicFmt(v.sentBackAt || v.SentBackAt || ''),
        accepted: String(v.accepted || '').trim(),
        reportPdfUrl: vAcceptedEditor ? String(vAcceptedEditor.reportPdfUrl || '').trim() : ''
      };
    });

    // 前バージョンの担当編集者情報（再投稿時の編集者指名フォームの初期値として使用）
    const currentVerNo = Number(msData.Ver_No || 1);
    if (currentVerNo > 1) {
      const prevVersion = allVersions.find(v => Number(v.Ver_No || 0) === currentVerNo - 1);
      if (prevVersion) {
        const prevMsVer = String(prevVersion.MsVer || '').trim();
        const prevEditorLogs = findAllRecordsByKey(ssId, EDITOR_LOG_SHEET_NAME, 'MsVer', prevMsVer);
        const prevAcceptedEditor = prevEditorLogs.find(
          log => String(log.edtOk || '').trim() === 'ok'
        );
        if (prevAcceptedEditor) {
          msData._prevEditorName  = String(prevAcceptedEditor.Editor_Name  || '').trim();
          msData._prevEditorEmail = String(prevAcceptedEditor.Editor_Email || '').trim();
        }
      }
    }
  }
  
  return msData;
}

/**
 * EIC全体管理: 全原稿の進捗サマリーを一括取得
 * spreadsheetCache 経由でシートを読み込み（prewarm 済みなら API 呼び出し不要）
 */
function getEicAllMsData(ssId) {
  const fmt = createDateFormatter(ssId);

  // spreadsheetCache に展開（既に prewarm 済みならキャッシュヒット = API 呼び出しなし）
  spreadsheetCache.prewarmSheets(ssId, [
    MANUSCRIPTS_SHEET_NAME,
    EDITOR_LOG_SHEET_NAME,
    REVIEW_LOG_SHEET_NAME
  ]);

  // ── キャッシュから全原稿行を取得
  const msCache = spreadsheetCache.getSheetData(ssId, MANUSCRIPTS_SHEET_NAME);
  if (!msCache || msCache.rows.length === 0) return [];
  const msHeaders = msCache.headers.map(function(h) { return String(h).trim(); });

  // ── キャッシュから Editor_log・Review_log を取得してオブジェクト配列に変換
  function cacheToObjArray(cacheEntry) {
    if (!cacheEntry) return [];
    var hdrs = cacheEntry.headers.map(function(h) { return String(h).trim(); });
    return cacheEntry.rows.map(function(row) {
      var obj = {};
      hdrs.forEach(function(h, i) { obj[h] = row[i]; });
      return obj;
    });
  }

  var allEditorLogs = cacheToObjArray(spreadsheetCache.getSheetData(ssId, EDITOR_LOG_SHEET_NAME));
  var allReviewLogs = cacheToObjArray(spreadsheetCache.getSheetData(ssId, REVIEW_LOG_SHEET_NAME));

  // ── 各原稿の進捗を集計
  var result = [];
  for (var mi = 0; mi < msCache.rows.length; mi++) {
    var ms = {};
    msHeaders.forEach(function(h, i) { ms[h] = msCache.rows[mi][i]; });

    var msVer = String(ms.MsVer || '').trim();
    if (!msVer) continue;

    var edLogs  = allEditorLogs.filter(function(log) { return String(log.MsVer || '').trim() === msVer; });
    var revLogs = allReviewLogs.filter(function(log) { return String(log.MsVer || '').trim() === msVer; });

    var acceptedEditor   = edLogs.find(function(log) { return String(log.edtOk || '').trim() === 'ok'; });
    var reviewerAccepted = revLogs.filter(function(log) { return String(log.revOk || '').trim() === 'ok'; }).length;
    var reviewCompleted  = revLogs.filter(function(log) { return String(log.Received_At || '').trim() !== ''; }).length;

    var status = determineProgressStatus(ms, edLogs, revLogs);

    result.push({
      MsVer:            msVer,
      MS_ID:            String(ms.MS_ID     || '').trim(),
      Ver_No:           String(ms.Ver_No    || '').trim(),
      TitleJP:          String(ms.TitleJP   || '').trim(),
      TitleEN:          String(ms.TitleEN   || '').trim(),
      CA_Name:          String(ms.CA_Name   || '').trim(),
      AuthorsJP:        String(ms.AuthorsJP || '').trim(),
      MS_Type:          String(ms.MS_Type   || '').trim(),
      Submitted_At:     fmt(ms.Submitted_At),
      sentBackAt:       fmt(ms.sentBackAt   || ms.SentBackAt || ''),
      score:            String(ms.score     || ms.Score      || '').trim(),
      finalStatus:      String(ms.finalStatus || '').trim(),
      eicKey:           String(ms.eicKey    || '').trim(),
      editorName:       acceptedEditor ? String(acceptedEditor.Editor_Name || '').trim() : '',
      reviewerAccepted: reviewerAccepted,
      reviewCompleted:  reviewCompleted,
      _progressStatus:  status
    });
  }

  // アクティブな原稿を優先してソート
  var sortOrder = [
    'submitted', 'editor_requested', 'editor_assigned',
    'reviewing', 'reviewed', 'final_review', 'eic_stopped',
    'decision', 'in_production'
  ];
  result.sort(function(a, b) {
    var ai = sortOrder.indexOf(a._progressStatus);
    var bi = sortOrder.indexOf(b._progressStatus);
    return (ai === -1 ? 99 : ai) - (bi === -1 ? 99 : bi);
  });

  return result;
}

/**
 * 担当編集者用データ取得ハンドラー
 */
function getEditorManuscriptData(ssId, key) {
  const editorLog = findRecordByKey(ssId, EDITOR_LOG_SHEET_NAME, 'editorKey', key);
  if (!editorLog) return null;
  
  const msVer = editorLog.MsVer || '';
  const { msId, verNo } = parseMsVer(msVer);
  // 再投稿時に Manuscripts シートへ新行が追加されるため、Ver_No で正しいバージョン行を特定する
  const allMsRows = findAllRecordsByKey(ssId, MANUSCRIPTS_SHEET_NAME, 'MS_ID', msId);
  const fullMs = allMsRows.find(r => Number(r.Ver_No || 1) === verNo)
              || allMsRows[allMsRows.length - 1]
              || null;
  
  const msData = Object.assign({}, fullMs || {}, editorLog);
  
  // 査読者の承諾状況を確認
  const reviewLogs = findAllRecordsByKey(ssId, REVIEW_LOG_SHEET_NAME, 'MsVer', msVer);
  const acceptedReviewers = reviewLogs.filter(log => String(log.revOk || '').trim() === 'ok');
  const completedReviews = reviewLogs.filter(log => String(log.Received_At || '').trim() !== '');
  
  msData._reviewerAcceptedCount = acceptedReviewers.length;
  msData._acceptedReviewerNames = acceptedReviewers.map(log => log.Rev_Name || '').join(', ');
  msData._reviewCompletedCount = completedReviews.length;
  
  const fmtDate = createDateFormatter(ssId);
  msData._reviewerList = buildReviewerList(reviewLogs, fmtDate);

  // 過去ラウンドの査読者（ver.2+のみ）。担当編集者ダッシュボードで参考表示・再指名に使う。
  msData._pastReviewersByRound = [];
  if (verNo > 1) {
    const prevVersions = allMsRows
      .filter(r => Number(r.Ver_No || 1) < verNo)
      .sort((a, b) => Number(a.Ver_No || 1) - Number(b.Ver_No || 1));
    for (const prevVer of prevVersions) {
      const prevMsVer = String(prevVer.MsVer || '').trim() || (msId + '-' + Number(prevVer.Ver_No || 1));
      const prevReviewLogs = findAllRecordsByKey(ssId, REVIEW_LOG_SHEET_NAME, 'MsVer', prevMsVer);
      msData._pastReviewersByRound.push({
        round:     Number(prevVer.Ver_No || 1),
        msVer:     prevMsVer,
        reviewers: prevReviewLogs.map(log => ({
          Rev_Name:    String(log.Rev_Name    || '').trim(),
          Rev_Email:   String(log.Rev_Email   || '').trim(),
          revOk:       String(log.revOk       || '').trim(),
          Received_At: fmtDate(log.Received_At),
          Score:       String(log.Score       || '').trim()
        }))
      });
    }
  }

  // 全査読者が結果を提出済みかどうか
  const hasIncomplete = msData._reviewerList.some(r =>
    r.revOk === '' || (r.revOk === 'ok' && r.Received_At === '')
  );
  msData._allReviewsIn = (msData._reviewCompletedCount > 0) && !hasIncomplete;
  
  // 進捗ステータスの決定
  // 注: Manuscripts シートの最終判定列は小文字 'score' で保存されるため両方チェック
  const msFinalScore     = (fullMs || {}).score || (fullMs || {}).Score;
  const msSentBackAt     = (fullMs || {}).sentBackAt;
  const msEicDecision    = (fullMs || {}).eicFinalDecision;
  const msInProduction   = (fullMs || {}).finalStatus === 'in_production';
  const editorRecScore   = editorLog.Score;

  // 全査読が完了したかどうか（承諾済み査読者全員が結果を提出）
  const allReviewsCompleted = acceptedReviewers.length > 0 &&
                             acceptedReviewers.length === completedReviews.length;

  if (msInProduction) {
    msData._progressStatus = 'in_production';
  } else if (msFinalScore || msSentBackAt || msEicDecision) {
    msData._progressStatus = 'decision';
  } else if (editorRecScore || allReviewsCompleted) {
    msData._progressStatus = 'reviewed';
  } else if (acceptedReviewers.length > 0) {
    msData._progressStatus = 'reviewing';
  } else {
    msData._progressStatus = 'editor_assigned';
  }
  
  msData._scoreOptions = getScoreOptions();
  
  return msData;
}

/**
 * 査読者用データ取得ハンドラー
 */
function getReviewerManuscriptData(ssId, key) {
  const reviewLog = findRecordByKey(ssId, REVIEW_LOG_SHEET_NAME, 'reviewKey', key);
  
  if (!reviewLog) return null;
  
  const msVer = reviewLog.MsVer || '';
  const { msId, verNo } = parseMsVer(msVer);
  // 再投稿時に Manuscripts シートへ新行が追加されるため、Ver_No で正しいバージョン行を特定する
  const allMsRows = findAllRecordsByKey(ssId, MANUSCRIPTS_SHEET_NAME, 'MS_ID', msId);
  const fullMs = allMsRows.find(r => Number(r.Ver_No || 1) === verNo)
              || allMsRows[allMsRows.length - 1]
              || null;
  
  const msData = Object.assign({}, fullMs || {}, reviewLog);

  // 担当編集者のキーも取得して含める。
  // 同じ MsVer に複数の候補者がいる場合（1人目が辞退→2人目が承諾）、
  // 承諾済み(edtOk='ok')の編集者を優先し、いなければ最後の行をフォールバック。
  const editorLogs = findAllRecordsByKey(ssId, EDITOR_LOG_SHEET_NAME, 'MsVer', msVer);
  const acceptedEditorLog = editorLogs.find(log => String(log.edtOk || '').trim() === 'ok')
                         || editorLogs[editorLogs.length - 1]
                         || null;
  if (acceptedEditorLog) {
    msData.editorKey = acceptedEditorLog.editorKey;
  }

  // 査読者本人のアップロードフォルダURL（添付ファイル表示用）
  msData.reviewerFolderUrl = reviewLog.reviewerUploadFolderUrl || null;
  msData._scoreOptions = getScoreOptions();
  const allReviewLogs = findAllRecordsByKey(ssId, REVIEW_LOG_SHEET_NAME, 'MsVer', msVer);
  msData._progressStatus = determineProgressStatus(fullMs || {}, editorLogs, allReviewLogs);

  return msData;
}

/**
 * 編集幹事用データ取得ハンドラー
 * managingEditorKey で Manuscripts シートを検索し、原稿データと担当編集者スコアを返す
 */
function getManagingEditorManuscriptData(ssId, managingEditorKey) {
  const msData = findRecordByKey(ssId, MANUSCRIPTS_SHEET_NAME, 'managingEditorKey', managingEditorKey);
  if (!msData) return null;

  const msVer = msData.MsVer || '';
  const editorLogs = findAllRecordsByKey(ssId, EDITOR_LOG_SHEET_NAME, 'MsVer', msVer);
  const reviewLogs = findAllRecordsByKey(ssId, REVIEW_LOG_SHEET_NAME, 'MsVer', msVer);

  msData._progressStatus = determineProgressStatus(msData, editorLogs, reviewLogs);

  const fmt = createDateFormatter(ssId);

  // 担当編集者リスト（コメント・レポートファイル含む）
  msData._editorList = editorLogs.map(log => ({
    Editor_Name:         String(log.Editor_Name         || '').trim(),
    Editor_Email:        String(log.Editor_Email        || '').trim(),
    edtOk:               String(log.edtOk               || '').trim(),
    Ask_At:              fmt(log.Ask_At),
    Answer_At:           fmt(log.Answer_At),
    Score:               String(log.Score               || '').trim(),
    Message:             String(log.Message             || '').trim(),
    ConfidentialMessage: String(log.ConfidentialMessage || '').trim(),
    reportPdfUrl:        String(log.reportPdfUrl        || '').trim(),
    reportWordUrl:       String(log.reportWordUrl       || '').trim(),
    reportFolderUrl:     String(log.reportFolderUrl     || '').trim()
  }));

  // 担当編集者の推薦情報（後方互換のため残す）
  const acceptedEditor = editorLogs.find(log => String(log.edtOk || '').trim() === 'ok');
  msData._editorScore = acceptedEditor ? String(acceptedEditor.Score || '').trim() : '';
  msData._editorName  = acceptedEditor ? String(acceptedEditor.Editor_Name || '').trim() : '';

  // 査読者リスト（コメント含む）
  msData._reviewerList = buildReviewerList(reviewLogs, fmt);

  return msData;
}

/**
 * メインの getManuscriptData() 関数（リファクタリング後）
 * 'managing-editor' ロールは getManuscriptDataBatch を経由できないため直接処理する。
 * それ以外は getManuscriptDataBatch 経由でキャッシュを活用する。
 */
function getManuscriptDataRefactored(role, key) {
  if (!getSpreadsheetId()) return null;

  const r = (role || '').toLowerCase();

  // managing-editor は managingEditorKey 列で検索するため直接処理
  if (r === 'managing-editor') {
    try {
      const ssId = getSpreadsheetId();
      spreadsheetCache.prewarmSheets(ssId, [
        MANUSCRIPTS_SHEET_NAME,
        EDITOR_LOG_SHEET_NAME,
        REVIEW_LOG_SHEET_NAME,
        DECISION_MAIL_SHEET_NAME
      ]);
      let msData = getManagingEditorManuscriptData(ssId, key);
      if (!msData) return null;
      msData = convertDatesToStrings(msData);
      msData = enrichWithFolderInfo(msData, ssId);
      return msData;
    } catch (err) {
      console.error('managing-editor data fetch error:', err);
      return null;
    }
  }

  const results = getManuscriptDataBatch([{ role, key }]);
  return results[0] || null;
}