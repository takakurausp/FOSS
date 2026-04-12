/**
 * CodeOptimized.js - パフォーマンス最適化版の主要関数
 */

/**
 * 最適化版 getSpreadsheetId - キャッシュ付き
 */
let spreadsheetIdCache = null;
let spreadsheetIdTimestamp = null;
const SPREADSHEET_ID_CACHE_EXPIRATION = 600; // 10分

function getSpreadsheetIdOptimized() {
  const now = new Date().getTime();
  
  // キャッシュチェック
  if (spreadsheetIdCache && spreadsheetIdTimestamp) {
    const age = (now - spreadsheetIdTimestamp) / 1000;
    if (age < SPREADSHEET_ID_CACHE_EXPIRATION) {
      return spreadsheetIdCache;
    }
  }
  
  // スクリプトプロパティから取得
  const propId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (propId) {
    spreadsheetIdCache = propId;
    spreadsheetIdTimestamp = now;
    return propId;
  }
  
  // フォールバック: fassSetting.json
  try {
    const files = DriveApp.getFilesByName("fassSetting.json");
    if (files.hasNext()) {
      const content = files.next().getBlob().getDataAsString();
      const id = JSON.parse(content).spreadsheetId;
      spreadsheetIdCache = id;
      spreadsheetIdTimestamp = now;
      return id;
    }
  } catch (e) {
    Logger.log('Error reading fassSetting.json: ' + e);
  }
  
  throw new Error('スプレッドシートIDが見つかりません。スクリプトプロパティに SPREADSHEET_ID を設定してください。');
}


/**
 * 最適化版 updateLogCell - バッチ更新対応
 */
function updateLogCellBatch(ssId, sheetName, updatesArray) {
  if (!updatesArray || updatesArray.length === 0) return;
  
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(sheetName);
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // ヘッダーインデックスのマップを作成
  const headerIndexMap = {};
  headers.forEach((h, i) => {
    headerIndexMap[String(h).toLowerCase().trim()] = i;
  });
  
  // 更新対象の行を特定
  const updatesByRow = {};
  
  updatesArray.forEach(update => {
    const { keyColName, keyValue, updates } = update;
    const keyColIndex = headerIndexMap[keyColName.toLowerCase().trim()];
    if (keyColIndex === undefined) return;
    
    // キー値に一致する行を検索
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][keyColIndex]).trim() === String(keyValue).trim()) {
        if (!updatesByRow[i]) {
          updatesByRow[i] = {};
        }
        
        Object.keys(updates).forEach(colName => {
          const colIndex = headerIndexMap[colName.toLowerCase().trim()];
          if (colIndex !== undefined) {
            updatesByRow[i][colIndex] = updates[colName];
          }
        });
        break;
      }
    }
  });
  
  // バッチ更新
  Object.keys(updatesByRow).forEach(rowIndexStr => {
    const rowIndex = parseInt(rowIndexStr);
    const rowUpdates = updatesByRow[rowIndex];
    
    Object.keys(rowUpdates).forEach(colIndexStr => {
      const colIndex = parseInt(colIndexStr);
      sheet.getRange(rowIndex + 1, colIndex + 1).setValue(rowUpdates[colIndex]);
    });
  });
  
  SpreadsheetApp.flush();
  spreadsheetCache.invalidate(ssId, sheetName);
}

/**
 * 最適化版 getManuscriptData - バッチ処理対応
 *
 * 全リクエストに必要なシートを Spreadsheet を1回だけ開いて spreadsheetCache に展開し、
 * 後続のロールハンドラー内の findRecordByKey がすべてキャッシュを参照するようにする。
 * これにより N 件のリクエストでも Spreadsheet.openById は1回、
 * getValues は対象シート数分（最大3回）で済む。
 */
function getManuscriptDataBatch(requests) {
  if (!requests || requests.length === 0) return [];

  const ssId = getSpreadsheetIdOptimized();

  // 全リクエストに必要なシートをまとめてキャッシュに展開（SS を1回だけ開く）
  spreadsheetCache.prewarmSheets(ssId, [
    MANUSCRIPTS_SHEET_NAME,
    EDITOR_LOG_SHEET_NAME,
    REVIEW_LOG_SHEET_NAME
  ]);

  // 各リクエストを処理（ロールハンドラーは warm キャッシュを参照するため API 呼び出しなし）
  return requests.map((req, index) => {
    try {
      const r = (req.role || '').toLowerCase();
      let msData = null;

      switch (r) {
        case 'author':          msData = getAuthorManuscriptData(ssId, req.key);          break;
        case 'eic':             msData = getEicManuscriptData(ssId, req.key);             break;
        case 'editor':          msData = getEditorManuscriptData(ssId, req.key);          break;
        case 'reviewer':        msData = getReviewerManuscriptData(ssId, req.key);        break;
        case 'managing-editor': msData = getManagingEditorManuscriptData(ssId, req.key); break;
        default: return null;
      }

      if (!msData) return null;
      msData = convertDatesToStrings(msData);
      msData = enrichWithFolderInfo(msData, ssId);
      return msData;
    } catch (e) {
      Logger.log(`getManuscriptDataBatch[${index}] error: ${e}`);
      return null;
    }
  });
}


/**
 * 汎用バッチ処理ユーティリティ
 */
function executeInBatches(items, batchSize, processor) {
  const results = [];
  for (let i = 0; i < items.length; i += batchSize) {
    const batch = items.slice(i, i + batchSize);
    results.push(...processor(batch));
  }
  return results;
}

/**
 * メール送信のバッチ処理
 *
 * sleep によるブロッキングを廃止し、sendEmailSafe に委譲する。
 * クォータ超過・送信エラー時は sendEmailSafe が Emails シートに保存し、
 * 時間ベーストリガー（retrySendingEmails）が後続で再送する。
 */
function sendEmailsBatch(emailOptionsArray) {
  if (!emailOptionsArray || emailOptionsArray.length === 0) return [];

  return emailOptionsArray.map((options, index) => {
    try {
      sendEmailSafe(options, `Batch email ${index}`);
      return { success: true };
    } catch (e) {
      Logger.log(`sendEmailsBatch[${index}] unexpected error: ${e.message}`);
      return { success: false, error: e.message };
    }
  });
}

// バッチAPIエンドポイント
function apiGetManuscriptDataBatch(data) {
  try {
    const requests = JSON.parse(data.requests || '[]');
    const results = getManuscriptDataBatch(requests);
    return { success: true, results: results };
  } catch (e) {
    Logger.log(`Error in batch API: ${e}`);
    return { success: false, error: e.message };
  }
}