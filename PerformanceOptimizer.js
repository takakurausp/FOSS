/**
 * PerformanceOptimizer.js - GASアプリのパフォーマンス最適化モジュール
 */

// キャッシュの有効期限（秒）
const CACHE_EXPIRATION = 300; // 5分

/**
 * スプレッドシートのヘッダー名を正規化するマップ。
 * シートによって大文字小文字が揺れる列名をコード側の正規名に統一する。
 * キーは小文字化したヘッダー名、値が正規名。
 */
const CANONICAL_COLUMN_NAMES = {
  // Editor_log
  'editorkey':    'editorKey',   // EditorKey, editorkey → editorKey（既存）
  'msver':        'MsVer',       // msver → MsVer
  // Review_log
  'revok':        'revOk',       // revok, RevOk → revOk
  'review_email': 'Rev_Email',   // Review_Email → Rev_Email（Rev_Email に統一）
  // Manuscripts
  'ver_no':       'Ver_No',      // ver_no, VER_NO → Ver_No
  'verno':        'Ver_No',      // verNo（camelCase 変異）→ Ver_No
  'sentbackat':   'sentBackAt',  // SentBackAt → sentBackAt（書き込みに合わせ camelCase に統一）
};

/**
 * スプレッドシートデータのキャッシュ管理クラス
 */
class SpreadsheetCache {
  constructor() {
    this.cache = {};
    this.timestamps = {};
  }
  
  /**
   * シートデータをキャッシュから取得または読み込み
   */
  getSheetData(ssId, sheetName) {
    const cacheKey = `${ssId}_${sheetName}`;
    const now = new Date().getTime();
    
    // キャッシュが有効なら返す
    if (this.cache[cacheKey] && this.timestamps[cacheKey]) {
      const age = (now - this.timestamps[cacheKey]) / 1000;
      if (age < CACHE_EXPIRATION) {
        return this.cache[cacheKey];
      }
    }
    
    // キャッシュが無効または存在しない場合は読み込み
    try {
      const sheet = SpreadsheetApp.openById(ssId).getSheetByName(sheetName);
      if (!sheet) {
        this.cache[cacheKey] = null;
        this.timestamps[cacheKey] = now;
        return null;
      }
      
      const data = sheet.getDataRange().getValues();
      this.cache[cacheKey] = {
        headers: data[0],
        rows: data.slice(1),
        lastRow: sheet.getLastRow(),
        lastColumn: sheet.getLastColumn()
      };
      this.timestamps[cacheKey] = now;
      
      return this.cache[cacheKey];
    } catch (e) {
      Logger.log(`Error loading sheet ${sheetName}: ${e}`);
      return null;
    }
  }
  
  /**
   * 複数シートをまとめてキャッシュに展開（Spreadsheet を1回だけ開く）
   * getManuscriptDataBatch から呼ばれ、後続の findRecordByKey がキャッシュを使えるようにする
   */
  prewarmSheets(ssId, sheetNames) {
    const now = new Date().getTime();
    const ss = SpreadsheetApp.openById(ssId);
    for (const sheetName of sheetNames) {
      const cacheKey = `${ssId}_${sheetName}`;
      if (this.cache[cacheKey] && this.timestamps[cacheKey]) {
        const age = (now - this.timestamps[cacheKey]) / 1000;
        if (age < CACHE_EXPIRATION) continue; // すでに有効なキャッシュがある
      }
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) continue;
      const data = sheet.getDataRange().getValues();
      this.cache[cacheKey] = {
        headers: data[0],
        rows: data.slice(1),
        lastRow: sheet.getLastRow(),
        lastColumn: sheet.getLastColumn()
      };
      this.timestamps[cacheKey] = now;
    }
  }

  /**
   * 特定のキャッシュを無効化
   */
  invalidate(ssId, sheetName) {
    const cacheKey = `${ssId}_${sheetName}`;
    delete this.cache[cacheKey];
    delete this.timestamps[cacheKey];
  }
  
  /**
   * 全キャッシュをクリア
   */
  clearAll() {
    this.cache = {};
    this.timestamps = {};
  }
}

/**
 * 設定データのキャッシュ管理
 */
class SettingsCache {
  constructor() {
    this.settingsCache = null;
    this.settingsTimestamp = null;
  }
  
  getSettings(ssId) {
    const now = new Date().getTime();
    
    if (this.settingsCache && this.settingsTimestamp) {
      const age = (now - this.settingsTimestamp) / 1000;
      if (age < CACHE_EXPIRATION) {
        return this.settingsCache;
      }
    }
    
    try {
      const sheet = SpreadsheetApp.openById(ssId).getSheetByName('Settings');
      if (!sheet) {
        this.settingsCache = {};
        this.settingsTimestamp = now;
        return {};
      }
      
      const data = sheet.getRange('A5:B100').getValues();
      const settings = {};
      data.forEach(row => {
        if (row[0]) {
          settings[row[0]] = row[1];
        }
      });
      
      this.settingsCache = settings;
      this.settingsTimestamp = now;
      
      return settings;
    } catch (e) {
      Logger.log(`Error loading settings: ${e}`);
      return {};
    }
  }
  
  invalidate() {
    this.settingsCache = null;
    this.settingsTimestamp = null;
  }
}

// グローバルキャッシュインスタンス
const spreadsheetCache = new SpreadsheetCache();
const settingsCache = new SettingsCache();

/**
 * 最適化版 findAllRecordsByKey
 * キャッシュを使用し、一度のAPI呼び出しで複数検索を処理
 */
function findAllRecordsByKeyOptimized(ssId, sheetName, keyColName, keyValue) {
  const cachedData = spreadsheetCache.getSheetData(ssId, sheetName);
  if (!cachedData) return [];
  
  const { headers, rows } = cachedData;
  const lowerKeyColName = keyColName.toLowerCase().trim();
  
  // ヘッダーインデックスの検索（キャッシュから取得可能ならさらに最適化）
  let keyColIndex = -1;
  for (let j = 0; j < headers.length; j++) {
    if (String(headers[j]).toLowerCase().trim() === lowerKeyColName) {
      keyColIndex = j;
      break;
    }
  }
  
  if (keyColIndex === -1) return [];
  
  const results = [];
  const searchValue = String(keyValue).trim();
  
  // 最適化されたループ
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (String(row[keyColIndex]).trim() === searchValue) {
      const record = {};
      for (let j = 0; j < headers.length; j++) {
        const canonical = CANONICAL_COLUMN_NAMES[String(headers[j]).toLowerCase().trim()] || headers[j];
        record[canonical] = row[j];
      }
      results.push(record);
    }
  }

  return results;
}

/**
 * 最適化版 findRecordByKey
 * 最初に見つかったレコードのみを返す
 */
function findRecordByKeyOptimized(ssId, sheetName, keyColName, keyValue) {
  const cachedData = spreadsheetCache.getSheetData(ssId, sheetName);
  if (!cachedData) return null;
  
  const { headers, rows } = cachedData;
  const lowerKeyColName = keyColName.toLowerCase().trim();
  
  let keyColIndex = -1;
  for (let j = 0; j < headers.length; j++) {
    if (String(headers[j]).toLowerCase().trim() === lowerKeyColName) {
      keyColIndex = j;
      break;
    }
  }
  
  if (keyColIndex === -1) return null;
  
  const searchValue = String(keyValue).trim();
  
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (String(row[keyColIndex]).trim() === searchValue) {
      const record = {};
      for (let j = 0; j < headers.length; j++) {
        const canonical = CANONICAL_COLUMN_NAMES[String(headers[j]).toLowerCase().trim()] || headers[j];
        record[canonical] = row[j];
      }
      return record;
    }
  }

  return null;
}

/**
 * バッチ検索 - 複数のキー値を一度に検索
 */
function findRecordsByKeysBatch(ssId, sheetName, keyColName, keyValues) {
  const cachedData = spreadsheetCache.getSheetData(ssId, sheetName);
  if (!cachedData) return {};
  
  const { headers, rows } = cachedData;
  const lowerKeyColName = keyColName.toLowerCase().trim();
  
  let keyColIndex = -1;
  for (let j = 0; j < headers.length; j++) {
    if (String(headers[j]).toLowerCase().trim() === lowerKeyColName) {
      keyColIndex = j;
      break;
    }
  }
  
  if (keyColIndex === -1) return {};
  
  const results = {};
  const searchValuesSet = new Set(keyValues.map(v => String(v).trim()));
  
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const rowValue = String(row[keyColIndex]).trim();
    
    if (searchValuesSet.has(rowValue)) {
      const record = {};
      for (let j = 0; j < headers.length; j++) {
        record[headers[j]] = row[j];
      }
      
      if (!results[rowValue]) {
        results[rowValue] = [];
      }
      results[rowValue].push(record);
    }
  }
  
  return results;
}

/**
 * 最適化版 getSettings
 */
function getSettingsOptimized(ssId) {
  return settingsCache.getSettings(ssId);
}

/**
 * Driveフォルダ検索のキャッシュ
 */
class DriveFolderCache {
  constructor() {
    this.folderCache = {};
  }
  
  getRootFolder(rootName) {
    const cacheKey = `__root__${rootName}`;
    if (Object.prototype.hasOwnProperty.call(this.folderCache, cacheKey)) {
      return this.folderCache[cacheKey];
    }
    const folders = DriveApp.getFoldersByName(rootName);
    const folder = folders.hasNext() ? folders.next() : null;
    this.folderCache[cacheKey] = folder;
    return folder;
  }

  getFolderByName(parentFolder, folderName) {
    const cacheKey = `${parentFolder.getId()}_${folderName}`;
    if (Object.prototype.hasOwnProperty.call(this.folderCache, cacheKey)) {
      return this.folderCache[cacheKey];
    }
    const folders = parentFolder.getFoldersByName(folderName);
    const folder = folders.hasNext() ? folders.next() : null;
    this.folderCache[cacheKey] = folder;
    return folder;
  }

  getOrCreateFolder(parentFolder, folderName) {
    let folder = this.getFolderByName(parentFolder, folderName);
    if (!folder) {
      folder = parentFolder.createFolder(folderName);
      this.folderCache[`${parentFolder.getId()}_${folderName}`] = folder;
    }
    return folder;
  }

  clear() {
    this.folderCache = {};
  }
}

const driveFolderCache = new DriveFolderCache();

/**
 * フォルダパスを一度のAPI呼び出しで解決
 */
function resolveFolderPathOptimized(settings, msData) {
  try {
    const rootName = settings.SUBFOLDER || 'Journal Files';
    const rootFolders = DriveApp.getFoldersByName(rootName);
    if (!rootFolders.hasNext()) return null;
    
    const root = rootFolders.next();
    
    // MS_IDフォルダ
    const msFolder = driveFolderCache.getFolderByName(root, msData.MS_ID);
    if (!msFolder) return null;
    
    const verNo = msData.Ver_No || 1;
    const verFolderName = `ver.${verNo}`;
    const verFolder = driveFolderCache.getFolderByName(msFolder, verFolderName);
    if (!verFolder) return null;
    
    const submittedFolder = driveFolderCache.getFolderByName(verFolder, 'submitted');
    if (!submittedFolder) return null;
    
    // 共有設定
    submittedFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return {
      folder: submittedFolder,
      url: submittedFolder.getUrl()
    };
  } catch (e) {
    Logger.log('Error resolving folder path: ' + e);
    return null;
  }
}

/**
 * パフォーマンス計測ユーティリティ
 */
class PerformanceMonitor {
  constructor() {
    this.timings = {};
    this.startTimes = {};
  }
  
  start(operationName) {
    this.startTimes[operationName] = new Date().getTime();
  }
  
  end(operationName) {
    const endTime = new Date().getTime();
    const startTime = this.startTimes[operationName];
    
    if (startTime) {
      const duration = endTime - startTime;
      if (!this.timings[operationName]) {
        this.timings[operationName] = [];
      }
      this.timings[operationName].push(duration);
      
      // ログ出力（開発時のみ）
      if (duration > 1000) { // 1秒以上かかった操作
        Logger.log(`SLOW OPERATION: ${operationName} took ${duration}ms`);
      }
      
      delete this.startTimes[operationName];
    }
  }
  
  getReport() {
    const report = {};
    for (const [operation, timings] of Object.entries(this.timings)) {
      if (timings.length > 0) {
        const avg = timings.reduce((a, b) => a + b, 0) / timings.length;
        const max = Math.max(...timings);
        const min = Math.min(...timings);
        report[operation] = {
          count: timings.length,
          avg: Math.round(avg),
          min: min,
          max: max,
          total: timings.reduce((a, b) => a + b, 0)
        };
      }
    }
    return report;
  }
}

// グローバルパフォーマンスモニター
const performanceMonitor = new PerformanceMonitor();

/**
 * パフォーマンス計測ラッパー
 */
function withPerformanceMonitoring(fn, operationName) {
  return function(...args) {
    performanceMonitor.start(operationName);
    try {
      return fn.apply(this, args);
    } finally {
      performanceMonitor.end(operationName);
    }
  };
}

// 主要関数のパフォーマンス計測を有効化
const monitoredFunctions = {
  getManuscriptDataRefactored: withPerformanceMonitoring(getManuscriptDataRefactored, 'getManuscriptData'),
  findAllRecordsByKey: withPerformanceMonitoring(findAllRecordsByKeyOptimized, 'findAllRecordsByKey'),
  findRecordByKey: withPerformanceMonitoring(findRecordByKeyOptimized, 'findRecordByKey'),
  getSettings: withPerformanceMonitoring(getSettingsOptimized, 'getSettings')
};

// エクスポート
if (typeof module !== 'undefined') {
  module.exports = {
    SpreadsheetCache,
    SettingsCache,
    DriveFolderCache,
    PerformanceMonitor,
    findAllRecordsByKeyOptimized,
    findRecordByKeyOptimized,
    findRecordsByKeysBatch,
    getSettingsOptimized,
    resolveFolderPathOptimized,
    performanceMonitor,
    withPerformanceMonitoring,
    monitoredFunctions
  };
}