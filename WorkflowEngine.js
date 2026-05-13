/**
 * WorkflowEngine.js - API 関数の共通パターンを抽出したモジュール
 *
 * 全 apiXxx() 関数に共通するボイラープレートを一元化し、
 * 各モジュールはビジネスロジックのみを記述するようにする。
 */

/**
 * アップロードファイルの共通バリデーション。
 * ファイル配列が null/empty の場合は何もしない。
 * @param {Array} files - { name, mimeType, content } の配列（null 可）
 */
function validateUploadedFiles(files) {
  if (!files || files.length === 0) return;
  files.forEach(function(file, i) { validateFileName(file.name, 'ファイル名 ' + (i + 1)); });
  validateFileSafety(files, '添付ファイル / Attachments');
  validateFileSize(files, MAX_ATTACHMENT_BYTES, '添付ファイル / Attachments');
}

/**
 * API 関数の共通コンテキストを取得する。
 * getManuscriptData() でデータを検索し、見つからなければ例外を投げる。
 *
 * @param {string} role   - getManuscriptData のロール ('editor','reviewer','managing-editor','eic','author')
 * @param {string} key    - 検索キー
 * @param {string} [label] - エラーメッセージ用の表示名
 * @returns {{ ssId: string, settings: Object, msData: Object }}
 */
function getApiContext(role, key, label) {
  enforceRateLimit(key);
  var ssId = getSpreadsheetId();
  var settings = getSettings();
  var msData = getManuscriptData(role, key);
  if (!msData) {
    throw new Error((label || 'Record') + ' not found.');
  }
  return { ssId: ssId, settings: settings, msData: msData };
}

/**
 * API 関数の共通コンテキストを取得する（Manuscripts.key / eicKey で検索する版）。
 * FeedbackModule, EditorAssignment で使用。
 *
 * @param {string} key    - 検索キー（key または eicKey）
 * @param {string} [label] - エラーメッセージ用の表示名
 * @returns {{ ssId: string, settings: Object, msData: Object }}
 */
function getApiContextByMsKey(key, label) {
  enforceRateLimit(key);
  var ssId = getSpreadsheetId();
  var settings = getSettings();
  var ms = findRecordByKey(ssId, MANUSCRIPTS_SHEET_NAME, 'key', key);
  if (!ms) ms = findRecordByKey(ssId, MANUSCRIPTS_SHEET_NAME, 'eicKey', key);
  if (!ms) {
    throw new Error((label || 'Manuscript') + ' not found for key: ' + key);
  }
  return { ssId: ssId, settings: settings, msData: ms };
}

/**
 * API 用のタイムスタンプ文字列を生成（Asia/Tokyo 固定）
 * @returns {string} "yyyy/MM/dd HH:mm" 形式
 */
function apiTimestamp() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
}

/**
 * レート制限の最小間隔（秒）
 */
var RATE_LIMIT_SECONDS = 5;

/**
 * 指定したキーに対するレート制限をチェックする。
 * 前回のリクエストから RATE_LIMIT_SECONDS 秒以内の場合は例外を投げる。
 * ScriptProperties にタイムスタンプを保存するため、
 * 実行間隔は全並行実行で共有される。
 *
 * @param {string} id - ユーザー識別子（editorKey, reviewKey, authorEmail など）
 */
function enforceRateLimit(id) {
  if (!id) return;
  var props = PropertiesService.getScriptProperties();
  var propKey = 'rl_' + id;
  var now = Date.now();
  var last = Number(props.getProperty(propKey)) || 0;
  if (now - last < RATE_LIMIT_SECONDS * 1000) {
    throw new Error('リクエストが早すぎます。少しお待ちください。/ Too many requests. Please wait a moment.');
  }
  props.setProperty(propKey, String(now));
}
