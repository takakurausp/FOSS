/**
 * ValidationUtils.js - システム境界での入力バリデーション
 *
 * GAS は全ファイルが同一グローバルスコープで実行されるため、
 * ここで定義した関数はすべてのモジュールから直接呼び出せる。
 */

/**
 * メールアドレスの検証
 * - 必須チェック
 * - 改行文字の禁止（メールヘッダーインジェクション対策）
 * - 基本的なフォーマット確認
 * @param {string} email
 * @param {string} fieldName  エラーメッセージに表示するフィールド名
 * @returns {string} trim済みの検証済みメールアドレス
 */
function validateEmail(email, fieldName) {
  const label = fieldName || 'Email';
  if (email === null || email === undefined || typeof email !== 'string') {
    throw new Error(label + ' は必須です。/ ' + label + ' is required.');
  }
  const trimmed = email.trim();
  if (trimmed.length === 0) {
    throw new Error(label + ' は必須です。/ ' + label + ' is required.');
  }
  if (/[\r\n]/.test(trimmed)) {
    throw new Error(label + ' に不正な文字が含まれています。/ ' + label + ' contains invalid characters.');
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(trimmed)) {
    throw new Error(label + ' の形式が正しくありません: ' + trimmed + ' / ' + label + ' format is invalid: ' + trimmed);
  }
  return trimmed;
}

/**
 * カンマ区切りのメールアドレスリストの検証（CC欄など）
 * 空文字列の場合はスキップする。
 * @param {string} emailList  カンマ区切りのメールアドレス
 * @param {string} fieldName
 * @returns {string} trim済みの検証済み文字列（元の形式を保持）
 */
function validateEmailList(emailList, fieldName) {
  if (!emailList || String(emailList).trim() === '') return '';
  const label = fieldName || 'Email list';
  String(emailList).split(',').forEach((email, i) => {
    const trimmed = email.trim();
    if (trimmed) validateEmail(trimmed, label + ' #' + (i + 1));
  });
  return String(emailList).trim();
}

/**
 * 必須文字列の検証
 * - 必須チェック
 * - 改行文字の禁止（単行フィールド：名前・キー類）
 * @param {*}      val
 * @param {string} fieldName
 * @returns {string} trim済みの検証済み文字列
 */
function validateRequiredString(val, fieldName) {
  const label = fieldName || 'Field';
  if (val === null || val === undefined || typeof val !== 'string') {
    throw new Error(label + ' は必須です。/ ' + label + ' is required.');
  }
  const trimmed = val.trim();
  if (trimmed.length === 0) {
    throw new Error(label + ' は必須です。/ ' + label + ' is required.');
  }
  if (/[\r\n]/.test(trimmed)) {
    throw new Error(label + ' に改行を含めることはできません。/ ' + label + ' must not contain line breaks.');
  }
  return trimmed;
}

/**
 * ファイル名の検証
 * - パストラバーサル（../ など）の禁止
 * - 絶対パスの禁止
 * - nullバイト・制御文字の禁止（Drive API クラッシュ防止）
 * - ファイル名の長さ制限（255文字）
 * @param {string} name
 * @param {string} fieldName
 * @returns {string} trim済みの検証済みファイル名
 */
function validateFileName(name, fieldName) {
  const label = fieldName || 'File name';
  if (!name || typeof name !== 'string') {
    throw new Error(label + ' は必須です。/ ' + label + ' is required.');
  }
  const trimmed = name.trim();
  if (trimmed.length === 0) {
    throw new Error(label + ' は必須です。/ ' + label + ' is required.');
  }
  if (trimmed.length > 255) {
    throw new Error(label + ' が長すぎます（最大255文字）。/ ' + label + ' is too long (max 255 characters).');
  }
  // eslint-disable-next-line no-control-regex
  if (/[\x00-\x1f\x7f]/.test(trimmed)) {
    throw new Error(label + ' に制御文字が含まれています。/ ' + label + ' contains control characters.');
  }
  if (/\.\.[\\/]|^[\\/]|^\.\.?$/.test(trimmed)) {
    throw new Error(label + ' に不正なパス文字が含まれています。/ ' + label + ' contains invalid path characters.');
  }
  return trimmed;
}

/**
 * HTML特殊文字のエスケープ（XSS対策）
 * メール本文など HTML 文字列にユーザー入力を埋め込む際に使用する。
 * @param {*} s
 * @returns {string}
 */
function escHtml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * MsVer 文字列（例: "JJEEZ3-1"、"JJEEZ3-2"）を MS_ID と verNo に分解する。
 *
 * MsVer のフォーマットは "{MS_ID}-{verNo}" であり、verNo は末尾の数値セグメント。
 * lastIndexOf('-') を使うことで MS_ID 自体にハイフンが含まれる場合も正しく動作する。
 *
 * @param {string} msVer
 * @returns {{ msId: string, verNo: number }}
 */
function parseMsVer(msVer) {
  const str = String(msVer || '').trim();
  const lastHyphen = str.lastIndexOf('-');
  if (lastHyphen === -1) return { msId: str, verNo: 1 };
  return {
    msId:  str.substring(0, lastHyphen),
    verNo: parseInt(str.substring(lastHyphen + 1), 10) || 1
  };
}
