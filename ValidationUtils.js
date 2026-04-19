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
 * 危険な拡張子のブロックリスト
 * - 直接実行されうる実行ファイル（exe, com, bat, scr, msi 等）
 * - スクリプト系（vbs, ps1, hta, wsf 等）
 * - マクロ入り Office ファイル（docm, xlsm, pptm 等）
 * - ディスクイメージ（iso, dmg, img）
 *
 * 注意: 拡張子のみで判定しているため、攻撃者がファイル名を偽装すれば回避可能。
 * 完全な安全保証ではなく「うっかり危険ファイルを送ってしまう」事故の防止が目的。
 * クライアント側 (scripts.html) の DANGEROUS_FILE_EXTENSIONS_CLIENT と必ず一致させること。
 */
const DANGEROUS_FILE_EXTENSIONS = [
  // Windows 実行形式
  'exe', 'com', 'bat', 'cmd', 'msi', 'msp', 'mst', 'scr', 'pif',
  'dll', 'sys', 'drv', 'ocx', 'cpl', 'reg', 'msc', 'gadget', 'scf', 'lnk',
  // スクリプト
  'vbs', 'vbe', 'jse', 'wsf', 'wsh', 'ws', 'ps1', 'ps2', 'psc1', 'psc2', 'hta',
  // Mac / Linux 実行形式・パッケージ
  'app', 'dmg', 'pkg', 'deb', 'rpm', 'run',
  // Java / その他
  'jar', 'class',
  // ディスクイメージ・コンテナ
  'iso', 'img', 'vhd', 'vhdx',
  // マクロ入り Office (m 末尾)
  'docm', 'dotm', 'xlsm', 'xltm', 'xlsb', 'xlam', 'pptm', 'potm', 'ppam', 'sldm'
];

/**
 * 危険な拡張子の検証（ブロックリスト方式）
 * @param {Array<{name:string}>} files
 * @param {string} fieldName
 */
function validateFileSafety(files, fieldName) {
  if (!Array.isArray(files) || files.length === 0) return;
  const label = fieldName || '添付ファイル / Attachments';
  const blocked = [];
  for (let i = 0; i < files.length; i++) {
    const name = String((files[i] && files[i].name) || '');
    const dot = name.lastIndexOf('.');
    if (dot < 0 || dot === name.length - 1) continue;
    const ext = name.substring(dot + 1).toLowerCase().trim();
    if (DANGEROUS_FILE_EXTENSIONS.indexOf(ext) !== -1) {
      blocked.push(name);
    }
  }
  if (blocked.length > 0) {
    throw new Error(
      label + ' に危険な拡張子のファイルが含まれています: ' + blocked.join(', ') +
      '。実行ファイル（.exe など）やマクロ入り Office ファイル（.docm, .xlsm など）はアップロードできません。' +
      '原稿や図表として提出する場合はマクロを除去した形式（.docx, .xlsx など）に変換してください。 / ' +
      label + ' contains files with dangerous extensions: ' + blocked.join(', ') +
      '. Executable files (e.g. .exe) and macro-enabled Office files (e.g. .docm, .xlsm) cannot be uploaded. ' +
      'For manuscripts and figures, please save in macro-free formats (.docx, .xlsx, etc.).'
    );
  }
}

/**
 * 添付ファイル合計サイズの上限（30 MB）
 * - Gmail 添付の 25 MB 制限と GAS の google.script.run payload 上限（約 50 MB）の中間値
 * - クライアント側 (scripts.html) の MAX_ATTACHMENT_BYTES_CLIENT と必ず一致させること
 */
const MAX_ATTACHMENT_BYTES = 30 * 1024 * 1024;

/**
 * 添付ファイル合計サイズの検証（サーバ側の多重防御）
 * - クライアント側でも同様のチェックを行うが、DevTools 等でバイパスされうるため
 *   サーバ側にも同等のバリデーションを置く。
 * - files の content は Base64 文字列。length から実バイト長を逆算する。
 *
 * @param {Array<{name:string, content:string, mimeType:string}>} files
 * @param {number} maxTotalBytes
 * @param {string} fieldName
 */
function validateFileSize(files, maxTotalBytes, fieldName) {
  if (!Array.isArray(files) || files.length === 0) return;
  const label = fieldName || '添付ファイル / Attachments';
  let total = 0;
  for (let i = 0; i < files.length; i++) {
    const s = String((files[i] && files[i].content) || '');
    if (!s) continue;
    const padding = s.endsWith('==') ? 2 : (s.endsWith('=') ? 1 : 0);
    total += Math.floor(s.length * 3 / 4) - padding;
  }
  if (total > maxTotalBytes) {
    const totalMB = (total / 1024 / 1024).toFixed(1);
    const limitMB = (maxTotalBytes / 1024 / 1024).toFixed(0);
    throw new Error(
      label + ' の合計サイズが上限 (' + limitMB + ' MB) を超えています（現在: ' + totalMB + ' MB）。' +
      'ファイルサイズを調整するか、Google Drive にアップロードしてリンクを書いたファイルを代わりに添付してください。 / ' +
      'Total size of ' + label + ' exceeds the limit (' + limitMB + ' MB) — currently ' + totalMB + ' MB. ' +
      'Please reduce file sizes or upload figures to Google Drive and attach a link file instead.'
    );
  }
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
