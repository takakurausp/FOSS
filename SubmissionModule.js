/**
 * SubmissionModule.gs - 論文投稿処理ロジック（修正版）
 */

/**
 * クライアントからの新規投稿リクエストを処理
 */
function apiSubmitSubmission(data) {
  // 入力バリデーション
  validateRequiredString(data.authorName,  '著者名 (authorName)');
  validateEmail(data.authorEmail,          '著者メールアドレス (authorEmail)');
  validateRequiredString(data.paperType,   '論文種別 (paperType)');
  if (!String(data.titleJp || '').trim() && !String(data.titleEn || '').trim()) {
    throw new Error('タイトル（日本語または英語）は最低1つ必須です。/ At least one of titleJp or titleEn is required.');
  }
  if (data.ccEmails) validateEmailList(data.ccEmails, 'CC メールアドレス (ccEmails)');
  if (data.files && data.files.length > 0) {
    data.files.forEach((file, i) => validateFileName(file.name, 'ファイル名 ' + (i + 1)));
  }

  const ssId = getSpreadsheetId();
  const settings = getSettings();
  
  // 1. IDの発行
  const lastId = findMaxValue('ID', ssId);
  const nowId = lastId + 1;
  const prefix = getPrefix(data.paperType, ssId);
  const msId = prefix + nowId;
  const msVer = msId + '-1';
  const verNo = 1;
  const msVerHex = getMsVerHEX(nowId, 1);
  const key = msVerHex + Utilities.getUuid().replace(/-/g, '');         // 著者用キー — CSPRNG (128 bit)
  const eicKey = 'E' + Utilities.getUuid().replace(/-/g, '');          // 編集委員長専用キー — CSPRNG (128 bit)

  const thisManuscript = {
    nowId: nowId,
    MS_ID: msId,
    MsVer: msVer,
    verNo: verNo,
    MsVerHex: msVerHex,
    key: key,
    eicKey: eicKey,
    authorName: data.authorName,
    authorEmail: data.authorEmail,
    authorAffiliation: data.authorAffiliation || '',
    authorAddress: data.authorAddress || '',
    authorsJp: data.authorsJp || '',
    authorsEn: data.authorsEn || '',
    ccEmails: data.ccEmails || '',
    paperType: data.paperType,
    titleJp: data.titleJp || '',
    titleEn: data.titleEn || '',
    runningTitle: data.runningTitle || '',
    abstractJp: data.abstractJp || '',
    abstractEn: data.abstractEn || '',
    letterToEditor: data.letterToEditor || '',
    englishEditing: data.englishEditing || '',
    reprintRequest: data.reprintRequest || '',
    sendDateTime: Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss'),
    submittedFiles: ''
  };

  // 2. フォルダ作成とファイル保存
  // submitted フォルダ: 投稿原稿を保存（担当編集者・査読者に共有）
  const folder = createSubmissionFolder(settings, thisManuscript);
  const fileNames = [];
  if (data.files && data.files.length > 0) {
    data.files.forEach(file => {
      try {
        const blob = Utilities.newBlob(
          Utilities.base64Decode(file.content),
          file.mimeType || 'application/octet-stream',
          file.name
        );
        folder.createFile(blob);
        fileNames.push(file.name);
        Logger.log(`Saved file to submitted folder: ${file.name}`);
      } catch (fileErr) {
        Logger.log(`Failed to save file "${file.name}": ${fileErr}`);
      }
    });
  } else {
    Logger.log(`New Submission ${msId}: data.files is empty or undefined`);
  }
  thisManuscript.submittedFiles = fileNames.join(', ');

  // submitted フォルダを共有設定（リンクを知っている全員が閲覧可）
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  thisManuscript.folderUrl = folder.getUrl();

  // receipt フォルダ: 受領票PDFを保存（編集委員長・印刷担当者のみに共有、担当編集者・査読者には不開示）
  const receiptFolder = createReceiptFolder(settings, thisManuscript);
  receiptFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  thisManuscript.receiptFolderUrl = receiptFolder.getUrl();

  // 3. データベースへの登録（submittedFiles / folderUrl / receiptFolderUrl 設定後に実行）
  addNewMsToDB(ssId, thisManuscript);

  writeLog(`New Submission: ${msId} by ${data.authorEmail}`);

  // 4. 受領通知とEICへの連絡（各メール送信は独立して実行）
  const emailWarnings = [];

  try {
    const receiptBlob = sendReceiptEmail(ssId, thisManuscript);
    // 受領票PDFを receipt フォルダに保存（submitted フォルダには保存しない）
    if (receiptBlob) {
      try { receiptFolder.createFile(receiptBlob); } catch (e) { Logger.log('Receipt PDF save error: ' + e); }
    }
  } catch (err) {
    Logger.log('sendReceiptEmail error: ' + err);
    emailWarnings.push('receipt');
  }

  try {
    sendEicNotification(thisManuscript);
  } catch (err) {
    Logger.log('sendEicNotification error: ' + err);
    emailWarnings.push('eic');
  }

  return {
    success: true,
    msId: msId,
    msVer: msVer,
    key: key,
    emailWarning: emailWarnings.length > 0 ? emailWarnings : null
  };
}

/**
 * IDの最大値を取得 (Manuscriptsシート)
 */
function findMaxValue(colName, ssId) {
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(MANUSCRIPTS_SHEET_NAME);
  if (!sheet) return 0;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = headers.findIndex(h => String(h).toLowerCase() === colName.toLowerCase());
  if (colIndex === -1) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const values = sheet.getRange(2, colIndex + 1, lastRow - 1).getValues();
  const numbers = values.flat().map(v => Number(v)).filter(v => !isNaN(v) && v > 0);
  return numbers.length > 0 ? Math.max(...numbers) : 0;
}

/**
 * 接頭辞を取得 (Settingsシート E5:F18)
 */
function getPrefix(paperType, ssId) {
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(SETTINGS_SHEET_NAME);
  const data = sheet.getRange('E5:F18').getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === paperType) return data[i][1];
  }
  return '';
}

/**
 * Manuscriptsシートに新規レコード追加
 */
function addNewMsToDB(ssId, ms) {
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(MANUSCRIPTS_SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newRow = new Array(headers.length).fill('');
  
  const mapping = {
    'ID': ms.nowId,
    'MS_ID': ms.MS_ID,
    'MsVer': ms.MsVer,
    'MsVerHex': ms.MsVerHex,
    'key': ms.key,
    'eicKey': ms.eicKey,
    'CA_Name': ms.authorName,
    'CA_Email': ms.authorEmail,
    'CA_Affiliation': ms.authorAffiliation,
    'Address': ms.authorAddress,
    'AuthorsJP': ms.authorsJp,
    'AuthorsEN': ms.authorsEn,
    'ccEmails': ms.ccEmails,
    'MS_Type': ms.paperType,
    'TitleJP': ms.titleJp,
    'TitleEN': ms.titleEn,
    'RunningTitle': ms.runningTitle,
    'AbstractJP': ms.abstractJp,
    'AbstractEN': ms.abstractEn,
    'LetterToEditor': ms.letterToEditor,
    'English_editing': ms.englishEditing,
    'Reprint request': ms.reprintRequest,
    'submittedFiles': ms.submittedFiles,
    'folderUrl': ms.folderUrl,
    'receiptFolderUrl': ms.receiptFolderUrl,
    'Ver_No': ms.verNo,
    'Submitted_At': ms.sendDateTime
  };
  
  // 大文字小文字・前後スペースを無視したマッチング
  headers.forEach((h, i) => {
    const headerLower = String(h).toLowerCase().trim();
    for (const key of Object.keys(mapping)) {
      if (key.toLowerCase() === headerLower) {
        if (mapping[key] !== undefined) newRow[i] = mapping[key];
        break;
      }
    }
  });
  
  sheet.appendRow(newRow);
}

/**
 * クライアントからの再投稿リクエストを処理
 */
function apiSubmitResubmission(data) {
  // 入力バリデーション
  validateRequiredString(data.msKey, '原稿キー (msKey)');
  if (data.authorEmail) validateEmail(data.authorEmail, '著者メールアドレス (authorEmail)');
  if (data.ccEmails)    validateEmailList(data.ccEmails, 'CC メールアドレス (ccEmails)');
  if (data.files && data.files.length > 0) {
    data.files.forEach((file, i) => validateFileName(file.name, 'ファイル名 ' + (i + 1)));
  }

  const ssId = getSpreadsheetId();
  const settings = getSettings();

  // 1. 前回（最新）の原稿データを取得
  const prevMs = getManuscriptData('author', data.msKey);
  if (!prevMs) throw new Error("Previous manuscript not found.");

  // 二重投稿防止: 同じ MS_ID でより新しいバージョンがすでに存在する場合はエラー
  const prevVerNo = Number(prevMs.Ver_No || 1);
  const prevMsId  = prevMs.MS_ID || '';
  if (prevMsId) {
    const allVersions = findAllRecordsByKey(ssId, MANUSCRIPTS_SHEET_NAME, 'MS_ID', prevMsId);
    const newerExists = allVersions.some(v => Number(v.Ver_No || 0) > prevVerNo);
    if (newerExists) {
      throw new Error('この原稿はすでに再投稿されています。二重投稿はできません。/ This manuscript has already been resubmitted. Duplicate submissions are not allowed.');
    }
  }
  
  // 2. 新しい版の情報を作成
  const nowId = prevMs.ID || prevMs.nowId;
  const verNo = parseMsVer(prevMs.MsVer).verNo + 1;
  const msId = prevMs.MS_ID;
  const msVer = msId + '-' + verNo;
  const msVerHex = getMsVerHEX(nowId, verNo);
  const key = msVerHex + Utilities.getUuid().replace(/-/g, '');         // 著者用キー — CSPRNG (128 bit)
  const eicKey = 'E' + Utilities.getUuid().replace(/-/g, '');           // 再投稿版にも新しいeicKeyを発行 — CSPRNG (128 bit)

  const thisManuscript = {
    nowId: nowId,
    MS_ID: msId,
    MsVer: msVer,
    verNo: verNo,
    MsVerHex: msVerHex,
    key: key,
    eicKey: eicKey,
    authorName: data.authorName || prevMs.CA_Name,
    authorEmail: data.authorEmail || prevMs.CA_Email,
    authorAffiliation: data.authorAffiliation || prevMs.CA_Affiliation || '',
    authorAddress: data.authorAddress || prevMs.Address || '',
    authorsJp: data.authorsJp || prevMs.AuthorsJP || '',
    authorsEn: data.authorsEn || prevMs.AuthorsEN || '',
    ccEmails: data.ccEmails || prevMs.ccEmails || '',
    paperType: data.paperType || prevMs.MS_Type,
    titleJp: data.titleJp || prevMs.TitleJP || '',
    titleEn: data.titleEn || prevMs.TitleEN || '',
    runningTitle: data.runningTitle || prevMs.RunningTitle || '',
    abstractJp: data.abstractJp || prevMs.AbstractJP || '',
    abstractEn: data.abstractEn || prevMs.AbstractEN || '',
    letterToEditor: data.letterToEditor || '',
    englishEditing: data.englishEditing || prevMs['English_editing'] || '',
    reprintRequest: data.reprintRequest || prevMs['Reprint request'] || '',
    sendDateTime: Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss')
  };

  // 3. フォルダ作成とファイル保存
  // submitted フォルダ: 投稿原稿を保存（担当編集者・査読者に共有）
  const folder = createSubmissionFolder(settings, thisManuscript);
  const fileNames = [];
  if (data.files && data.files.length > 0) {
    data.files.forEach(file => {
      try {
        const blob = Utilities.newBlob(
          Utilities.base64Decode(file.content),
          file.mimeType || 'application/octet-stream',
          file.name
        );
        folder.createFile(blob);
        fileNames.push(file.name);
        Logger.log(`Saved file to submitted folder: ${file.name}`);
      } catch (fileErr) {
        Logger.log(`Failed to save file "${file.name}": ${fileErr}`);
      }
    });
  } else {
    Logger.log(`Resubmission ${msVer}: data.files is empty or undefined`);
  }
  thisManuscript.submittedFiles = fileNames.join(', ');

  // submitted フォルダを共有設定（リンクを知っている全員が閲覧可）
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  thisManuscript.folderUrl = folder.getUrl();

  // receipt フォルダ: 受領票PDFを保存（編集委員長・印刷担当者のみに共有、担当編集者・査読者には不開示）
  const receiptFolder = createReceiptFolder(settings, thisManuscript);
  receiptFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  thisManuscript.receiptFolderUrl = receiptFolder.getUrl();

  // 4. データベースへの登録（folderUrl / receiptFolderUrl 設定後に実行）
  addNewMsToDB(ssId, thisManuscript);

  writeLog(`Resubmission: ${msVer} by ${thisManuscript.authorEmail}`);

  // 5. 受領通知と通知先の分岐（各メール送信は独立して実行）
  const emailWarnings = [];

  try {
    const receiptBlob = sendReceiptEmail(ssId, thisManuscript);
    // 受領票PDFを receipt フォルダに保存（submitted フォルダには保存しない）
    if (receiptBlob) {
      try { receiptFolder.createFile(receiptBlob); } catch (e) { Logger.log('Receipt PDF save error: ' + e); }
    }
  } catch (err) {
    Logger.log('sendReceiptEmail error: ' + err);
    emailWarnings.push('receipt');
  }

  // ★ 最終チェックフローへの経路判定
  // 条件1: 前バージョンが採択済み (in_production後の最終修正)
  // 条件2: 前バージョンが最終チェック中 (ルートaで著者に差し戻された後の再投稿)
  const wasAccepted = (String(prevMs.accepted || '').toLowerCase() === 'yes' || String(prevMs.isAccepted || '').toLowerCase() === 'yes');
  const wasInFinalReview = String(prevMs.finalStatus || '').trim() === 'final_review';

  if (wasAccepted || wasInFinalReview) {
    // 【最終チェックフロー】 担当編集者を介さず直接編集幹事へ
    try {
      if (!settings.managingEditorEmail) {
        writeLog('[ERROR] Resubmission Routing: 受理後の再投稿ですが、managingEditorEmail が未設定です。');
        sendEicNotification(thisManuscript); // フォールバック: EICへ
      } else {
        const managingEditorKey = Utilities.getUuid();
        // DBを更新（幹事用キーとステータスをセット）
        updateLogCell(ssId, MANUSCRIPTS_SHEET_NAME, 'key', thisManuscript.key, {
          'managingEditorKey': managingEditorKey,
          'finalStatus':       'final_review'
        });
        // 編集幹事へ通知
        sendResubmittedAcceptedNotificationToManagingEditor(thisManuscript, settings, managingEditorKey);
        writeLog(`Resubmission Routing: ${msVer} - Accepted manuscript resubmitted. Routed to Managing Editor.`);
      }
    } catch (err) {
      Logger.log('Managing Editor notification error: ' + err);
      emailWarnings.push('me');
    }
  } else {
    // 【通常の再投稿フロー】 従来通りEICへ通知して担当編集者指名へ
    try {
      sendEicNotification(thisManuscript);
    } catch (err) {
      Logger.log('sendEicNotification error: ' + err);
      emailWarnings.push('eic');
    }
  }

  return {
    success: true,
    msId: msId,
    msVer: msVer,
    key: key,
    emailWarning: emailWarnings.length > 0 ? emailWarnings : null
  };
}


function getMsVerHEX(id, ver) {
  const hex_id = ('0000' + id.toString(16)).slice(-4);
  const hex_ver = ver.toString(16);
  const result = (hex_id + hex_ver).length < 5 ? ('0' + hex_id + hex_ver) : (hex_id + hex_ver);
  return 'H' + result;
}

/**
 * 投稿用フォルダ作成
 * submitted フォルダ: 投稿原稿を格納。担当編集者・査読者に URL を共有する。
 */
function createSubmissionFolder(settings, ms) {
  const rootName = settings.SUBFOLDER || 'Journal Files';
  const root = driveFolderCache.getRootFolder(rootName) || DriveApp.createFolder(rootName);
  const msFolder = driveFolderCache.getOrCreateFolder(root, ms.MS_ID);
  const verFolder = driveFolderCache.getOrCreateFolder(msFolder, 'ver.' + ms.verNo);
  return driveFolderCache.getOrCreateFolder(verFolder, 'submitted');
}

/**
 * 受領票用フォルダ作成
 * receipt フォルダ: 受領票PDFを格納。編集委員長・印刷担当者のみに URL を共有する。
 * 担当編集者・査読者には URL を伝えないため、submitted フォルダとは分離する。
 */
function createReceiptFolder(settings, ms) {
  const rootName = settings.SUBFOLDER || 'Journal Files';
  const root = driveFolderCache.getRootFolder(rootName) || DriveApp.createFolder(rootName);
  const msFolder = driveFolderCache.getOrCreateFolder(root, ms.MS_ID);
  const verFolder = driveFolderCache.getOrCreateFolder(msFolder, 'ver.' + ms.verNo);
  return driveFolderCache.getOrCreateFolder(verFolder, 'receipt');
}
