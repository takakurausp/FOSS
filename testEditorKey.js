/**
 * Test function to check Editor_log sheet column names
 */
function testEditorKeyColumns() {
  try {
    const ssId = getSpreadsheetId();
    const sheetName = 'Editor_log';
    
    const sheet = SpreadsheetApp.openById(ssId).getSheetByName(sheetName);
    if (!sheet) {
      Logger.log('Sheet not found: ' + sheetName);
      return;
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log('Editor_log headers:');
    headers.forEach((header, index) => {
      Logger.log((index + 1) + ': "' + header + '"');
    });
    
    // Check for editorKey/EditorKey
    const hasEditorKey = headers.some(h => String(h).toLowerCase().trim() === 'editorkey');
    Logger.log('Has editorKey (case-insensitive): ' + hasEditorKey);
    
    // Find exact matches
    const exactEditorKey = headers.find(h => String(h).trim() === 'editorKey');
    const exactEditorKeyCapital = headers.find(h => String(h).trim() === 'EditorKey');
    
    Logger.log('Exact "editorKey" match: ' + (exactEditorKey || 'NOT FOUND'));
    Logger.log('Exact "EditorKey" match: ' + (exactEditorKeyCapital || 'NOT FOUND'));
    
    return headers;
  } catch (error) {
    Logger.log('Error in testEditorKeyColumns: ' + error.toString());
    throw error;
  }
}