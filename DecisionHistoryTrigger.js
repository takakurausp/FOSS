function logDecisionHistory(e) {
  if (!e || !e.range || !e.source) {
    throw new Error('Invalid event object');
  }

  const sheet = e.range.getSheet();
  if (sheet.getName() !== 'Decisions') return;

  const row = e.range.getRow();
  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const historySheet = e.source.getSheetByName('DECISION_HISTORY');
  historySheet.appendRow([
    row,
    JSON.stringify({
      decision: e.oldValue,
      isAccepted: rowData[2], 
      resubmit: rowData[3]
    }),
    JSON.stringify({
      decision: e.value,
      isAccepted: rowData[2],
      resubmit: rowData[3]
    }),
    Session.getActiveUser().getEmail(),
    new Date()
  ]);
}
