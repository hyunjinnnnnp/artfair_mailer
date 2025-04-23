function handleErrorMessage(error, contextMessage='', row) {
  const fullMessage = [
    '❌ 오류 발생',
    contextMessage && `📍위치: ${contextMessage}`,
    `메시지: ${error.message}`,
    `스택 추적:\n${error.stack}`
  ].filter(Boolean).join('\n');
  Logger.log(fullMessage);

  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(row, COL_NUM.STATUS).setValue(STATUS.ERROR);
  sheet.getRange(row, COL_NUM.EMAIL_SENT_AT).setValue(new Date());
  sheet.getRange(row, COL_NUM.ERROR).setValue(error.message);

  try{
    SpreadsheetApp.getUi().alert(`:\n\n${error.message}`);
  }catch(_){
    // UI 사용 불가능한 상황에서는 아무 동작 x
  }
}