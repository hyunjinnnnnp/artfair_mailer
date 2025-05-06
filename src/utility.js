// ui 알림창은 버튼 클릭했을 때만 사용 가능
// 이메일 보낸 후 에러 혹은 성공 처리는(로깅, 시트에 결과 반영하기) handler에게 위임한다.



function handleErrorMessage(error, contextMessage='', rowNum) {
  const fullMessage = [
    '❌ 오류 발생',
    contextMessage && `📍위치: ${contextMessage}`,
    `메시지: ${error.message}`,
    `스택 추적:\n${error.stack}`
  ].filter(Boolean).join('\n');

  Logger.log(fullMessage);

  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(rowNum, COL_NUM.STATUS).setValue(STATUS.ERROR);
  sheet.getRange(rowNum, COL_NUM.EMAIL_SENT_AT).setValue(new Date());
  // sheet.getRange(rowNum, COL_NUM.EMAIL_SENT_AT).setNumberFormat("yyyy. m. d 오전/오후 h:mm:ss");
  sheet.getRange(rowNum, COL_NUM.ERROR).setValue(error.message);

}


function handleSuccessMessage(rowNum){
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(rowNum, COL_NUM.STATUS).setValue(STATUS.SENT);
  sheet.getRange(rowNum, COL_NUM.EMAIL_SENT_AT).setValue(new Date());
  sheet.getRange(rowNum, COL_NUM.ERROR).setValue("");
  
  try{
  SpreadsheetApp.getUi().alert(`이메일 발송 완료`);
  }catch(_){
    // UI 사용 불가능한 상황에서는 아무 동작 x
  }
}