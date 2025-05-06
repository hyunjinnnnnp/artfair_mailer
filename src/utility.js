// ui 알림창은 '메일 보내기 버튼' 클릭했을 때만 사용 가능함.
// 이메일 발송 후 에러/성공 처리는(로깅, 시트 업데이트) handler에게 위임한다.

function handleLogger(error, contextMessage){
  const fullMessage = [
      '❌ 오류 발생',
      contextMessage && `📍위치: ${contextMessage}`,
      `메시지: ${error.message}`,
      `스택 추적:\n${error.stack}`
    ].filter(Boolean).join('\n');

    Logger.log(fullMessage);  
}

function handleErrorMessage(errors, contextMessage='') {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // errors = [{row, error}] || {row, error}
  if(Array.isArray(errors)){
    Logger.log(errors)
    errors.map(({ error, rowNum }) => {
      sheet.getRange(rowNum, COL_NUM.STATUS).setValue(STATUS.ERROR);
      sheet.getRange(rowNum, COL_NUM.EMAIL_SENT_AT).setValue(new Date());
      sheet.getRange(rowNum, COL_NUM.ERROR).setValue(`${contextMessage}. ${error.message}`);
      handleLogger(error, contextMessage);
    })
  }else {
    const { error, rowNum } = errors;
    sheet.getRange(rowNum, COL_NUM.STATUS).setValue(STATUS.ERROR);
    sheet.getRange(rowNum, COL_NUM.EMAIL_SENT_AT).setValue(new Date());
    sheet.getRange(rowNum, COL_NUM.ERROR).setValue(`${contextMessage}. ${error.message}`);
    handleLogger(error, contextMessage);
  }
}


function handleSuccessMessage(rowNum){
  const sheet = SpreadsheetApp.getActiveSheet();
  if(Array.isArray(rowNum)){
    rowNum.map((row)=> {
      sheet.getRange(row, COL_NUM.STATUS).setValue(STATUS.SENT);
      sheet.getRange(row, COL_NUM.EMAIL_SENT_AT).setValue(new Date());
      sheet.getRange(row, COL_NUM.ERROR).setValue("");
    })
  }else{
    sheet.getRange(rowNum, COL_NUM.STATUS).setValue(STATUS.SENT);
    sheet.getRange(rowNum, COL_NUM.EMAIL_SENT_AT).setValue(new Date());
    sheet.getRange(rowNum, COL_NUM.ERROR).setValue("");
  }
}