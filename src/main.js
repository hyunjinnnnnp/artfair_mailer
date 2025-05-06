// TO DO: 스크립트 네임, 스프레드 시트 네임, 폴더 네임 관계 자동으로 만들 수 있을까?
// TO DO: 작가 이름 철자 체크 쉽게 상단에서 진행할 것


/**
 * 버튼 클릭 시 이메일 발송
 */

function handleSendButtonClick() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const pendingRows = data
    .map((row, index) => ({ row, index })) // 원래 행 보존
    .slice(1) // 헤더 제외
    .filter(obj => {
      const row = obj.row;
      return row[COL_INDEX.STATUS] !== STATUS.SENT && 
              row[COL_INDEX.EMAIL] && 
              row[COL_INDEX.NAME] && 
              row[COL_INDEX.ARTISTS];
    });

  try {
    const fileMap = drive_getPdfFileMap();
  
    pendingRows.forEach((obj) => {
      const row = obj.row;
      const rowNum = obj.index + 1;
      email_handleRowSend(row, rowNum, fileMap);
    });
  } catch (error) {
    Logger.log("🚨 이메일 발송 전체 처리 중 오류 발생: " + error.message);
    ui.alert("❌ 이메일 발송 중 중 오류가 발생했습니다" + error.message);
  }
}

/**
 * 고객이 구글 폼 응답시 이메일 자동 발송
 */
function onFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = e.range.getRow();

  try {
    const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    const fileMap = drive_getPdfFileMap();
    email_handleRowSend(rowData, row, fileMap);
    
  } catch (error) {
    handleErrorMessage(error, '폼 응답 처리 중 에러 발생', row)
  }
}

/**
 * MEMO 컬럼을 제외한 다른 컬럼에 대한 수정이 발생하면, 해당 수정은 자동으로 이전 값으로 되돌려집니다.
 */
function onEdit(e) {
  const memoCol = COL_NUM.MEMO;
  const editedCol = e.range.getColumn();

  if (editedCol !== memoCol) {
    const oldValue = e.oldValue;
    e.range.setValue(oldValue);
  }
}


/** 
 * 시트 헤더명 초기화
 */
function initializeHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const COL_LENGTH = Object.keys(COL_NUM).length;
  const firstRow = sheet.getRange(1, 1, 1, COL_LENGTH);
  const headers = Object.keys(COL_NUM);
  firstRow.setValues([headers]);
}


function onOpen() {
  protectColumns();
  initializeHeaders();
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("🖼 갤러리 도구")
  .addItem("✅ 구글 드라이브 폴더명 체크하기", "drive_checkFolderExistence")
  .addItem("📧 이메일 발송 시작", "handleSendButtonClick")
  .addToUi();
}
