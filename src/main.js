// TO DO: 스크립트 네임, 스프레드 시트 네임, 폴더 네임 관계 자동으로 만들 수 있을까?
// TO DO: 작가 이름 철자 체크 쉽게 상단에서 진행할 것


/**
 * 버튼 클릭 시 이메일 발송
 */

function handleSendButtonClick() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues(); // data 인덱스는 0부터 시작하는 값

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

    const fileMap = drive_getPdfFileMap();
  

    // ++++++ 하나의 이메일 보낼 때마다 함수 호출 x
    // []을 넘겨주고 안에서 처리한다
    // email_handleRow 원래 사용하던 함수들 다 바꿔줘야 함
    // try catch????
    email_handleRows(pendingRows, fileMap);
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
    const errorDetails = { error, row }
    handleErrorMessage(errorDetails, '폼 응답 처리 중 에러 발생')
  }
}

/**
 * MEMO 컬럼을 제외한 다른 컬럼에 대한 수정이 발생하면, 해당 수정은 자동으로 이전 값으로 되돌려집니다.
 */
// function onEdit(e) {
//   const memoCol = COL_NUM.MEMO;
//   const editedCol = e.range.getColumn();

//   if (editedCol !== memoCol) {
//     const oldValue = e.oldValue;
//     e.range.setValue(oldValue);
//   }
// }


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

/** 
 * 이메일 발송시간 형식 변환
 */
function formatEmailSentAtColumn() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const colIndex = COL_NUM.EMAIL_SENT_AT;  // 예: 5열이라면 5

  // 해당 열 전체 범위 가져오기 (예: A:A, B:B ...)
  const range = sheet.getRange(2, colIndex, sheet.getMaxRows() - 1); // 헤더 제외

  // 날짜/시간 포맷 설정
  range.setNumberFormat("yyyy. m. d 오전/오후 h:mm:ss");
}

function onOpen() {
  protectColumns();
  initializeHeaders();
  formatEmailSentAtColumn();
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("🖼 갤러리 도구")
  .addItem("✅ 구글 드라이브 폴더명 체크하기", "drive_checkFolderExistence")
  .addItem("📧 이메일 발송 시작", "handleSendButtonClick")
  .addToUi();
}
