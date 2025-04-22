// TO DO: 스크립트 네임, 스프레드 시트 네임, 폴더 네임 관계 자동으로 만들 수 있을까?
// TO DO: 작가 이름 철자 체크 쉽게 상단에서 진행할 것

/**
 * 공통: 스프레드 시트 제목으로 구글 드라이브 내부에 해당 아트페어 폴더가 있는지 검색.
 */
function getTargetFolder() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetName = sheet.getName();

    const parentFolders = DriveApp.getFoldersByName(parentFolderName);
    if (!parentFolders.hasNext()) {
      throw new Error(`❌ Google 드라이브에 '${parentFolderName}'라는 폴더가 없습니다.`);
    }
    const parentFolder = parentFolders.next();

    const fairFolders = parentFolder.getFoldersByName(spreadsheetName);
    if (!fairFolders.hasNext()) {
      throw new Error(`❌ Google 드라이브 '${parentFolderName}' 폴더 안에서 '${spreadsheetName}' 폴더를 찾을 수 없습니다.`);
    }

    return fairFolders.next(); // 정상적으로 찾은 폴더 반환   
}


/**
 * 공통: PDF 파일 맵 생성
 */
function getPdfFileMap() {
    const fairFolder = getTargetFolder();
    const fileMap = new Map();
    const files = fairFolder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      fileMap.set(file.getName(), file);
    }
    return fileMap;
}

/**
 * 공통: 이메일 발송 함수
 */
function sendArtistEmail(email, name, artistList, fileMap) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetName = sheet.getName();
  const attachments = [];
  artistList.forEach(artist => {
    const file = fileMap.get(`${artist}.pdf`);
    if (file) {
      attachments.push(file.getAs(MimeType.PDF));
    }
  });
 
  if (attachments.length > 0) {
    const subject = `${spreadsheetName} - 작가 작품 정보`;
    const body = `${name}님 안녕하세요,\n${spreadsheetName}에서 관심 주신 작가님의 PDF를 첨부드립니다:\n\n${artistList.join(", ")}`;
    GmailApp.sendEmail(email, subject, body, { attachments });
    return true;
  } else {
    return false;
  }
}

/**
 * 버튼 클릭 시 발송
 */

function handleSendButtonClick() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  let rowNum = null;
  // TO DO: 모든 데이터를 가져와서 확인하는 게 아니라, 전송 안된 목록만 가져와서 처리하는 게 효율적이지 않을까?

  try {
    const fileMap = getPdfFileMap();
  
    data.slice(1).forEach((row, idx) => {
      rowNum = idx + 2;
      const email = row[COL_INDEX.EMAIL];
      const name = row[COL_INDEX.NAME];
      const artistsRaw = row[COL_INDEX.ARTISTS];
      const status = row[COL_INDEX.STATUS];

      if (status === "전송됨" || !email || !name || !artistsRaw){
        return;
      }

      const artistList = artistsRaw.split(",").map(a => a.trim());
      const sent = sendArtistEmail(email, name, artistList, fileMap);
      const now = new Date();

      if (!sent) {
        throw new Error(`❌ 이메일 전송 실패 (행 ${rowNum}): ${email}`);
      }
      sheet.getRange(rowNum, COL_NUM.STATUS).setValue("전송됨");
      sheet.getRange(rowNum, COL_NUM.EMAIL_SENT_AT).setValue(now);
    });
    ui.alert("✅ 이메일 발송이 완료되었습니다.")

  } catch (error) {
    Logger.log("🚨 전체 처리 중 오류 발생: " + error.message);
    ui.alert("❌ 이메일 발송 중 중 오류가 발생했습니다" + error.message);
    if (rowNum !== null) {
      sheet.getRange(rowNum, COL_NUM.STATUS).setValue("전체 처리 오류");
      sheet.getRange(rowNum, COL_NUM.EMAIL_SENT_AT).setValue(new Date());
      sheet.getRange(rowNum, COL_NUM.ERROR).setValue(error.message);
    }
  }
}

/**
 * 폼 응답시 자동 발송
 */
function onFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = e.range.getRow();

  try {
    const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    const email = rowData[COL_INDEX.EMAIL];
    const name = rowData[COL_INDEX.NAME];
    const artistsRaw = rowData[COL_INDEX.ARTISTS];
    const status = rowData[COL_INDEX.STATUS];
    const fileMap = getPdfFileMap();

    if (status !== "전송됨" && email && name && artistsRaw) {
      const artistList = artistsRaw.split(",").map(a => a.trim());
      const sent = sendArtistEmail(email, name, artistList, fileMap);

      if (sent) {
        sheet.getRange(row, COL_NUM.STATUS).setValue("전송됨");
        const formattedDate = Utilities.formatDate(new Date(), TIMEZONE.SEOUL, "yyyy-MM-dd HH:mm:ss");
        sheet.getRange(row, COL_NUM.EMAIL_SENT_AT).setValue(formattedDate);
      }
    }
  } catch (error) {
    Logger.log("🚨 오류 발생: " + error.message);
    sheet.getRange(row, COL_NUM.ERROR).setValue(error.message);
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

/**
 * 스프레드시트의 고객 정보 수정 제한
 */
function protectColumns() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const protection = sheet.protect().setDescription('Sample protected sheet');

  // Ensure the current user is an editor before removing others. Otherwise, if
  // the user's edit permission comes from a group, the script throws an exception
  // upon removing the group.
  const me = Session.getEffectiveUser();
  protection.setWarningOnly(false); // 설정해야만 add, remove editor 사용 가능
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
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


function onOpen() {
  protectColumns();
  initializeHeaders();
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🖼 갤러리 도구')
    .addItem('이메일 발송 시작', 'handleSendButtonClick')
    .addToUi();
}

// const artfair_mailer = {
//   onOpen,
//   handleSendButtonClick,
// }
