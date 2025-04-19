const parentFolderName = "아트페어_PDF";

/**
 * 공통: PDF 파일 맵 생성
 */
function getPdfFileMap() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetName = sheet.getName();
  const parentFolders = DriveApp.getFoldersByName(parentFolderName);

  if (!parentFolders.hasNext()) {
    throw new Error(`Google 드라이브에 '${parentFolderName}'라는 폴더가 없습니다.`);
  }
  const parentFolder = parentFolders.next();

  const fairFolders = parentFolder.getFoldersByName(spreadsheetName);
  if (!fairFolders.hasNext()) {
    throw new Error(`'${parentFolderName}' 폴더 안에 '${spreadsheetName}'이라는 폴더가 없습니다.`);
  }
  const fairFolder = fairFolders.next();

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
function sendArtistEmail(email, artistList, fileMap) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetName = sheet.getName();
  const attachments = [];
  artistList.forEach(name => {
    const file = fileMap.get(`${name}.pdf`);
    if (file) {
      attachments.push(file.getAs(MimeType.PDF));
    }
  });

  if (attachments.length > 0) {
    const subject = `${spreadsheetName} - 작가 작품 정보`;
    const body = `안녕하세요,\n${spreadsheetName}에서 관심 주신 작가님의 PDF를 첨부드립니다:\n\n${artistList.join(", ")}`;
    GmailApp.sendEmail(email, subject, body, { attachments });
    return true;
  } else {
    return false;
  }
}

/**
 * 버튼 클릭 시 발송
 */
function sendArtistPdfsUsingSheetName() {
  const ui = SpreadsheetApp.getUi();
  // get active sheet로 할 때만 작동 !!!
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  try {
    const fileMap = getPdfFileMap();

    data.slice(1).forEach((row, idx) => {
      const rowNum = idx + 2;
      const email = row[1];
      const artistsRaw = row[2];
      const status = row[3];

      if (status !== "전송됨" && email && artistsRaw) {
        const artistList = artistsRaw.split(",").map(a => a.trim());
        const sent = sendArtistEmail(email, artistList, fileMap);

        if (sent) {
          sheet.getRange(rowNum, 4).setValue("전송됨");
          sheet.getRange(rowNum, 5).setValue(new Date());
        }
      }
    });

    ui.alert("✅ 이메일 발송이 완료되었습니다.");
  } catch (error) {
    Logger.log("🚨 오류 발생: " + error.message);
    ui.alert("⚠️ 이메일 발송 중 문제가 발생했습니다.\n\n" + error.message);
  }
}

/**
 * 폼 응답시 자동 발송
 */
function onFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // const ui = SpreadsheetApp.getUi();
  // 폼서브밋은 UI를 사용할 수 없는 백그라운드에서 실행되기 때문에 여기 작성하면 안됨
  const row = e.range.getRow();

  try {
    const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    const email = rowData[1];
    const artistsRaw = rowData[2];
    const status = rowData[3];
    const fileMap = getPdfFileMap();

    sheet.getRange(row, 4).setValue("폼 트리거 작동함");
      if (status !== "전송됨" && email && artistsRaw) {
        const artistList = artistsRaw.split(",").map(a => a.trim());
        const sent = sendArtistEmail(email, artistList, fileMap);

        if (sent) {
          sheet.getRange(row, 4).setValue("전송됨");
          sheet.getRange(row, 5).setValue(new Date());
        }
      }
    

    // ui.alert("✅ 이메일 발송이 완료되었습니다.");
  } catch (error) {
    Logger.log("🚨 오류 발생: " + error.message);
    // ui.alert("⚠️ 이메일 발송 중 문제가 발생했습니다.\n\n" + error.message);
  }
}

// 스프레드시트의 고객 정보 수정 제한
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

// 스프레드시트 값이 임의로 수정되었을 경우 롤백
function onEdit(e) {
  const protectedCols = [1, 2, 3, 4, 5];
  // 이후에 추가되는 내용이 있을지도 모르니 타임스탬프~발송일시까지만 수정을 막는다
  const col = e.range.getColumn();      // 수정된 열 번호

  // 수정된 열이 보호된 열 중 하나라면
  if (protectedCols.includes(col)) {
    const oldValue = e.oldValue;
    // 수정된 값을 이전 값으로 되돌림 (롤백)
    e.range.setValue(oldValue);
  }
}

function onOpen() {
  protectColumns();
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🖼 갤러리 도구')
    .addItem('이메일 발송 시작', 'sendArtistPdfsUsingSheetName')
    .addToUi();
}

// const artfair_mailer = {
//   onOpen,
//   sendArtistPdfsUsingSheetName,
// }