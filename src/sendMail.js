function sendArtistPdfsUsingSheetName() {
  const parentFolderName = "아트페어_PDF";
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const spreadsheetName = SpreadsheetApp.getActiveSpreadsheet().getName();
  const data = sheet.getDataRange().getValues();

  try {
    // 상위 폴더 확인
    const parentFolders = DriveApp.getFoldersByName(parentFolderName);
    if (!parentFolders.hasNext()) {
      throw new Error(`Google 드라이브에 '${parentFolderName}'라는 폴더가 없습니다. \n\n📁 '내 드라이브' 최상단에 해당 폴더가 존재하는지 확인해주세요.`);
    }
    const parentFolder = parentFolders.next();

    // 아트페어 폴더 확인
    const fairFolders = parentFolder.getFoldersByName(spreadsheetName);
    if (!fairFolders.hasNext()) {
      throw new Error(`'${parentFolderName}' 폴더 안에 '${spreadsheetName}'이라는 이름의 아트페어 폴더가 없습니다.\n\n📁 폴더명이 구글 시트 제목과 정확히 일치하는지 확인해주세요.`);
    }
    const fairFolder = fairFolders.next();

    Logger.log(`✅ 상위 폴더: ${parentFolderName}, 아트페어 폴더: ${spreadsheetName}`);

    // 폴더 내 파일 맵 만들기
    const fileMap = new Map();
    const files = fairFolder.getFiles();
    let fileCount = 0;
    while (files.hasNext()) {
      const file = files.next();
      fileMap.set(file.getName(), file);
      fileCount++;
    }
    Logger.log(`📁 '${spreadsheetName}' 폴더 내 파일 수: ${fileCount}`);
    Logger.log(`🗂 파일 목록: ${[...fileMap.keys()].join(", ")}`);

    // 본격적으로 이메일 발송
    data.slice(1).forEach((row, idx) => {
      const rowNum = idx + 2; // 시트의 실제 행 번호
      const email = row[1];
      const artistsRaw = row[2];
      const status = row[3];

      Logger.log(`\n🔽 [Row ${rowNum}]`);
      Logger.log(`📬 이메일: ${email}`);
      Logger.log(`🎨 작가 입력값: "${artistsRaw}"`);
      Logger.log(`📦 현재 상태: ${status}`);

      if (status !== "전송됨" && email && artistsRaw) {
        const artistList = artistsRaw.split(",").map(a => a.trim());
        Logger.log(`🎯 파싱된 작가 리스트: ${artistList.join(", ")}`);

        const attachments = [];

        artistList.forEach(name => {
          const expectedFileName = `${name}.pdf`;
          const file = fileMap.get(expectedFileName);

          if (file) {
            attachments.push(file.getAs(MimeType.PDF));
            Logger.log(`✅ 파일 첨부됨: ${expectedFileName}`);
          } else {
            Logger.log(`❌ PDF 파일 없음: ${expectedFileName}`);
          }
        });

        if (attachments.length > 0) {
          const subject = `${spreadsheetName} - 작가 작품 정보`;
          const body = `안녕하세요,\n${spreadsheetName}에서 관심 주신 작가님의 PDF를 첨부드립니다:\n\n${artistList.join(", ")}`;

          GmailApp.sendEmail(email, subject, body, {
            attachments: attachments
          });

          Logger.log(`📨 이메일 발송 완료: ${email}`);

          sheet.getRange(rowNum, 4).setValue("전송됨");
          sheet.getRange(rowNum, 5).setValue(new Date());
        } else {
          Logger.log(`⚠️ 첨부할 파일이 없어 이메일 미발송`);
        }
      } else {
        Logger.log(`⏭️ 조건 불충족으로 건너뜀`);
      }
    });

    ui.alert("✅ 이메일 발송이 완료되었습니다.");
  } catch (error) {
    Logger.log("🚨 오류 발생: " + error.message);
    ui.alert("⚠️ 이메일 발송 중 문제가 발생했습니다.\n\n" + error.message);
  }
}

// 스프레드시트의 고객 정보 수정 제한
function protectColumns() {
  // Protect the active sheet, then remove all other users from the list of
  // editors.
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
  const sheet = e.range.getSheet();
  const protectedCols = [1, 2, 3, 4];  // 보호할 열 (1~4열: 타임스탬프, 이름, 이메일, 작가 목록)
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

const artfair_mailer = {
  onOpen,
  sendArtistPdfsUsingSheetName,
}