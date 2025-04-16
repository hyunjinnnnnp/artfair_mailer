function sendArtistPdfsUsingSheetName() {
  const parentFolderName = "아트페어_PDF"; // 상위 폴더명
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const spreadsheetName = SpreadsheetApp.getActiveSpreadsheet().getName(); // 시트 제목 = 아트페어 이름
  const data = sheet.getDataRange().getValues();

  const parentFolders = DriveApp.getFoldersByName(parentFolderName);
  if (!parentFolders.hasNext()) {
    SpreadsheetApp.getUi().alert("상위 폴더가 없습니다: " + parentFolderName);
    return;
  }

  const parentFolder = parentFolders.next();
  const fairFolders = parentFolder.getFoldersByName(spreadsheetName);
  if (!fairFolders.hasNext()) {
    SpreadsheetApp.getUi().alert(
      "아트페어 폴더가 없습니다: " + spreadsheetName
    );
    return;
  }

  const fairFolder = fairFolders.next();
  for (let i = 1; i < data.length; i++) {
    const email = data[i][1];
    const artistsRaw = data[i][2];
    const status = data[i][3];

    if (status !== "전송됨" && email && artistsRaw) {
      const artistList = artistsRaw.split(",").map((a) => a.trim());
      const attachments = [];

      // 드라이브 전체에서 폴더명으로 검색
      // artistList.forEach(name => {
      //   const files = fairFolder.getFilesByName(`${name}.pdf`);
      //   if (files.hasNext()) {
      //     attachments.push(files.next().getAs(MimeType.PDF));
      //   } else {
      //     Logger.log(`PDF 파일 없음: ${name}.pdf in ${spreadsheetName}`);
      //   }
      // });

      // 현 폴더 내부에서 폴더명 검색
      artistList.forEach((name) => {
        let found = false;
        const files = fairFolder.getFiles();
        while (files.hasNext()) {
          const file = files.next();
          console.log(file.getName());
          if (file.getName() === `${name}.pdf`) {
            attachments.push(file.getAs(MimeType.PDF));
            found = true;
            break;
          }
        }
        if (!found) {
          Logger.log(`PDF 파일 없음: ${name}.pdf in ${spreadsheetName}`);
        }
      });

      if (attachments.length > 0) {
        const subject = `${spreadsheetName} - 작가 작품 정보`;
        const body = `안녕하세요,\n${spreadsheetName}에서 관심 주신 작가님의 PDF를 첨부드립니다:\n\n${artistList.join(
          ", "
        )}`;

        GmailApp.sendEmail(email, subject, body, {
          attachments: attachments,
        });

        // 발송 상태 및 시간 기록
        sheet.getRange(i + 1, 4).setValue("전송됨");
        sheet.getRange(i + 1, 5).setValue(new Date());
      }
    }
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🖼 갤러리 도구")
    .addItem("이메일 발송 시작", "sendArtistPdfsUsingSheetName")
    .addToUi();
}
