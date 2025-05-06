
/**
 * 이메일 발송 함수
 */
function email_sendArtistEmail(email, name, artistList, fileMap) {
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
 * 이메일 발송을 시도하고 결과(성공/실패)를 시트에 기록.
 */
function email_handleRows(pendingRows, fileMap){
  const errors = [];
  const ui = SpreadsheetApp.getUi();

    pendingRows.forEach((obj)=> {
      const row = obj.row;
      const rowNum = obj.index + 1;
      const email = row[COL_INDEX.EMAIL];
      const name = row[COL_INDEX.NAME];
      const artistsRaw = row[COL_INDEX.ARTISTS];
      const artistList = artistsRaw.split(",").map(a => a.trim());
      // 하나의 이메일이 발송실패해도 다음 이메일은 발송되어야 한다
      try {
        const sent = email_sendArtistEmail(email, name, artistList, fileMap);
        if(!sent){
          ui.alert("이메일 발송 실패::: 알 수 없는 오류")
        }
        handleSuccessMessage(rowNum);
      } catch (error) {
        errors.push({ error, rowNum });
        handleErrorMessage(errors, '이메일 발송 실패');
      }
    })

  
  if (errors.length < 1) {
    ui.alert("✅ 모든 이메일 발송 성공");
  }else{
    ui.alert("❌ 발송에 실패한 이메일이 있습니다");
  }

}
