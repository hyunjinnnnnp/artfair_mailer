
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
 * 한 행의 이메일 발송을 시도하고 결과(성공/실패)를 시트에 기록.
 */
function email_handleRowSend(row, rowNum, fileMap){
  const email = row[COL_INDEX.EMAIL];
  const name = row[COL_INDEX.NAME];
  const artistsRaw = row[COL_INDEX.ARTISTS];
  const artistList = artistsRaw.split(",").map(a => a.trim());

  const sent = email_sendArtistEmail(email, name, artistList, fileMap);

  if (!sent) {
    throw new Error(`❌ 이메일 전송 실패 (행 ${rowNum}): ${email}`);
  }
  handleSuccessMessage(rowNum);
}
