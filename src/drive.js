/**
 * 구글 드라이브에 두 폴더가 존재하는지만 체크 후 ui 알림 (데이터 반환 X)
 * @returns {boolean}
 */
function drive_checkFolderExistence() {
  const ui = SpreadsheetApp.getUi();
  const parentFolders = DriveApp.getFoldersByName(USER.GOOGLE_PARENT_FOLDER);
  if (!parentFolders.hasNext()) {
    ui.alert(`구글 드라이브에서 ${USER.GOOGLE_PARENT_FOLDER} 폴더를 찾을 수 없습니다.
    폴더 이름이 정확한지 다시 확인해 주세요.`);
    return;
  };

  const parentFolder = parentFolders.next();
  const fairFolders = parentFolder.getFoldersByName(USER.GOOGLE_FAIR_FOLDER);
  if(!fairFolders.hasNext()){
    ui.alert(`구글 드라이브에서 ${USER.GOOGLE_PARENT_FOLDER}/${USER.GOOGLE_FAIR_FOLDER} 폴더를 찾을 수 없습니다.
    폴더 이름이 정확한지 다시 확인해 주세요.`);
    return;
  }
  ui.alert(`✅ 구글 드라이브에서 '${USER.GOOGLE_PARENT_FOLDER}' 폴더 안에 '${USER.GOOGLE_FAIR_FOLDER}' 폴더를 확인했습니다.`);
}

/**
 * 스프레드 시트 제목으로 구글 드라이브에서 아트페어 폴더를 검색, 해당 폴더를 반환한다.
 * @returns {Folder} 구글 드라이브에서 찾은 아트페어 폴더
 */
function drive_getTargetFolder() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetName = sheet.getName();

    const parentFolders = DriveApp.getFoldersByName(parentFolderName);
    if (!parentFolders.hasNext()) {
      throw new Error(`❌ Google 드라이브에 '${parentFolderName}'라는 폴더가 없습니다.`);
    }
    const parentFolder = parentFolders.next();

    const fairFolders = parentFolder.getFoldersByName(spreadsheetName);
    if (!fairFolders.hasNext()) {
      throw new Error(`❌ Google 드라이브에서 '${parentFolderName}'/'${spreadsheetName}' 폴더를 찾을 수 없습니다.`);
    }

    return fairFolders.next(); // 정상적으로 찾은 폴더 반환   
}


/**
 * PDF 파일 맵 생성
 */
function drive_getPdfFileMap() {
    const fairFolder = drive_getTargetFolder();
    if(!fairFolder){
      return;
    }
    const fileMap = new Map();
    const files = fairFolder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      fileMap.set(file.getName(), file);
    }
    return fileMap;
}