
/**
 * 스프레드 시트 제목으로 구글 드라이브 내부에 해당 아트페어 폴더가 있는지 검색.
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
      throw new Error(`❌ Google 드라이브 '${parentFolderName}' 폴더 안에서 '${spreadsheetName}' 폴더를 찾을 수 없습니다.`);
    }

    return fairFolders.next(); // 정상적으로 찾은 폴더 반환   
}


/**
 * PDF 파일 맵 생성
 */
function drive_getPdfFileMap() {
    const fairFolder = drive_getTargetFolder();
    const fileMap = new Map();
    const files = fairFolder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      fileMap.set(file.getName(), file);
    }
    return fileMap;
}