/**
 * 사용자에게 폴더를 찾을 수 없다는 메시지와 구글 드라이브 링크를 포함한 모달을 띄우는 함수
 * @param {string} folderName - 확인하려는 폴더 이름
 */
function showFolderNotFoundDialog(folderName) {
  const htmlOutput = HtmlService.createHtmlOutput(
    `<style>
      .dialog-content {
        max-width: 100%;
        max-height: 400px;
        overflow-y: auto;
      }
    </style>
    <div class="dialog-content">
    <p>❌ Google 드라이브에 '${folderName}'라는 폴더가 없습니다.</p>
    <p>폴더 이름과 위치가 정확한지 다시 확인해 주세요.</p>
    <br>
    <p><strong>폴더 구조는 다음과 같습니다:</strong></p>
    <p>Google 드라이브</p>
    <p>&nbsp;&nbsp;└── ${USER.GOOGLE_PARENT_FOLDER_NAME} (최상위 폴더 이름)</p>
    <p>&nbsp;&nbsp;&nbsp;&nbsp;└── ${USER.GOOGLE_FAIR_FOLDER_NAME} (스프레드시트 이름과 동일해야 함)</p>
    <p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;└── 작가명_파일들.pdf (여기에 작가별 PDF가 저장됨)</p>
    <br>
    <p>구글 드라이브로 이동하려면 <a href="https://drive.google.com" target="_blank">여기</a>를 클릭하세요.</p>
    </div>`
  );

  htmlOutput.setWidth(400).setHeight(300);
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(htmlOutput, '폴더를 찾을 수 없습니다.');
}

/**
 * 구글 드라이브에 두 폴더가 존재하는지만 체크 후 ui 알림 (데이터 반환 X)
 * @returns {boolean}
 */
function drive_checkFolderExistence() {
  const ui = SpreadsheetApp.getUi();
  const parentFolders = DriveApp.getFoldersByName(USER.GOOGLE_PARENT_FOLDER_NAME);

  if (!parentFolders.hasNext()) {
    showFolderNotFoundDialog(USER.GOOGLE_PARENT_FOLDER_NAME);
    return;
  };

  const parentFolder = parentFolders.next();
  const fairFolders = parentFolder.getFoldersByName(USER.GOOGLE_FAIR_FOLDER_NAME);
  if(!fairFolders.hasNext()){
    showFolderNotFoundDialog(USER.GOOGLE_FAIR_FOLDER_NAME);
    return;
  }
  ui.alert(`✅ 구글 드라이브에서 '${USER.GOOGLE_PARENT_FOLDER_NAME}' 폴더 안에 '${USER.GOOGLE_FAIR_FOLDER_NAME}' 폴더를 확인했습니다.`);
}

/**
 * 스프레드 시트 제목으로 구글 드라이브에서 아트페어 폴더를 검색, 해당 폴더를 반환한다.
 * @returns {Folder} 구글 드라이브에서 찾은 아트페어 폴더
 */
function drive_getTargetFolder() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetName = sheet.getName();

    const parentFolderName = USER.GOOGLE_PARENT_FOLDER_NAME;
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