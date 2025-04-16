# artfair_mailer - 아트페어 PDF 메일링 툴

`artfair_mailer`는 구글 스프레드시트를 기반으로 작가의 작품 정보를 포함한 PDF 파일을 이메일로 자동으로 발송하는 구글 Apps Script 코드입니다. 각 아트페어에 맞는 스프레드시트를 사용하여 이메일 발송 상태를 관리하고, 관련 파일을 구글 드라이브에서 자동으로 찾아 이메일로 전송합니다.


## 주요 기능

- 구글 스프레드시트에 작성된 손님 정보를 통해 관심 작가의 PDF 작품을 발송합니다.
- 각 아트페어별로 스프레드시트를 관리하며, 이메일 발송 상태와 발송 시간을 기록합니다.
- 구글 드라이브에서 아트페어 폴더를 검색하여 PDF 파일을 자동으로 첨부하고 발송합니다.
- `클래스프`(Clasp)를 사용해 Google Apps Script 프로젝트를 로컬에서 관리하고 배포할 수 있습니다.




## clasp를 사용한 사용 방법
- 이 과정은 생략하려면 sendMail.js의 코드를 사용하실 구글 스프레드 시트에서 생성된 Gooogle Apps Script에 단순 복붙해서 사용할 수 있습니다.

1. **이 프로젝트를 클론합니다.**

   ```bash
   git clone https://github.com/hyunjinnnnnp/artfair_mailer.git
   ```

2. **프로젝트 디렉토리로 이동합니다.**

    ```bash
    cd artfair_mailer
    ```

3. **clasp(Google Apps Script CLI)를 설치합니다.**

    ```bash
    npm install -g @google/clasp
    ```

4. **clasp를 통해 Google Apps Script 프로젝트를 설정합니다.**
- clasp를 사용하여 Google Apps Script 프로젝트를 연결하고 설정합니다.

    ```bash
    clasp create --title "artfair_mailer" --type sheets
    ```

5. **스크립트를 Google Apps Script에 배포합니다.**
- 로컬에서 수정하고 Google Apps Script 프로젝트에 업로드합니다.

    ```bash
    clasp push
    ```

- Google Apps Script에서 수정된 내역을 로컬에 받아올 수 있습니다.

    ```bash
    clasp pull
    ```

6. **배포된 스크립트의 Apps Script 고유 식별자를 스프레드 시트에 직접 연결된 Apps Script에 연결시켜줍니다.**
- 배포된 스크립트의 고유 식별자 값(설정 > ID)을 복사합니다.
- 사용 할 스프레드 시트 > 화면 상단의 확장 프로그램 > App Script > 라이브러리에 추가


7. **스프레드시트에 직접 연결된 Apps Script 파일에 새로운 스크립트를 만들어 아래 코드를 업로드합니다**
- 좌측 편집기'< >' 파일 추가 후 스크립트 생성

    ```javascript
    function onOpen() {
    artfair_mailer.onOpen();
    }

    function sendArtistPdfsUsingSheetName() {
    artfair_mailer.sendArtistPdfsUsingSheetName();
    }
    ```

8. **스프레드시트에서 사용할 Google Drive 폴더 설정**

- 이메일 발송에 필요한 PDF 파일을 저장할 아트페어_PDF라는 폴더를 구글 드라이브 최상단에 생성합니다.
- <u>하위 폴더(ex: art_paris_2025)를 스프레드시트 이름과 동일하게</u> 생성하고 그 안에 PDF 파일들을 배치합니다.
- <u>각 PDF 파일명은 스프레드 시트에 작성된 작가명과 일치해야 합니다.</u>


## 사용법

1. **구글 스프레드시트에서 이메일 발송을 원하는 아트페어 스프레드시트를 엽니다.**
- 스프레드시트에서 각 작가와 이메일 정보, PDF 파일 목록을 입력합니다.
- 이 과정은 구글 폼을 이용해서 손님에게 직접 입력하도록 요구할 수 있습니다.

2. **🖼 갤러리 도구 메뉴에서 이메일 발송 시작을 클릭합니다.**

3. **이메일 발송을 시작하면, 스프레드시트에 입력된 이메일 주소로 작가의 PDF 파일이 첨부되어 발송됩니다.**
- 이미 발송된 내역은 무시됩니다.
- 발송 상태는 자동으로 업데이트됩니다.

