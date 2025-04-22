const parentFolderName = "아트페어_PDF";


/**
 * ✅ PDF 파일은 Google 드라이브 내의 특정 폴더에 저장되어야 하며,
 *    폴더 구조와 이름은 아래와 같은 규칙을 따라야 합니다.
 *
 * ✅ 폴더 구조 예시:
 * Google 드라이브
 * └── 아트페어_PDF              ← (최상위 폴더 이름: 고정)
 *     └── 아트파리_2025         ← (아트페어별 폴더: 스프레드시트 이름과 동일해야 함)
 *         └── 작가명_파일들.pdf  ← (여기에 작가별 PDF가 저장됨)
 *
 * ✅ 규칙 요약:
 * 1. 최상위 폴더 이름은 반드시 "아트페어_PDF"여야 합니다.
 * 2. 그 안에 있는 하위 폴더 이름은, 사용 중인 스프레드시트 이름과 정확히 같아야 합니다.
 *    예: 스프레드시트 이름이 "아트파리_2025"면 → 드라이브 내에도 동일한 이름의 폴더가 있어야 합니다.
 * 3. 각 PDF 파일의 이름은 작가명으로 저장되어야 하며, 작가명 철자는 정확하게 입력해야 합니다.
 *
 * ✍️ 예시:
 * 스프레드시트 이름: 아트파리_2025
 * 드라이브 내 저장 위치: 아트페어_PDF/아트파리_2025/이우환.pdf
 */

/**
 * 지역 시간대 설정
 */
const TIMEZONE = {
  PARIS: "Europe/Paris",
  SEOUL: "Asia/Seoul",
  // 다른 지역을 여기에 추가할 수 있습니다
};

/**
 * 스프레드 시트 컬럼 인덱스 - 배열 접근용 (0-based)
 */
const COL_INDEX = {
  TIMESTAMP: 0,
  EMAIL: 1,
  NAME: 2,
  ARTISTS: 3,
  STATUS: 4,
  EMAIL_SENT_AT: 5,
  ERROR: 6,
  MEMO: 7, // 👈 사용자 자유 기입용 메모 칸
};

/**
 * 스프레드 시트 컬럼 넘버
 * sheet.getRange(rowNum, COL_NUM.STATUS) → Range 관련 작업에 사용 (1-based)
 */
const COL_NUM = {};
Object.keys(COL_INDEX).forEach(key => {
  COL_NUM[key] = COL_INDEX[key] + 1;
});
