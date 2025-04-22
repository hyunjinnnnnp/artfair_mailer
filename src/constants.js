
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

const STATUS = {
  SENT: "✅ 전송됨",
  ERROR: "❌ 오류 발생",
  // FORM_TRIGGERED: "폼 트리거 작동함",
  // SEND_FAILED: "전송 오류",
  // PROCESS_ERROR: "전체 처리 오류"
};
