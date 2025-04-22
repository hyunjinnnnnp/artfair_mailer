
/**
 * 스프레드시트의 고객 정보 수정 제한
 */
function protectColumns() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const protection = sheet.protect().setDescription('고객 정보 보호');

  // Ensure the current user is an editor before removing others. Otherwise, if the user's edit permission comes from a group, the script throws an exception upon removing the group.
  const me = Session.getEffectiveUser();
  protection.setWarningOnly(false); // 설정해야만 add, remove editor 사용 가능
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
    // 👉 조직(예: 같은 회사 도메인)의 다른 사람들이 수정할 수 있는 기능 off
  }
}
