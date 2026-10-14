/**
 * 모바일 초대장 참석 회신 → 구글 시트 저장
 *
 * [설정 방법]
 * 1. 구글 시트를 새로 만들고, 메뉴 [확장 프로그램] → [Apps Script] 를 엽니다.
 * 2. 기존 코드를 지우고 이 파일 내용을 그대로 붙여넣고 저장합니다.
 * 3. 오른쪽 위 [배포] → [새 배포] → 유형 선택(톱니바퀴) → [웹 앱]
 *      - 다음 사용자 인증 정보로 실행: 나
 *      - 액세스 권한이 있는 사용자: 모든 사용자
 *    → [배포] → 권한 승인 → 나오는 "웹 앱 URL"(…/exec)을 복사합니다.
 * 4. index.html 의 CONFIG.rsvpUrl 에 그 주소를 붙여넣습니다.
 *
 * ※ 코드를 고친 뒤에는 [배포] → [배포 관리] → 연필 아이콘 → 버전 "새 버전" 으로 다시 배포해야 반영됩니다.
 * ※ 행사 담당자에게는 시트만 [공유](보기 또는 편집 권한)로 넘기면 됩니다. 웹앱은 만든 계정(나)으로 계속 실행됩니다.
 * ※ 개인정보 안내대로 행사 종료 후 시트의 회신 내용을 삭제해 주세요.
 */
const SHEET_NAME = '참석회신';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) {
      sh = ss.insertSheet(SHEET_NAME);
      sh.appendRow(['번호', '접수일시', '참석여부', '성함', '소속', '직위/직책', '연락처', '인원', '남기신 말씀', '개인정보 동의']);
      sh.setFrozenRows(1);
      sh.getRange('1:1').setFontWeight('bold').setBackground('#eef4df');
    }
    // 예전 형식(번호 열 없음) 시트면 맨 앞에 '번호' 열을 넣고 기존 회신에 번호를 매김 — 한 번만 실행됨
    if (sh.getRange(1, 1).getValue() !== '번호') {
      sh.insertColumnBefore(1);
      sh.getRange(1, 1).setValue('번호').setFontWeight('bold').setBackground('#eef4df');
      const n = sh.getLastRow() - 1;
      if (n > 0) sh.getRange(2, 1, n, 1).setValues(Array.from({ length: n }, (_, i) => [i + 1]));
    }
    // 예전 형식(직위/직책 열 없음) 시트면 '소속'(E열) 다음에 열을 넣음 — 한 번만 실행됨
    if (sh.getRange(1, 6).getValue() !== '직위/직책') {
      sh.insertColumnAfter(5);
      sh.getRange(1, 6).setValue('직위/직책').setFontWeight('bold').setBackground('#eef4df');
    }
    const p = e.parameter;
    if (p.agree !== '동의') {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'consent required' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    // 일련번호: 마지막 줄 번호 + 1 (중간 줄을 지워도 번호가 겹치지 않음)
    sh.getRange('A:A').setNumberFormat('0');   // 번호 열이 날짜 형식으로 보이지 않게
    const last = sh.getLastRow(), prev = last > 1 ? sh.getRange(last, 1).getValue() : 0;
    const no = (Number.isInteger(prev) && prev >= 0 ? prev : last - 1) + 1;
    const clean = v => String(v || '').replace(/^[=+\-@]/, "'$&").slice(0, 500); // 수식 주입 방지
    sh.appendRow([
      no,
      new Date(),
      clean(p.attend),
      clean(p.name),
      clean(p.org),
      clean(p.position),
      clean(p.tel),
      Number(p.count) || 0,
      clean(p.message),
      '동의',
    ]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/** 시트 맨 아래에 참석 인원 합계를 보고 싶다면: =SUMIF(참석회신!C:C,"참석",참석회신!H:H)  (다른 탭에 입력) */
