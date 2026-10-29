/**
 * 2026 직원 건강관리 설문조사 (본사용) - 응답 수집 백엔드
 *
 * 사용 방법
 * 1) 본사용 응답을 모을 구글 스프레드시트를 열고 [확장 프로그램] > [Apps Script] 클릭
 * 2) 기본 생성된 코드를 지우고 이 파일 내용 전체를 붙여넣기
 * 3) 저장 (Ctrl+S)
 * 4) 상단 [배포] > [새 배포] > 유형: "웹 앱" 선택
 *    - 실행 사용자: 나
 *    - 액세스 권한이 있는 사용자: 전체 허용(Anyone)  ← 로그인 없이 깃허브 페이지에서 호출 가능해야 하므로 반드시 "Anyone"으로 설정
 * 5) 배포 후 나오는 웹 앱 URL을 복사해서 hq/index.html 상단의 SCRIPT_URL 값에 붙여넣기
 */

// 본사 설문 응답 전용 스프레드시트 (매장용과는 완전히 분리된 시트)
const SPREADSHEET_ID = '1NL1oZeLXY6Vvd_mv7oNYac53UTiFHN3hCCmtGIpxPos';
const SHEET_NAME = '본사_응답';

const HEADERS = [
  '제출시각',
  '구분',
  'Q1_건강관리활동_도움정도(1-5)',
  'Q2_건강개선여부',
  'Q2_개선된항목',
  'Q3_회사지원인식(1-5)',
  'Q4_만족활동',
  'Q5_희망프로그램',
  'Q5_기타의견'
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet_();

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
    } else {
      // 설문 문항이 수정되어 헤더 구성이 바뀐 경우, 기존 헤더 행을 최신 구성으로 갱신
      const currentHeaderRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
      const isSame = HEADERS.every((h, i) => currentHeaderRow[i] === h);
      if (!isSame) {
        sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      }
    }

    sheet.appendRow([
      new Date(),
      data.site || '본사',
      data.q1 || '',
      data.q2 || '',
      (data.q2_sub || []).join(', '),
      data.q3 || '',
      (data.q4 || []).join(', '),
      (data.q5 || []).join(', '),
      data.q5_other || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: '본사용 설문 제출 웹앱이 정상 동작 중입니다.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  return sheet;
}
