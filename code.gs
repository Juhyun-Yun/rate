/**
 * ============================================
 *  비율 탐험대 - Google Apps Script 백엔드
 * ============================================
 *
 *  ▶ 사용 방법 (선생님용)
 *
 *  1) Google Sheets에서 새 시트 만들기
 *
 *  2) 첫 번째 탭 이름을 "학생명단" 으로 바꾸기
 *     - A1 셀에 "이름" 입력
 *     - A2 부터 학생 이름 한 줄에 한 명씩 입력
 *
 *  3) 두 번째 탭 만들기, 이름을 "활동기록" 으로 바꾸기
 *     - 첫 행에 헤더 입력:
 *       A1: 날짜시간 | B1: 학생이름 | C1: 이벤트 | D1: 단계 | E1: 점수 | F1: 상세
 *
 *  4) 시트 ID 복사하기
 *     - 시트 URL: https://docs.google.com/spreadsheets/d/[시트ID]/edit#gid=0
 *     - "[시트ID]" 부분이 ID
 *
 *  5) 아래 SHEET_ID 값을 시트 ID로 바꿔주세요
 *
 *  6) Sheets 메뉴 → 확장 프로그램 → Apps Script 클릭
 *     - 기본 코드 모두 지우고 이 파일 내용을 붙여넣기
 *     - 저장
 *
 *  7) 배포하기
 *     - 우측 상단 "배포" → "새 배포"
 *     - 톱니바퀴 → 유형: "웹 앱" 선택
 *     - 다음 사람으로 실행: "나"
 *     - 액세스 권한: "모든 사용자"
 *     - "배포" 클릭
 *     - 권한 검토 → 본인 계정으로 허용
 *     - "웹 앱 URL"을 복사
 *
 *  8) 복사한 URL을 index.html 의 SCRIPT_URL 변수에 붙여넣기
 *     - index.html을 메모장으로 열기
 *     - 맨 위쪽에 const SCRIPT_URL = ''; 줄을 찾기
 *     - 따옴표 안에 URL 붙여넣기
 *
 *  완료! 학생들이 비율탐험대 앱을 열면 이름 선택 화면이 나오고,
 *  활동 기록이 "활동기록" 탭에 자동으로 쌓입니다.
 */

const SHEET_ID = '1oBZ2Bslv7WV6ECYsgFEy8QeVDfOgOsc46AYIHDtddE4';

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'list';
    if (action === 'list') {
      return jsonResponse({ ok: true, names: getStudentNames() });
    }
    if (action === 'log') {
      logActivity(e.parameter);
      return jsonResponse({ ok: true });
    }
    return jsonResponse({ error: 'unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'log') {
      logActivity(data);
      return jsonResponse({ ok: true });
    }
    return jsonResponse({ error: 'unknown action: ' + data.action });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function getStudentNames() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('학생명단');
  if (!sheet) throw new Error('"학생명단" 탭을 찾을 수 없어요. 탭 이름을 확인하세요.');
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange('A2:A' + last).getValues().flat().filter(function (v) { return v && String(v).trim(); });
}

function logActivity(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('활동기록');
  if (!sheet) {
    sheet = ss.insertSheet('활동기록');
    sheet.appendRow(['날짜시간', '학생이름', '이벤트', '단계', '점수', '상세']);
  }
  sheet.appendRow([
    new Date(),
    data.name || '',
    data.event || '',
    data.stage || '',
    data.stars || '',
    data.detail || ''
  ]);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
