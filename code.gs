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
 *  9) (선택) 스프레드시트를 새로 열면 상단에 "비율 탐험대" 메뉴가 생겨요.
 *     - "비율 탐험대 → 학생 활동 사이드바" 를 누르면 좌측에 통계 사이드바가 열려요.
 *     - "비율 탐험대 → 대시보드 열기" 를 누르면 큰 팝업 대시보드가 열려요.
 *
 *  완료! 학생들이 비율탐험대 앱을 열면 이름 선택 화면이 나오고,
 *  활동 기록이 "활동기록" 탭에 자동으로 쌓입니다.
 */

const SHEET_ID = '1oBZ2Bslv7WV6ECYsgFEy8QeVDfOgOsc46AYIHDtddE4';

// ============ 메뉴 자동 생성 ============
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('비율 탐험대')
    .addItem('📊 학생 활동 사이드바', 'showSidebar')
    .addItem('📈 대시보드 열기', 'showDashboard')
    .addSeparator()
    .addItem('🔄 통계 새로고침', 'refreshStats')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutput(getSidebarHtml())
    .setTitle('비율 탐험대 — 학생 활동')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showDashboard() {
  const html = HtmlService.createHtmlOutput(getDashboardHtml())
    .setWidth(1100)
    .setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, '📈 비율 탐험대 — 대시보드');
}

function refreshStats() {
  SpreadsheetApp.getActiveSpreadsheet().toast('통계가 새로고침되었어요.', '비율 탐험대', 3);
}

// ============ 웹 앱 ============
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
    if (action === 'stats') {
      return jsonResponse({ ok: true, stats: getStats() });
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

// ============ 통계 ============
function getStats() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const logSheet = ss.getSheetByName('활동기록');
  const nameSheet = ss.getSheetByName('학생명단');

  const allNames = nameSheet
    ? nameSheet.getRange('A2:A' + Math.max(2, nameSheet.getLastRow())).getValues().flat().filter(function (v) { return v && String(v).trim(); })
    : [];

  const byStudent = {};
  allNames.forEach(function (n) {
    byStudent[n] = { name: n, total: 0, correct: 0, logins: 0, stages: { s1: 0, s2: 0, s3: 0 }, badges: 0, lastSeen: null };
  });

  let totalEvents = 0;
  let totalCorrect = 0;
  let totalLogins = 0;
  let recent = [];

  if (logSheet && logSheet.getLastRow() > 1) {
    const rows = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 6).getValues();
    rows.forEach(function (r) {
      const when = r[0], name = r[1], event = r[2], stage = r[3], stars = r[4], detail = r[5];
      if (!name) return;
      if (!byStudent[name]) {
        byStudent[name] = { name: name, total: 0, correct: 0, logins: 0, stages: { s1: 0, s2: 0, s3: 0 }, badges: 0, lastSeen: null };
      }
      const s = byStudent[name];
      s.total += 1;
      totalEvents += 1;
      if (event === 'correct') { s.correct += 1; totalCorrect += 1; }
      if (event === 'login') { s.logins += 1; totalLogins += 1; }
      if (event === 'stage_complete' && stage && s.stages[stage] !== undefined) {
        s.stages[stage] = Math.max(s.stages[stage], Number(stars) || 0);
      }
      if (event === 'badge') s.badges += 1;
      if (when instanceof Date) {
        if (!s.lastSeen || when > s.lastSeen) s.lastSeen = when;
      }
      recent.push({ when: when instanceof Date ? when.toISOString() : String(when), name: name, event: event, stage: stage, stars: stars, detail: detail });
    });
  }

  recent = recent.slice(-30).reverse();

  const students = Object.keys(byStudent).map(function (k) {
    const s = byStudent[k];
    return {
      name: s.name,
      total: s.total,
      correct: s.correct,
      logins: s.logins,
      stages: s.stages,
      stars: (s.stages.s1 || 0) + (s.stages.s2 || 0) + (s.stages.s3 || 0),
      badges: s.badges,
      lastSeen: s.lastSeen ? s.lastSeen.toISOString() : null,
    };
  }).sort(function (a, b) { return b.correct - a.correct; });

  return {
    totals: { events: totalEvents, correct: totalCorrect, logins: totalLogins, students: allNames.length },
    students: students,
    recent: recent,
  };
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============ Sidebar HTML ============
function getSidebarHtml() {
  return '<!DOCTYPE html><html><head><base target="_top">'
    + '<style>'
    + 'body{font-family:"Apple SD Gothic Neo","Noto Sans KR",sans-serif;background:#FFF6E5;color:#2D1810;padding:14px;margin:0;}'
    + 'h2{margin:0 0 8px;color:#92400E;font-size:18px;}'
    + '.totals{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px;}'
    + '.box{background:white;border:2px solid #FCD34D;border-radius:10px;padding:8px 10px;flex:1;min-width:90px;text-align:center;}'
    + '.box b{display:block;font-size:20px;color:#B45309;}'
    + '.box span{font-size:12px;color:#78716C;}'
    + 'table{width:100%;border-collapse:collapse;background:white;border-radius:8px;overflow:hidden;border:2px solid #FED7AA;}'
    + 'th,td{padding:6px 8px;text-align:left;border-bottom:1px solid #FED7AA;font-size:12px;}'
    + 'th{background:#FEF3C7;color:#9A3412;}'
    + 'tr:last-child td{border-bottom:none;}'
    + '.row-empty{padding:14px;text-align:center;color:#78716C;}'
    + '.stars{color:#F59E0B;}'
    + '.refresh{background:#F59E0B;color:white;border:none;border-radius:8px;padding:6px 12px;font-size:13px;cursor:pointer;margin-bottom:10px;}'
    + '.refresh:hover{background:#D97706;}'
    + '.muted{color:#78716C;font-size:11px;margin-top:6px;}'
    + '</style></head><body>'
    + '<h2>📊 학생 활동 요약</h2>'
    + '<button class="refresh" onclick="load()">🔄 새로고침</button>'
    + '<div id="content">불러오는 중...</div>'
    + '<script>'
    + 'function fmtStars(n){var s="★".repeat(n||0)+"☆".repeat(Math.max(0,3-(n||0)));return s;}'
    + 'function render(d){'
    + 'var t=d.totals;'
    + 'var html="<div class=\\"totals\\">"'
    + '+"<div class=\\"box\\"><b>"+t.students+"</b><span>학생</span></div>"'
    + '+"<div class=\\"box\\"><b>"+t.correct+"</b><span>총 정답</span></div>"'
    + '+"<div class=\\"box\\"><b>"+t.logins+"</b><span>로그인</span></div>"'
    + '+"</div>";'
    + 'if(!d.students.length){html+="<div class=\\"row-empty\\">학생 활동이 아직 없어요</div>";}'
    + 'else{html+="<table><tr><th>이름</th><th>정답</th><th>별</th><th>1</th><th>2</th><th>3</th></tr>";'
    + 'for(var i=0;i<d.students.length;i++){var s=d.students[i];'
    + 'html+="<tr><td>"+s.name+"</td><td>"+s.correct+"</td><td>"+s.stars+"</td>"'
    + '+"<td class=\\"stars\\">"+fmtStars(s.stages.s1)+"</td>"'
    + '+"<td class=\\"stars\\">"+fmtStars(s.stages.s2)+"</td>"'
    + '+"<td class=\\"stars\\">"+fmtStars(s.stages.s3)+"</td></tr>";}'
    + 'html+="</table>";}'
    + 'html+="<div class=\\"muted\\">마지막 갱신: "+new Date().toLocaleString()+"</div>";'
    + 'document.getElementById("content").innerHTML=html;}'
    + 'function load(){document.getElementById("content").innerText="불러오는 중...";'
    + 'google.script.run.withSuccessHandler(render).withFailureHandler(function(e){document.getElementById("content").innerText="오류: "+e.message;}).getStats();}'
    + 'load();'
    + '</script></body></html>';
}

// ============ Dashboard HTML (큰 팝업) ============
function getDashboardHtml() {
  return '<!DOCTYPE html><html><head><base target="_top">'
    + '<style>'
    + 'body{font-family:"Apple SD Gothic Neo","Noto Sans KR",sans-serif;background:linear-gradient(135deg,#FFF6E5,#FDE8B5);color:#2D1810;padding:24px;margin:0;}'
    + 'h1{margin:0 0 4px;color:#92400E;font-size:28px;}'
    + '.sub{color:#78716C;margin-bottom:16px;font-size:14px;}'
    + '.row{display:flex;gap:14px;flex-wrap:wrap;}'
    + '.col{flex:1;min-width:280px;}'
    + '.card{background:white;border:3px solid #F4B942;border-radius:18px;padding:18px;box-shadow:0 6px 18px rgba(74,44,31,0.1);}'
    + '.totals{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:14px;}'
    + '.box{background:white;border:3px solid #FCD34D;border-radius:14px;padding:14px 18px;flex:1;min-width:110px;text-align:center;}'
    + '.box b{display:block;font-size:32px;color:#B45309;font-weight:bold;}'
    + '.box span{font-size:13px;color:#78716C;}'
    + 'table{width:100%;border-collapse:collapse;}'
    + 'th,td{padding:8px 10px;text-align:left;border-bottom:1px solid #FED7AA;font-size:13px;}'
    + 'th{background:#FEF3C7;color:#9A3412;position:sticky;top:0;}'
    + '.stars{color:#F59E0B;font-size:14px;}'
    + 'h2{margin:0 0 10px;color:#9A3412;font-size:18px;}'
    + '.bar-wrap{background:#FEF3C7;border-radius:6px;overflow:hidden;height:14px;}'
    + '.bar{background:linear-gradient(90deg,#F59E0B,#EA580C);height:100%;}'
    + '.refresh{background:#F59E0B;color:white;border:none;border-radius:10px;padding:8px 16px;font-size:14px;cursor:pointer;margin-bottom:10px;float:right;}'
    + '.refresh:hover{background:#D97706;}'
    + '.recent{max-height:280px;overflow-y:auto;}'
    + '.event{padding:6px 10px;border-bottom:1px solid #FED7AA;font-size:12px;display:flex;justify-content:space-between;gap:8px;}'
    + '.event:last-child{border-bottom:none;}'
    + '.ev-name{font-weight:bold;color:#9A3412;}'
    + '.ev-time{color:#78716C;font-size:11px;}'
    + '.muted{color:#78716C;font-size:12px;}'
    + '</style></head><body>'
    + '<button class="refresh" onclick="load()">🔄 새로고침</button>'
    + '<h1>📈 비율 탐험대 대시보드</h1>'
    + '<div class="sub" id="updated">불러오는 중...</div>'
    + '<div id="content"></div>'
    + '<script>'
    + 'function fmtStars(n){return "★".repeat(n||0)+"☆".repeat(Math.max(0,3-(n||0)));}'
    + 'function pct(n,d){if(!d)return 0;return Math.round(n/d*100);}'
    + 'function render(d){'
    + 'var t=d.totals;'
    + 'var maxCorrect=Math.max(1,Math.max.apply(null,d.students.map(function(s){return s.correct;})));'
    + 'var html="<div class=\\"totals\\">"'
    + '+"<div class=\\"box\\"><b>"+t.students+"</b><span>전체 학생</span></div>"'
    + '+"<div class=\\"box\\"><b>"+t.correct+"</b><span>총 정답</span></div>"'
    + '+"<div class=\\"box\\"><b>"+t.events+"</b><span>총 이벤트</span></div>"'
    + '+"<div class=\\"box\\"><b>"+t.logins+"</b><span>로그인 수</span></div>"'
    + '+"</div>";'
    + 'html+="<div class=\\"row\\"><div class=\\"col\\"><div class=\\"card\\"><h2>🏆 학생 순위 & 진행</h2>";'
    + 'html+="<div style=\\"max-height:420px;overflow-y:auto;\\"><table><tr><th>이름</th><th>정답</th><th>막대</th><th>비의숲</th><th>비율</th><th>%</th></tr>";'
    + 'if(!d.students.length){html+="<tr><td colspan=\\"6\\" class=\\"muted\\">아직 활동이 없어요</td></tr>";}'
    + 'else{for(var i=0;i<d.students.length;i++){var s=d.students[i];'
    + 'var w=pct(s.correct,maxCorrect);'
    + 'html+="<tr><td><b>"+s.name+"</b></td><td>"+s.correct+"</td>"'
    + '+"<td><div class=\\"bar-wrap\\"><div class=\\"bar\\" style=\\"width:"+w+"%;\\"></div></div></td>"'
    + '+"<td class=\\"stars\\">"+fmtStars(s.stages.s1)+"</td>"'
    + '+"<td class=\\"stars\\">"+fmtStars(s.stages.s2)+"</td>"'
    + '+"<td class=\\"stars\\">"+fmtStars(s.stages.s3)+"</td></tr>";}}'
    + 'html+="</table></div></div></div>";'
    + 'html+="<div class=\\"col\\"><div class=\\"card\\"><h2>📜 최근 활동 (최근 30건)</h2><div class=\\"recent\\">";'
    + 'if(!d.recent.length){html+="<div class=\\"muted\\">활동이 없어요</div>";}'
    + 'else{for(var j=0;j<d.recent.length;j++){var r=d.recent[j];'
    + 'var when=r.when?new Date(r.when).toLocaleString():"";'
    + 'var label=r.event;if(r.event==="stage_complete")label="단계 클리어("+r.stage+", "+r.stars+"★)";'
    + 'if(r.event==="correct")label="정답!";'
    + 'if(r.event==="login")label="로그인";'
    + 'if(r.event==="badge")label="메달: "+(r.detail||"");'
    + 'if(r.event==="milestone")label="레벨: "+(r.detail||"");'
    + 'if(r.event==="download_note")label="탐험노트 받기";'
    + 'html+="<div class=\\"event\\"><span><span class=\\"ev-name\\">"+r.name+"</span> · "+label+"</span><span class=\\"ev-time\\">"+when+"</span></div>";}}'
    + 'html+="</div></div></div></div>";'
    + 'document.getElementById("content").innerHTML=html;'
    + 'document.getElementById("updated").innerText="마지막 갱신: "+new Date().toLocaleString();}'
    + 'function load(){document.getElementById("updated").innerText="불러오는 중...";'
    + 'google.script.run.withSuccessHandler(render).withFailureHandler(function(e){document.getElementById("updated").innerText="오류: "+e.message;}).getStats();}'
    + 'load();'
    + '</script></body></html>';
}
