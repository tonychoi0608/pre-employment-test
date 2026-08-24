/**
 * 고강이엔지 인적성검사 — 결과 수신 서버 (Google Apps Script)
 *
 * 하는 일:
 *  1) 검사 페이지가 보낸 결과를 받아 구글 스프레드시트에 한 줄씩 기록
 *  2) 결과 전문(AI 분석 프롬프트 + 데이터)을 txt 파일로 첨부해 이메일 발송
 *
 * 설치 방법은 같은 폴더의 설정방법.md 참고.
 */

var TO_EMAIL = 'gokang.korea@gmail.com';        // 결과를 받을 메일 주소
var TOKEN = 'gk-1379';                           // 검사 페이지와 맞춰 둔 확인용 코드
var SHEET_NAME = '고강이엔지 인적성검사 결과';   // 자동 생성될 스프레드시트 이름

function doGet() {
  return ContentService.createTextOutput('ok');
}

function doPost(e) {
  var out = { ok: false };
  try {
    var d = JSON.parse(e.postData.contents);
    if (!d || d.token !== TOKEN) {
      out.err = 'bad token';
      return json_(out);
    }
    var when = d.ts ? new Date(d.ts) : new Date();
    var s = d.summary || {};
    var mailErr = '';
    try {
      sendMail_(d, s, when);
    } catch (err) {
      mailErr = String(err);
    }
    try {
      appendRow_(d, s, when, mailErr ? '실패: ' + mailErr : '발송됨');
    } catch (err2) {
      out.err = 'sheet: ' + String(err2) + (mailErr ? ' / mail: ' + mailErr : '');
      out.ok = !mailErr; // 메일이라도 나갔으면 접수는 된 것으로 처리
      return json_(out);
    }
    out.ok = true;
    if (mailErr) out.warn = 'mail: ' + mailErr;
  } catch (err3) {
    out.err = String(err3);
  }
  return json_(out);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var props = PropertiesService.getScriptProperties();
    var id = props.getProperty('sheetId');
    var ss = null;
    if (id) {
      try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
    }
    if (!ss) {
      ss = SpreadsheetApp.create(SHEET_NAME);
      props.setProperty('sheetId', ss.getId());
      ss.getSheets()[0].appendRow(['접수일시', '이름', '지원직무', '경력', '소요시간', '신뢰도',
        '자기보고(100점)', '상황판단 평균', '실무감각', '타고난 동기', '동기 괴리', '메일']);
      ss.getSheets()[0].setFrozenRows(1);
    }
    return ss;
  } finally {
    lock.releaseLock();
  }
}

function appendRow_(d, s, when, mailStat) {
  var ss = getSheet_();
  ss.getSheets()[0].appendRow([
    Utilities.formatDate(when, 'Asia/Seoul', 'yyyy-MM-dd HH:mm'),
    d.name || '', d.role || '', d.career || '', d.dur || '', d.grade || '',
    s.self || '', s.sjt != null ? s.sjt : '', s.mech || '', s.drive || '', s.gap || '',
    mailStat
  ]);
}

function sendMail_(d, s, when) {
  var dateStr = Utilities.formatDate(when, 'Asia/Seoul', 'yyyy-MM-dd');
  var subject = '[인적성검사] ' + (d.name || '이름없음') + ' · ' + (d.role || '') + ' · ' + dateStr;
  var sheetUrl = '';
  try { sheetUrl = getSheet_().getUrl(); } catch (e) {}
  var body =
    '고강이엔지 인적성검사 결과가 접수되었습니다.\n\n' +
    '■ 지원자: ' + (d.name || '') + ' (' + (d.role || '') + ' / ' + (d.career || '') + ')\n' +
    '■ 응시일시: ' + Utilities.formatDate(when, 'Asia/Seoul', 'yyyy-MM-dd HH:mm') + '\n' +
    '■ 소요시간: ' + (d.dur || '') + '\n' +
    '■ 신뢰도: ' + (d.grade || '') + '\n' +
    '■ 자기보고(100점): ' + (s.self || '') + '\n' +
    '■ 상황판단 평균: ' + (s.sjt != null ? s.sjt : '') + '\n' +
    '■ 실무감각: ' + (s.mech || '') + '\n' +
    '■ 타고난 동기: ' + (s.drive || '') + '\n' +
    '■ 동기 괴리: ' + (s.gap || '') + '\n\n' +
    '▶ AI 분석 방법\n' +
    '첨부된 텍스트 파일을 열어 내용 전체를 복사한 뒤,\n' +
    'Claude(claude.ai) 대화창에 붙여넣으면 분석 리포트가 나옵니다.\n\n' +
    (sheetUrl ? '▶ 전체 접수 현황(스프레드시트)\n' + sheetUrl + '\n' : '');
  var fname = '인적성검사_' + (d.name || '지원자') + '_' + dateStr + '.txt';
  var options = {};
  if (d.package) {
    options.attachments = [Utilities.newBlob(d.package, 'text/plain', fname)];
  }
  MailApp.sendEmail(TO_EMAIL, subject, body, options);
}

/** 편집기에서 이 함수를 실행하면 설정이 잘 됐는지 테스트 메일이 갑니다. */
function testMail() {
  sendMail_(
    { name: '테스트', role: '생산직', career: '신입', dur: '0분 0초', grade: 'A',
      package: '테스트 첨부입니다. 이 메일이 보이면 설정 성공!' },
    { self: '테스트', sjt: 0, mech: '0/6', drive: '-', gap: '-' },
    new Date()
  );
}
