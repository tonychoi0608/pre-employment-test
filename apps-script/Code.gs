/**
 * 고강이엔지 인적성검사 — 결과 수신 서버 (Google Apps Script)
 *
 * 하는 일:
 *  1) 검사 페이지가 보낸 결과를 구글 스프레드시트에 한 줄씩 기록 (요약 + 원본 데이터 전체)
 *  2) 메일에서 바로 읽을 수 있는 결과지(HTML)를 만들어 발송
 *  3) 관리자가 검사 페이지에서 담당자 코드를 입력하면, 이 시트에 쌓인 결과를
 *     어떤 기기(폰 포함)에서든 목록/상세로 조회할 수 있게 하는 조회 API(doGet)
 *
 * 설치 방법은 같은 폴더의 설정방법.md 참고.
 * ⚠ 이 파일을 고치면 반드시 "배포 → 배포 관리 → 편집 → 새 버전"으로 다시 배포해야 반영됩니다.
 */

var TO_EMAIL = 'gokang.korea@gmail.com';        // 결과를 받을 메일 주소
var TOKEN = 'gk-1379';                           // 검사 페이지와 맞춰 둔 확인용 코드
var SHEET_NAME = '고강이엔지 인적성검사 결과';   // 자동 생성될 스프레드시트 이름
var MAX_CELL = 45000;                            // 시트 한 셀에 담는 텍스트의 안전 길이 한도

/* 색상 */
var C_INK = '#1B2430', C_MUTED = '#62707E', C_LINE = '#DCE1E5';
var C_BG = '#F3F4F2', C_TRACK = '#E6E9E7', C_STEEL = '#2E4057';
var C_GOOD = '#2E7D46', C_MID = '#2E4057', C_WARN = '#B07A22', C_BAD = '#B3402A';

/* 시트 컬럼 순서. 예전 시트에 없던 컬럼(뒤 3개)은 열 때 자동으로 추가됩니다. */
var HEADERS = ['접수일시', '이름', '지원직무', '경력', '소요시간', '신뢰도',
  '자기보고(100점)', '상황판단 평균', '실무감각', '타고난 동기', '동기 괴리', '메일',
  '접수ID', '원본데이터', '패키지'];

/* ================= 진입점 ================= */

/** 관리자 조회용. ?action=list 또는 ?action=get&id=... + token 필요. 없으면 헬스체크. */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var cb = p.callback;
  function respond(obj) {
    var body = JSON.stringify(obj);
    if (cb) {
      return ContentService.createTextOutput(cb + '(' + body + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return json_(obj);
  }
  if (!p.action) return respond({ ok: true });
  if (p.token !== TOKEN) return respond({ ok: false, err: 'unauthorized' });
  try {
    if (p.action === 'list') return respond({ ok: true, items: listResults_() });
    if (p.action === 'get') return respond({ ok: true, item: getResult_(p.id) });
    return respond({ ok: false, err: 'unknown action' });
  } catch (err) {
    return respond({ ok: false, err: String(err) });
  }
}

/** 검사 완료 시 결과 수신. */
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
function trim_(t) {
  t = String(t == null ? '' : t);
  return t.length > MAX_CELL ? t.slice(0, MAX_CELL) + '\n…(길이 제한으로 생략됨 — 전체 내용은 메일 첨부 파일 참고)' : t;
}

/* ================= 스프레드시트 ================= */
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
      ss.getSheets()[0].appendRow(HEADERS);
      ss.getSheets()[0].setFrozenRows(1);
    } else {
      migrateHeaders_(ss.getSheets()[0]);
    }
    return ss;
  } finally {
    lock.releaseLock();
  }
}

/** 예전에 만들어진 시트에 새 컬럼(접수ID/원본데이터/패키지)이 없으면 헤더를 맞춰 확장합니다. */
function migrateHeaders_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var cur = sheet.getRange(1, 1, 1, Math.max(lastCol, HEADERS.length)).getValues()[0];
  var need = false;
  for (var i = 0; i < HEADERS.length; i++) {
    if (cur[i] !== HEADERS[i]) { need = true; break; }
  }
  if (need) sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
}

function appendRow_(d, s, when, mailStat) {
  var ss = getSheet_();
  ss.getSheets()[0].appendRow([
    Utilities.formatDate(when, 'Asia/Seoul', 'yyyy-MM-dd HH:mm'),
    d.name || '', d.role || '', d.career || '', d.dur || '', d.grade || '',
    s.self || '', s.sjt != null ? s.sjt : '', s.mech || '', s.drive || '', s.gap || '',
    mailStat,
    d.ts != null ? String(d.ts) : '',
    d.data ? trim_(JSON.stringify(d.data)) : '',
    d.package ? trim_(d.package) : ''
  ]);
}

/* ================= 관리자 조회 API ================= */
function sheetRows_() {
  var sheet = getSheet_().getSheets()[0];
  var last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
}
function colIdx_(name) { return HEADERS.indexOf(name); }

/** 목록 조회용 — 가벼운 요약만 반환(원본데이터·패키지는 제외), 최신순. */
function listResults_() {
  var rows = sheetRows_();
  var iId = colIdx_('접수ID'), iWhen = colIdx_('접수일시'), iName = colIdx_('이름'),
    iRole = colIdx_('지원직무'), iCareer = colIdx_('경력'), iDur = colIdx_('소요시간'),
    iGrade = colIdx_('신뢰도'), iSelf = colIdx_('자기보고(100점)'), iSjt = colIdx_('상황판단 평균'),
    iMech = colIdx_('실무감각'), iDrive = colIdx_('타고난 동기'), iGap = colIdx_('동기 괴리'),
    iData = colIdx_('원본데이터');
  var items = [];
  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];
    if (!row[iName]) continue;
    items.push({
      id: row[iId] ? String(row[iId]) : ('row' + r),
      when: row[iWhen], name: row[iName], role: row[iRole], career: row[iCareer],
      dur: row[iDur], grade: row[iGrade], self: row[iSelf], sjt: row[iSjt],
      mech: row[iMech], drive: row[iDrive], gap: row[iGap],
      hasDetail: !!row[iData]
    });
  }
  items.reverse();
  return items;
}

/** 상세 조회용 — 접수ID로 해당 행의 원본데이터(JSON)와 패키지(AI 분석용 전체 텍스트)를 반환. */
function getResult_(id) {
  if (!id) throw new Error('id가 필요합니다');
  var rows = sheetRows_();
  var iId = colIdx_('접수ID'), iData = colIdx_('원본데이터'), iPkg = colIdx_('패키지'),
    iName = colIdx_('이름'), iWhen = colIdx_('접수일시');
  for (var r = rows.length - 1; r >= 0; r--) {
    var row = rows[r];
    var rowId = row[iId] ? String(row[iId]) : ('row' + r);
    if (rowId === String(id)) {
      if (!row[iData]) throw new Error('상세 데이터가 없습니다 (이 기능 도입 전에 접수된 과거 결과입니다 — 검사 PC에서 "클라우드로 전송" 버튼을 한 번 눌러 다시 보내 주세요)');
      return { data: JSON.parse(row[iData]), package: row[iPkg] || '', name: row[iName], when: row[iWhen] };
    }
  }
  throw new Error('해당 접수 건을 찾을 수 없습니다');
}

/* ================= 메일 ================= */
function sendMail_(d, s, when) {
  var dateStr = Utilities.formatDate(when, 'Asia/Seoul', 'yyyy-MM-dd');
  var subject = '[인적성검사] ' + (d.name || '이름없음') + ' · ' + (d.role || '') + ' · ' + dateStr;
  var sheetUrl = '';
  try { sheetUrl = getSheet_().getUrl(); } catch (e) {}

  var plain =
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
    (sheetUrl ? '전체 접수 현황: ' + sheetUrl + '\n' : '');

  var options = {};
  if (d.package) {
    options.attachments = [Utilities.newBlob(d.package, 'text/plain',
      '인적성검사_' + (d.name || '지원자') + '_' + dateStr + '.txt')];
  }
  if (d.data) {
    try { options.htmlBody = buildReportHtml_(d, s, when, sheetUrl); } catch (e) {}
  }
  options.name = '고강이엔지 인적성검사';
  MailApp.sendEmail(TO_EMAIL, subject, plain, options);
}

/* ================= 결과지 HTML ================= */
function esc_(t) {
  return String(t == null ? '' : t)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function nl2br_(t) { return esc_(t).replace(/\n/g, '<br>'); }

function scoreColor_(v) {
  if (v >= 75) return C_GOOD;
  if (v >= 60) return C_MID;
  if (v >= 45) return C_WARN;
  return C_BAD;
}
function gradeColor_(g) {
  if (g === '양호') return C_GOOD;
  if (g === '주의') return C_WARN;
  if (g === '낮음') return C_BAD;
  return C_STEEL;
}

/** 가로 막대 한 줄 */
function bar_(label, v, color, note) {
  var w = Math.max(1, Math.min(100, Math.round(v)));
  return '<tr>' +
    '<td style="padding:7px 10px 7px 0;font-size:14px;color:' + C_INK + ';white-space:nowrap;">' + esc_(label) + '</td>' +
    '<td style="padding:7px 0;width:100%;">' +
      '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:separate;background:' + C_TRACK + ';border-radius:7px;">' +
        '<tr><td width="' + w + '%" style="background:' + color + ';height:13px;line-height:13px;font-size:1px;border-radius:7px;">&nbsp;</td>' +
        '<td style="font-size:1px;line-height:13px;">&nbsp;</td></tr>' +
      '</table>' +
    '</td>' +
    '<td align="right" style="padding:7px 0 7px 12px;font-size:15px;font-weight:700;color:' + color + ';white-space:nowrap;">' + Math.round(v) +
      (note ? '<span style="font-size:12px;font-weight:400;color:' + C_MUTED + ';"> ' + esc_(note) + '</span>' : '') +
    '</td></tr>';
}

function section_(title, inner, sub) {
  return '<tr><td style="padding:26px 28px 0 28px;">' +
    '<div style="font-size:12px;letter-spacing:.08em;color:' + C_MUTED + ';font-weight:700;text-transform:uppercase;">' + esc_(title) + '</div>' +
    (sub ? '<div style="font-size:12px;color:' + C_MUTED + ';margin-top:3px;">' + esc_(sub) + '</div>' : '') +
    '<div style="height:10px;"></div>' + inner +
  '</td></tr>';
}

function chip_(text, bg, fg) {
  return '<span style="display:inline-block;background:' + bg + ';color:' + fg +
    ';font-size:12px;font-weight:700;padding:4px 10px;border-radius:20px;">' + esc_(text) + '</span>';
}

function buildReportHtml_(d, s, when, sheetUrl) {
  var x = d.data;
  var meta = x.검사정보 || {}, info = x.지원자정보 || {}, rel = x.신뢰도지표 || {};
  var sum = x.점수요약_100점 || {};
  var combined = sum.종합_자기보고60_상황판단40 || {};
  var selfS = sum.자기보고 || {};
  var sjt = sum.상황판단 || {};
  var motive = sum.동기프로필 || {};
  var h = [];

  h.push('<div style="background:' + C_BG + ';padding:20px 0;margin:0;">');
  h.push('<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:' + C_BG + ';">');
  h.push('<tr><td align="center" style="padding:0 10px;">');
  h.push('<table width="640" cellpadding="0" cellspacing="0" border="0" style="max-width:640px;width:100%;background:#FFFFFF;border-radius:14px;overflow:hidden;font-family:-apple-system,BlinkMacSystemFont,\'Malgun Gothic\',\'맑은 고딕\',sans-serif;color:' + C_INK + ';">');

  /* 헤더 */
  h.push('<tr><td style="background:' + C_STEEL + ';padding:22px 28px;">' +
    '<div style="color:#9FB3C8;font-size:11px;letter-spacing:.14em;font-weight:700;">GOKANG ENG · 인적성검사 결과지</div>' +
    '<div style="color:#FFFFFF;font-size:26px;font-weight:800;margin-top:6px;">' + esc_(info.이름) + '</div>' +
    '<div style="color:#C8D5E0;font-size:14px;margin-top:4px;">' + esc_(info.지원직무) + ' · ' + esc_(info.경력) + '</div>' +
  '</td></tr>');

  /* 메타 */
  var partT = meta.파트별소요초 || {};
  var partStr = [];
  for (var pk in partT) partStr.push(pk + ' ' + Math.round(partT[pk] / 60) + '분');
  h.push('<tr><td style="padding:14px 28px;background:#F7F8F7;border-bottom:1px solid ' + C_LINE + ';font-size:13px;color:' + C_MUTED + ';">' +
    '응시 ' + esc_(meta.응시일시 || '') + ' &nbsp;·&nbsp; 총 ' + esc_(meta.총소요 || '') +
    (partStr.length ? '<div style="margin-top:5px;font-size:12px;">' + esc_(partStr.join('  /  ')) + '</div>' : '') +
  '</td></tr>');

  /* 신뢰도 */
  var gc = gradeColor_(rel.종합등급);
  var flags = rel.감점요인 || [];
  h.push(section_('응답 신뢰도',
    '<div>' + chip_('신뢰도 ' + (rel.종합등급 || '-'), gc, '#FFFFFF') +
      (flags.length ? '<span style="font-size:13px;color:' + C_MUTED + ';margin-left:10px;">' + esc_(flags.join(' · ')) + '</span>'
                    : '<span style="font-size:13px;color:' + C_MUTED + ';margin-left:10px;">특이사항 없음</span>') +
    '</div>' +
    (flags.length ? '<div style="margin-top:8px;font-size:12px;color:' + C_MUTED + ';line-height:1.6;">' +
      '위 항목이 있으면 자기보고 점수는 다소 높게 나올 수 있습니다. 상황판단·서술형을 함께 보세요.</div>' : '')
  ));

  /* 종합 역량 */
  var rows = '';
  for (var k in combined) rows += bar_(k, combined[k], scoreColor_(combined[k]));
  h.push(section_('종합 역량', '<table width="100%" cellpadding="0" cellspacing="0" border="0">' + rows + '</table>',
    '자기보고 60% + 상황판단 40%'));

  /* 자기보고 vs 상황판단 */
  var sjtBy = sjt.역량별 || {};
  var cmp = '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;font-size:13px;">' +
    '<tr style="background:#F7F8F7;">' +
    '<th align="left" style="padding:8px 10px;border-bottom:1px solid ' + C_LINE + ';font-size:12px;color:' + C_MUTED + ';">역량</th>' +
    '<th align="center" style="padding:8px 10px;border-bottom:1px solid ' + C_LINE + ';font-size:12px;color:' + C_MUTED + ';">자기보고</th>' +
    '<th align="center" style="padding:8px 10px;border-bottom:1px solid ' + C_LINE + ';font-size:12px;color:' + C_MUTED + ';">상황판단</th>' +
    '<th align="center" style="padding:8px 10px;border-bottom:1px solid ' + C_LINE + ';font-size:12px;color:' + C_MUTED + ';">차이</th></tr>';
  var alias = { '성실·안전': '성실안전', '정서 안정': '정서안정' };
  for (var ck in combined) {
    var sv = selfS[ck] != null ? selfS[ck] : selfS[alias[ck]];
    var jv = sjtBy[ck];
    var diff = (sv != null && jv != null) ? (jv - sv) : null;
    var dTxt = diff == null ? '-' : (diff > 0 ? '+' + diff : String(diff));
    var dCol = diff == null ? C_MUTED : (Math.abs(diff) >= 20 ? C_WARN : C_MUTED);
    cmp += '<tr>' +
      '<td style="padding:9px 10px;border-bottom:1px solid #F0F2F1;">' + esc_(ck) + '</td>' +
      '<td align="center" style="padding:9px 10px;border-bottom:1px solid #F0F2F1;color:' + (sv != null ? scoreColor_(sv) : C_MUTED) + ';font-weight:700;">' + (sv != null ? sv : '-') + '</td>' +
      '<td align="center" style="padding:9px 10px;border-bottom:1px solid #F0F2F1;color:' + (jv != null ? scoreColor_(jv) : C_MUTED) + ';font-weight:700;">' + (jv != null ? jv : '-') + '</td>' +
      '<td align="center" style="padding:9px 10px;border-bottom:1px solid #F0F2F1;color:' + dCol + ';">' + dTxt + '</td></tr>';
  }
  cmp += '</table>' +
    '<div style="margin-top:8px;font-size:12px;color:' + C_MUTED + ';line-height:1.6;">' +
    '자기보고는 본인이 답한 성향, 상황판단은 실제 현장 상황에서의 선택입니다. 차이가 ±20 이상이면 그 역량은 면접에서 확인해 보세요.</div>';
  h.push(section_('자기보고 vs 상황판단', cmp, '상황판단 평균 ' + (sjt.평균 != null ? sjt.평균 : '-') + '점'));

  /* 실무감각 */
  var mechRows = x.실무감각응답 || [];
  var wrong = [];
  for (var mi = 0; mi < mechRows.length; mi++) if (mechRows[mi].정오 !== '정답') wrong.push(mechRows[mi]);
  var mechTxt = String(sum.실무감각 || '');
  var mechNum = parseInt(mechTxt.split('/')[0], 10);
  var mechTot = parseInt(mechTxt.split('/')[1], 10) || 6;
  var mechPct = mechTot ? Math.round(mechNum / mechTot * 100) : 0;
  var mechHtml = '<div style="font-size:22px;font-weight:800;color:' + scoreColor_(mechPct) + ';">' + esc_(mechTxt) +
    '<span style="font-size:13px;font-weight:400;color:' + C_MUTED + ';"> 정답</span></div>';
  if (wrong.length) {
    mechHtml += '<div style="margin-top:10px;font-size:13px;color:' + C_MUTED + ';">틀린 문항</div>';
    for (var wi = 0; wi < wrong.length; wi++) {
      mechHtml += '<div style="margin-top:6px;padding:10px 12px;background:#FBF7F2;border-left:3px solid ' + C_WARN + ';border-radius:0 6px 6px 0;font-size:13px;line-height:1.6;">' +
        esc_(wrong[wi].문항) +
        '<div style="margin-top:4px;color:' + C_MUTED + ';">지원자 답: ' + esc_(wrong[wi].지원자답 || '무응답') + ' &nbsp;/&nbsp; 정답: ' + esc_(wrong[wi].정답) + '</div></div>';
    }
  }
  h.push(section_('실무 감각 (기계 이해)', mechHtml));

  /* 동기 */
  var limb = motive.타고난동기_림빅양자택일 || {};
  var pct = limb.강도100 || {};
  var wish = motive.직장에바라는조건_우선순위기반 || motive.우선순위기반 || {};
  var mHtml = '';
  if (pct.지배욕 != null) {
    mHtml += '<div style="font-size:13px;color:' + C_MUTED + ';margin-bottom:2px;">타고난 동기 (양자택일 12문항)</div>' +
      '<table width="100%" cellpadding="0" cellspacing="0" border="0">' +
      bar_('지배욕 (성취·인정)', pct.지배욕, C_STEEL) +
      bar_('자극욕 (변화·도전)', pct.자극욕, '#7B6CA8') +
      bar_('균형욕 (안정·조화)', pct.균형욕, '#3E8E8A') +
      '</table>';
  }
  if (wish.지배욕 != null) {
    mHtml += '<div style="font-size:13px;color:' + C_MUTED + ';margin:14px 0 2px;">직장에 바라는 조건 (우선순위 기반)</div>' +
      '<table width="100%" cellpadding="0" cellspacing="0" border="0">' +
      bar_('지배욕', wish.지배욕, C_STEEL) +
      bar_('자극욕', wish.자극욕, '#7B6CA8') +
      bar_('균형욕', wish.균형욕, '#3E8E8A') +
      '</table>';
  }
  var rank = (x.동기응답 || {}).우선순위_직장에바라는조건 || [];
  if (rank.length) {
    mHtml += '<div style="margin-top:12px;font-size:13px;line-height:1.9;">';
    for (var ri = 0; ri < rank.length; ri++) mHtml += '<div>' + esc_(rank[ri]) + '</div>';
    mHtml += '</div>';
  }
  var gapTxt = motive.타고난동기와_직장조건의_괴리 || '';
  if (gapTxt.indexOf('있음') === 0) {
    mHtml += '<div style="margin-top:12px;padding:12px 14px;background:#FBF3EE;border-left:3px solid ' + C_BAD + ';border-radius:0 6px 6px 0;font-size:13px;line-height:1.7;">' +
      '<b style="color:' + C_BAD + ';">타고난 동기와 바라는 조건의 괴리 있음</b><br>' +
      '본능적으로 원하는 것과 직장에 바라는 것이 다릅니다. 지금 조건에 맞춰 참고 있을 가능성이 있어, 면접에서 실제로 원하는 근무 형태를 확인해 보세요.</div>';
  }
  if (motive.보상선택) {
    mHtml += '<div style="margin-top:12px;font-size:13px;color:' + C_MUTED + ';">보상 선택: <b style="color:' + C_INK + ';">' + esc_(motive.보상선택) + '</b></div>';
  }
  if (mHtml) h.push(section_('일 동기', mHtml));

  /* 상황판단 상세 */
  var sjtRows2 = x.상황판단응답 || [];
  if (sjtRows2.length) {
    var st = '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;font-size:13px;">';
    for (var si = 0; si < sjtRows2.length; si++) {
      var r2 = sjtRows2[si];
      var p2 = r2.점수100;
      st += '<tr><td style="padding:10px 0;border-bottom:1px solid #F0F2F1;">' +
        '<div style="color:' + C_MUTED + ';font-size:12px;">' + esc_(r2.번호 + '. ' + (r2.평가역량 || '')) + '</div>' +
        '<div style="margin-top:3px;line-height:1.6;">' + esc_(String(r2.상황 || '').slice(0, 70)) + (String(r2.상황 || '').length > 70 ? '…' : '') + '</div>' +
        '<div style="margin-top:5px;color:' + C_MUTED + ';font-size:12px;">할 것 <b style="color:' + C_INK + ';">' + esc_(r2.할것같은행동 || '-') + '</b>' +
        ' &nbsp; 안 할 것 <b style="color:' + C_INK + ';">' + esc_(r2.하지않을행동 || '-') + '</b>' +
        ' &nbsp; <span style="color:' + (p2 != null ? scoreColor_(p2) : C_MUTED) + ';font-weight:700;">' + (p2 != null ? p2 + '점' : '-') + '</span></div>' +
      '</td></tr>';
    }
    st += '</table>';
    h.push(section_('상황 판단 상세', st));
  }

  /* 서술형 */
  var essays = x.서술형응답 || [];
  if (essays.length) {
    var eh = '';
    for (var ei = 0; ei < essays.length; ei++) {
      var e2 = essays[ei];
      var empty = !e2.답변 || e2.답변 === '(무응답)';
      eh += '<div style="margin-bottom:12px;padding:14px 16px;background:#F7F8F7;border-radius:10px;">' +
        '<div style="font-size:12px;color:' + C_MUTED + ';">' + esc_((e2.관련역량 || '')) + ' · ' + esc_(e2.글자수 + '자') + '</div>' +
        '<div style="font-size:13px;font-weight:700;margin-top:4px;line-height:1.6;">' + esc_(e2.질문) + '</div>' +
        '<div style="font-size:14px;margin-top:8px;line-height:1.75;color:' + (empty ? C_MUTED : C_INK) + ';">' + nl2br_(e2.답변) + '</div>' +
      '</div>';
    }
    h.push(section_('서술형 답변', eh, '지원자가 직접 쓴 문장 — 태도와 표현력을 함께 보세요'));
  }

  /* 푸터 */
  h.push('<tr><td style="padding:26px 28px 28px 28px;">' +
    '<div style="border-top:1px solid ' + C_LINE + ';padding-top:18px;font-size:13px;color:' + C_MUTED + ';line-height:1.8;">' +
    '<b style="color:' + C_INK + ';">AI 상세 분석을 원하시면</b><br>' +
    '첨부된 텍스트 파일을 열어 전체를 복사한 뒤 Claude(claude.ai) 대화창에 붙여넣으세요.' +
    (sheetUrl ? '<br><br><b style="color:' + C_INK + ';">전체 접수 현황</b><br><a href="' + sheetUrl + '" style="color:' + C_STEEL + ';">스프레드시트 열기</a>' : '') +
    '</div></td></tr>');

  h.push('</table>');
  h.push('<div style="font-size:11px;color:#98A4AE;padding:14px 0 4px;">고강이엔지 인적성검사 · 채용 평가 목적 자료</div>');
  h.push('</td></tr></table></div>');
  return h.join('');
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
