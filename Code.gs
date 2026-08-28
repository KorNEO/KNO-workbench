/**
 * KNO Workbench v1.0 — 신어 판별 워크벤치 백엔드 (Google Apps Script)
 * 실행: 소유자(USER_DEPLOYING). 신원: 개인 링크 토큰(?u=...). 데이터: 프로젝트 단위 시트(항목_<id>).
 */

// ── 상수 ───────────────────────────────────────────────
var SHEET_USERS = '연구원';
var SHEET_GUIDE = '지침';
var SHEET_LOG = '변경로그';
var SHEET_FAILLOG = '저장실패로그';
var SHEET_PROJECTS = '프로젝트';
var PEPPER = 'KNO_v1_pepper';
var TZ = 'Asia/Seoul';
var TS_FMT = 'yyyy-MM-dd HH:mm:ss';

var HEADERS = [
  'ID', '신어 후보', '작업자', '검수자', '배정 주차',
  '1차 판별', '1차 일시', '1차 메모', '1차 고유명 여부', '1차 혐오 표현 여부',
  '2차 판별', '2차 일시', '2차 메모', '2차 고유명 여부', '2차 혐오 표현 여부',
  '상태', '작업 구분', '출처', '추출 시기',
  'LLM 판단 결과', 'LLM 판단 기준', 'LLM 판단 근거',
  '용례', '용례 일자', '용례 URL', '검색 URL'
];
var VERDICTS = ['신어', '비신어', '판단 보류'];
var CHECK_COLS = ['1차 고유명 여부', '1차 혐오 표현 여부', '2차 고유명 여부', '2차 혐오 표현 여부'];
function chkOk_(v) { return v === 'Y' || v === 'N'; }
var STATUS = { NONE: '미작업', FIRST: '1차완료', SECOND: '2차완료' };
var LOG_HEADERS = ['일시', 'ID', '신어 후보', '단계', '행위자', '판별', '메모', '이전 상태', '새 상태'];
var FAILLOG_HEADERS = ['일시', '행위자', '기능', 'ID', '판별', '메모', '에러', '토큰'];
var SAVE_FNS = { saveFirst: '1차', saveSecond: '2차', saveWrite: '집필' };

// 집필(M5) 항목 스키마 — 판별과 별개. 입력 필드는 1차/2차 쌍으로 저장('1차 조어법'/'2차 조어법' …), 부(형태부/의미부)별 상태·일시.
var WRITE_KINDS = ['집필', '집필 테스트'];
function isWriteKind_(k) { return WRITE_KINDS.indexOf(String(k || '').trim()) >= 0; }
var WRITE_PARTS = ['형태부', '의미부'];
var WRITE_PART_FIELDS = {
  '형태부': ['색인표제어', '단어/구', '조어법', '품사', '일상어/전문어', '전문 분야', '등재표제어', '고유어', '원어', '어종 표시', '어원', '의미 범주', '집필 메모(형태부)', '검수 메모(형태부)'],
  '의미부': ['뜻풀이', '용례', '참고 용례', '참고 용례 URL', 'X년 Y월 신어', '집필 메모(의미부)', '검수 메모(의미부)']
};
var WRITE_LLM = ['LLM 색인표제어', 'LLM 단어/구', 'LLM 조어법', 'LLM 품사', 'LLM 일상어/전문어', 'LLM 전문 분야', 'LLM 의미 범주', 'LLM 뜻풀이', 'LLM 용례'];   // 컬럼명 잠정(M5 출력 확정 시 맞춤)
var WRITE_REF = ['출처', '추출 시기', '판별 작업자', '판별 검수자', '1차 판별', '1차 메모', '1차 고유명 여부', '1차 혐오 표현 여부', '2차 판별', '2차 메모', '2차 고유명 여부', '2차 혐오 표현 여부', '용례', '용례 일자', '용례 URL', '검색 URL'];   // 판별 참조
var WRITE_META = ['ID', '신어 후보', '작업자', '검수자', '형태부 주차', '의미부 주차', '형태부 상태', '의미부 상태', '작업 구분', '형태부 1차 일시', '형태부 2차 일시', '의미부 1차 일시', '의미부 2차 일시'];
// 단계별 저장 컬럼: 집필 메모는 1차만, 검수 메모는 2차만
function writeStageCols_(part, stage) {
  return WRITE_PART_FIELDS[part].filter(function (f) { return !(stage === 1 && f.indexOf('검수 메모') === 0) && !(stage === 2 && f.indexOf('집필 메모') === 0); })
    .map(function (f) { return stage + '차 ' + f; });
}
var HEADERS_WRITE = (function () {
  var h = WRITE_META.concat(WRITE_REF, WRITE_LLM);
  [1, 2].forEach(function (st) { WRITE_PARTS.forEach(function (pt) { h = h.concat(writeStageCols_(pt, st)); }); });
  return h;
})();
// 집필 테스트: 이 명단은 작업자·검수자 양쪽 편집 가능
var WRITE_TEST_BOTH = ['관리자', '안진산', '이수진', '백미경'];
function canEditWrite_(me, kind, worker, reviewer, stage) {
  if (me.isManager) return true;
  if (kind === '집필 테스트' && WRITE_TEST_BOTH.indexOf(me.name) >= 0) return true;
  return stage === 1 ? (!!me.name && me.name === String(worker || '').trim()) : (!!me.name && me.name === String(reviewer || '').trim());
}
function headersFor_(kind) { return isWriteKind_(kind) ? HEADERS_WRITE : HEADERS; }
// 시트에 없는 헤더를 끝에 추가(기존 프로젝트 호환) 후 최신 인덱스 반환
function ensureCols_(sh, headers) {
  var idx = headerIndex_(sh), missing = headers.filter(function (h) { return !(h in idx); });
  if (missing.length) { sh.getRange(1, sh.getLastColumn() + 1, 1, missing.length).setValues([missing]); styleHeader_(sh, sh.getLastColumn()); idx = headerIndex_(sh); }
  return idx;
}

// 연구원 스키마: 구글계정1 이름2 역할3 초대일시4 응답상태5 수락일시6 token7 개인링크8 아이디9 비번해시10 소속11 성별12
var USER_HDR = ['구글 계정', '이름', '역할', '초대 일시', '응답 상태', '수락 일시', '토큰', '개인 링크', '아이디', '비밀번호 해시', '소속', '성별'];
var PROJ_HEADERS = ['프로젝트 ID', '이름', '유형', '등록일', '마감일', '상태', '파일 ID', '단계'];   // 등록일·마감일 = 관리자 설정, file_id = 프로젝트 전용 스프레드시트 ID, 단계 = 집필(형태부/의미부)
var DRIVE_ROOT = 'KNO 워크벤치';   // 최상위 드라이브 폴더(하위: 프로젝트/·원본 업로드/, 작업유형별)
var PRES_TTL = 100;
var WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbxGQ25QDvzAdXOCdYWXihv3Lkdj6zVXyq5M0KiGjccGJTRbiY1XRMvRjCHKrmlFdWLZ/exec';   // doPost API 엔드포인트(프론트가 fetch)
var PAGES_URL = 'https://korneo.github.io/KNO-workbench/';   // 프론트(GitHub Pages) — 개인/공통 링크는 이 주소

// ── 진입점 ─────────────────────────────────────────────
function doGet(e) {   // /exec 접근 시 Pages 프론트로 리다이렉트
  var token = (e && e.parameter && e.parameter.u) ? String(e.parameter.u).trim() : '';
  var url = PAGES_URL + (token ? '?u=' + encodeURIComponent(token) : '');
  return HtmlService.createHtmlOutput('<script>location.replace(' + JSON.stringify(url) + ');</script><p><a href="' + url + '">KNO Workbench로 이동</a></p>')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
// ── 외부 프론트(GitHub Pages)용 JSON API ──
// Pages 프론트가 fetch(단순요청)로 호출. 익명 배포 ContentService 응답은 CORS * 허용, text/plain 본문이라 프리플라이트 없음.
var API = {
  getBootstrap: getBootstrap, ping: ping, getPresence: getPresence,
  getGuide: getGuide, setGuide: setGuide,
  getProjects: getProjects, createProject: createProject, updateProject: updateProject, deleteProject: deleteProject, exportProject: exportProject, getTemplate: getTemplate, setProjectPhase: setProjectPhase,
  getResearchers: getResearchers, saveResearchers: saveResearchers,
  getAssignees: getAssignees, genAgree: genAgree, genReal: genReal,
  getProgress: getProgress, getItems: getItems, getItem: getItem,
  saveFirst: saveFirst, saveSecond: saveSecond, saveWrite: saveWrite, addWriteItem: addWriteItem, deleteWriteItem: deleteWriteItem, logClientFail: logClientFail,
  requestOtp: requestOtp, registerAccount: registerAccount, login: login, resetPassword: resetPassword, adminResetPassword: adminResetPassword
};
function doPost(e) {
  var out;
  try {
    var body = (e && e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : {};
    var fn = API[body.fn];
    if (typeof fn !== 'function') throw new Error('허용되지 않은 함수: ' + body.fn);
    out = { ok: true, data: fn.apply(null, body.args || []) };
  } catch (err) { out = { ok: false, error: String(err && err.message ? err.message : err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ── 공통 헬퍼 ──────────────────────────────────────────
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name) { var sh = ss_().getSheetByName(name); if (!sh) throw new Error('시트 없음: ' + name); return sh; }
function headerIndex_(sh) {
  var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0], map = {};
  for (var i = 0; i < hdr.length; i++) map[String(hdr[i]).trim()] = i;
  return map;
}
// 체크 컬럼(고유명·혐오)이 없는 기존 프로젝트 시트에 자동 추가 후 최신 헤더 인덱스 반환.
function ensureCheckCols_(sh) {
  var idx = headerIndex_(sh), missing = [];
  for (var i = 0; i < CHECK_COLS.length; i++) if (!(CHECK_COLS[i] in idx)) missing.push(CHECK_COLS[i]);
  if (missing.length) {
    sh.getRange(1, sh.getLastColumn() + 1, 1, missing.length).setValues([missing]);
    styleHeader_(sh, sh.getLastColumn());
    idx = headerIndex_(sh);
  }
  return idx;
}
// 첫행: 고정 + 하늘색 배경 + 데이터 필터
function styleHeader_(sh, ncols) {
  try {
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, ncols).setBackground('#cfe2f3').setFontWeight('bold');
    var ex = sh.getFilter(); if (ex) ex.remove();
    sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), ncols).createFilter();
  } catch (e) {}
}
function now_() { return Utilities.formatDate(new Date(), TZ, TS_FMT); }
function cacheGet_(key) { try { var v = CacheService.getScriptCache().get(key); return v ? JSON.parse(v) : null; } catch (e) { return null; } }
function cachePut_(key, obj, ttl) { try { CacheService.getScriptCache().put(key, JSON.stringify(obj), ttl || 60); } catch (e) {} }
function cacheDel_(keys) { try { CacheService.getScriptCache().removeAll([].concat(keys)); } catch (e) {} }
// 멱등키: 응답 유실로 같은 저장이 재전송돼도 1회만 적용. 적용 성공 후에만 마킹.
function opSeen_(opId) { if (!opId) return false; try { return !!CacheService.getScriptCache().get('op:' + opId); } catch (e) { return false; } }
function opMark_(opId) { if (opId) try { CacheService.getScriptCache().put('op:' + opId, '1', 3600); } catch (e) {} }
function getAppUrl_() { return PAGES_URL; }   // 개인/공통 링크는 Pages 프론트 주소
// 개인 링크: ?u=토큰 + &authuser=이메일 → 멀티계정 브라우저에서도 그 계정으로 열려 라우팅 오류 회피.
function personalLink_(url, token, email) { return (!url || !token) ? (url || '') : url + '?u=' + token + (email ? '&authuser=' + encodeURIComponent(email) : ''); }

// ── 인증 ───────────────────────────────────────────────
function hashPw_(pw) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pw) + '|' + PEPPER)); }
// OTP 검증: 일치하면 true(코드는 소비하지 않음 → 이후 검증 실패에도 재입력 가능). 불일치 시 실패 카운트, 5회 초과면 코드 폐기.
function otpVerify_(email, code) {
  var em = String(email || '').trim().toLowerCase(), c = CacheService.getScriptCache();
  var cached = c.get('otp:' + em);
  if (cached && cached === String(code || '').trim()) return true;
  var fkey = 'otpfail:' + em, fails = parseInt(c.get(fkey) || '0', 10) + 1;
  if (fails >= 5) { c.remove('otp:' + em); c.remove(fkey); throw new Error('인증번호를 5회 이상 틀려 무효화됐습니다. [인증번호 받기]로 다시 발급하세요.'); }
  c.put(fkey, String(fails), 600);
  throw new Error('인증번호가 올바르지 않거나 만료됐습니다.');
}
// OTP 소비: 작업 성공 후 코드·실패카운트 제거.
function otpConsume_(email) { var em = String(email || '').trim().toLowerCase(), c = CacheService.getScriptCache(); c.remove('otp:' + em); c.remove('otpfail:' + em); }
// 임시 비밀번호 생성(혼동되기 쉬운 0/O/1/l/I 제외).
function genTempPw_() { var cs = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789', s = ''; for (var i = 0; i < 8; i++) s += cs.charAt(Math.floor(Math.random() * cs.length)); return s; }
function usersSheet_() { return sheet_(SHEET_USERS); }
function ensureUserHeaders_(sh) { sh.getRange(1, 9, 1, 2).setValues([['아이디', '비밀번호 해시']]); }
function whoByToken_(token) {
  token = String(token || '').trim();
  var out = { token: token, known: false, email: '', name: '', role: '', isManager: false, _row: -1 };
  if (!token) return out;
  var hit = cacheGet_('who:' + token); if (hit) return hit;
  var sh = ss_().getSheetByName(SHEET_USERS);
  if (!sh || sh.getLastRow() < 2) return out;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][6]).trim() === token) {
      out.email = String(rows[i][0]).trim(); out.name = String(rows[i][1]).trim(); out.role = String(rows[i][2]).trim();
      out.isManager = /관리자/.test(out.role); out.known = true; out._row = i + 2; break;
    }
  }
  if (out.known) cachePut_('who:' + token, out, 300);
  return out;
}
function me_(token) { var me = whoByToken_(token); if (!me.known) throw new Error('세션이 만료됐거나 잘못된 링크입니다. 개인 링크로 다시 접속하세요.'); return me; }
function assertManager_(me) { if (!me.isManager) throw new Error('관리자 전용 기능입니다.'); }
function assertCanEditStage_(me, worker, reviewer, stage) {
  if (me.isManager) return;
  if (stage === 1 && me.name && me.name === String(worker).trim()) return;
  if (stage === 2 && me.name && me.name === String(reviewer).trim()) return;
  throw new Error('권한 없음: 배정된 담당자만 입력할 수 있습니다.');
}
function requestOtp(email) {
  email = String(email || '').trim();
  if (!email) throw new Error('이메일을 입력하세요.');
  var sh = usersSheet_(), n = sh.getLastRow() - 1;
  var rows = n > 0 ? sh.getRange(2, 1, n, 1).getValues() : [], found = false;
  for (var i = 0; i < rows.length; i++) if (String(rows[i][0]).trim().toLowerCase() === email.toLowerCase()) { found = true; break; }
  if (!found) throw new Error('등록된 연구원 이메일이 아닙니다. 관리자에게 문의하세요.');
  var ck = CacheService.getScriptCache(), ckey = 'otpsent:' + email.toLowerCase();
  if (ck.get(ckey)) throw new Error('인증번호를 방금 보냈습니다. 잠시 후 다시 시도하세요.');
  ck.put(ckey, '1', 60);
  var code = String(Math.floor(Math.random() * 900000) + 100000);
  ck.put('otp:' + email.toLowerCase(), code, 600);
  ck.remove('otpfail:' + email.toLowerCase());
  GmailApp.sendEmail(email, '[KNO Workbench] 인증번호',
    '신어 판별 및 집필 워크벤치 계정 등록을 위한 인증번호입니다.\n\n' + code + '\n\n10분 이내에 입력해 주세요. 😊',
    { name: 'KNO Workbench', htmlBody: '<p>신어 판별 및 집필 워크벤치 계정 등록을 위한 인증번호입니다.</p><p style="font-size:24px;font-weight:bold">' + code + '</p><p>10분 이내에 입력해 주세요. 😊</p>' });
  return { ok: true };
}
function registerAccount(email, code, id, pw, name) {
  email = String(email || '').trim(); id = String(id || '').trim();
  if (id.length < 2) throw new Error('아이디는 2자 이상이어야 합니다.');
  if (String(pw || '').length < 4) throw new Error('비밀번호는 4자 이상이어야 합니다.');
  otpVerify_(email, code);
  var lock = LockService.getDocumentLock(); lock.waitLock(15000);
  try {
    var sh = usersSheet_(); ensureUserHeaders_(sh);
    var n = sh.getLastRow() - 1, data = sh.getRange(2, 1, n, 10).getValues(), myRow = -1;
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][8]).trim().toLowerCase() === id.toLowerCase() && String(data[i][0]).trim().toLowerCase() !== email.toLowerCase())
        throw new Error('이미 사용 중인 아이디입니다.');
      if (String(data[i][0]).trim().toLowerCase() === email.toLowerCase()) myRow = i;
    }
    if (myRow < 0) throw new Error('등록된 연구원 이메일이 아닙니다.');
    if (name && String(data[myRow][1]).trim() !== String(name).trim()) throw new Error('이름이 등록 정보와 일치하지 않습니다.');
    if (String(data[myRow][8]).trim()) throw new Error('이미 등록된 계정입니다. 로그인하거나 [비밀번호 재설정]을 이용하세요.');
    var tok = String(data[myRow][6]).trim();
    if (!tok) { tok = Utilities.getUuid().replace(/-/g, '').slice(0, 10); sh.getRange(myRow + 2, 7).setValue(tok); }
    sh.getRange(myRow + 2, 9).setValue(id);
    sh.getRange(myRow + 2, 10).setValue(hashPw_(pw));
    otpConsume_(email);
    cacheDel_('who:' + tok);
    SpreadsheetApp.flush();
    return { ok: true, token: tok };
  } finally { lock.releaseLock(); }
}
function login(id, pw) {
  id = String(id || '').trim();
  if (!id) throw new Error('아이디를 입력하세요.');
  var sh = usersSheet_(), n = sh.getLastRow() - 1;
  if (n <= 0) throw new Error('등록된 계정이 없습니다.');
  var data = sh.getRange(2, 1, n, 10).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][8]).trim().toLowerCase() === id.toLowerCase()) {
      if (String(data[i][9]).trim() && String(data[i][9]).trim() === hashPw_(pw)) return { ok: true, token: String(data[i][6]).trim() };
      break;
    }
  }
  throw new Error('아이디 또는 비밀번호가 올바르지 않습니다.');
}
// 비밀번호 재설정: 이메일 OTP로 본인 확인 후 비번 해시 컬럼(10)만 갱신 — 아이디·토큰·작업 데이터는 그대로 보존.
function resetPassword(email, code, pw) {
  email = String(email || '').trim();
  if (!email) throw new Error('이메일을 입력하세요.');
  if (String(pw || '').length < 4) throw new Error('비밀번호는 4자 이상이어야 합니다.');
  otpVerify_(email, code);
  var lock = LockService.getDocumentLock(); lock.waitLock(15000);
  try {
    var sh = usersSheet_(); ensureUserHeaders_(sh);
    var n = sh.getLastRow() - 1;
    if (n <= 0) throw new Error('등록된 계정이 없습니다.');
    var data = sh.getRange(2, 1, n, 10).getValues(), myRow = -1;
    for (var i = 0; i < data.length; i++)
      if (String(data[i][0]).trim().toLowerCase() === email.toLowerCase()) { myRow = i; break; }
    if (myRow < 0) throw new Error('등록된 연구원 이메일이 아닙니다.');
    if (!String(data[myRow][8]).trim()) throw new Error('아직 등록되지 않은 계정입니다. [최초 등록]을 진행하세요.');
    sh.getRange(myRow + 2, 10).setValue(hashPw_(pw));
    var tok = String(data[myRow][6]).trim();
    if (!tok) { tok = Utilities.getUuid().replace(/-/g, '').slice(0, 10); sh.getRange(myRow + 2, 7).setValue(tok); }
    otpConsume_(email);
    SpreadsheetApp.flush();
    return { ok: true, token: tok };
  } finally { lock.releaseLock(); }
}
// 관리자 수동 초기화: 임시 비밀번호 발급 후 비번 해시(10)만 갱신 — 아이디·토큰·작업 데이터 보존. 임시비번은 관리자에게만 반환.
function adminResetPassword(token, email) {
  assertManager_(me_(token));
  email = String(email || '').trim();
  if (!email) throw new Error('이메일이 필요합니다.');
  var lock = LockService.getDocumentLock(); lock.waitLock(15000);
  try {
    var sh = usersSheet_(); ensureUserHeaders_(sh);
    var n = sh.getLastRow() - 1;
    if (n <= 0) throw new Error('등록된 계정이 없습니다.');
    var data = sh.getRange(2, 1, n, 10).getValues(), myRow = -1;
    for (var i = 0; i < data.length; i++)
      if (String(data[i][0]).trim().toLowerCase() === email.toLowerCase()) { myRow = i; break; }
    if (myRow < 0) throw new Error('해당 이메일의 연구원이 없습니다.');
    var id = String(data[myRow][8]).trim();
    if (!id) throw new Error('아직 계정을 등록하지 않은 연구원입니다. (초기화할 비밀번호가 없습니다.)');
    var tempPw = genTempPw_();
    sh.getRange(myRow + 2, 10).setValue(hashPw_(tempPw));
    SpreadsheetApp.flush();
    return { ok: true, id: id, tempPw: tempPw };
  } finally { lock.releaseLock(); }
}

// ── 부트스트랩 / 접속 현황 ─────────────────────────────
function getBootstrap(token) {
  var me = me_(token);
  return { me: me, verdicts: VERDICTS, appUrl: getAppUrl_() };
}
function ping(token) {
  var me = whoByToken_(token);
  if (me.known && me.name) { try { CacheService.getScriptCache().put('pres:' + me.name, String(Date.now()), PRES_TTL); } catch (e) {} }
  return { ok: true };
}
function getPresence(token) {
  var me = me_(token);
  try { CacheService.getScriptCache().put('pres:' + me.name, String(Date.now()), PRES_TTL); } catch (e) {}
  var roster = cacheGet_('presence:roster');
  if (!roster) {
    roster = [];
    var sh = ss_().getSheetByName(SHEET_USERS);
    if (sh && sh.getLastRow() >= 2) {
      var rows = sh.getRange(2, 2, sh.getLastRow() - 1, 2).getValues();
      for (var i = 0; i < rows.length; i++) { var nm = String(rows[i][0]).trim(); if (nm) roster.push({ name: nm, '역할': String(rows[i][1] || '').trim() }); }
    }
    cachePut_('presence:roster', roster, 300);
  }
  var keys = roster.map(function (r) { return 'pres:' + r.name; }), got = {};
  try { got = CacheService.getScriptCache().getAll(keys) || {}; } catch (e) { got = {}; }
  return roster.map(function (r) { return { name: r.name, '역할': r['역할'], online: !!got['pres:' + r.name] }; });
}

// ── 프로젝트 ───────────────────────────────────────────
function projRegSheet_() {
  var sh = ss_().getSheetByName(SHEET_PROJECTS);
  if (!sh) { sh = ss_().insertSheet(SHEET_PROJECTS); sh.getRange(1, 1, 1, PROJ_HEADERS.length).setValues([PROJ_HEADERS]); styleHeader_(sh, PROJ_HEADERS.length); return sh; }
  var hdr = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), PROJ_HEADERS.length)).getValues()[0];
  var h = hdr.map(function (x) { return String(x).trim(); });
  if (h.indexOf('마감일') < 0) {   // 구 스키마 → 신 스키마: 업로드 일자→등록일, 마감일 열 추가 (데이터 무손실 이관)
    var upIdx = h.indexOf('업로드 일자'); if (upIdx < 0) upIdx = h.indexOf('등록일');
    if (upIdx >= 0) {
      sh.getRange(1, upIdx + 1).setValue('등록일');                                    // 업로드 일자 → 등록일 (같은 열, 기존 값 유지)
      sh.insertColumnAfter(upIdx + 1); sh.getRange(1, upIdx + 2).setValue('마감일');   // 등록일 뒤에 마감일 삽입 (상태·파일 ID 자동 우측 이동)
      styleHeader_(sh, sh.getLastColumn());
    } else if (sh.getLastRow() < 2) {
      sh.getRange(1, 1, 1, PROJ_HEADERS.length).setValues([PROJ_HEADERS]); styleHeader_(sh, PROJ_HEADERS.length);   // 빈 시트: 헤더만 재설정
    }
  }
  if (h.indexOf('단계') < 0) { sh.getRange(1, sh.getLastColumn() + 1).setValue('단계'); styleHeader_(sh, sh.getLastColumn()); }   // 집필 단계 열 자동 추가
  return sh;
}
function fmtDate_(v) {   // Date/문자열 → 'YYYY.MM.DD'
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy.MM.dd');
  var s = String(v || '').trim(), m = s.match(/(\d{4})[-.](\d{1,2})[-.](\d{1,2})/);
  return m ? m[1] + '.' + ('0' + m[2]).slice(-2) + '.' + ('0' + m[3]).slice(-2) : s;
}
function kindLabel_(kind) { return kind === '일치도' ? '연구자 일치도 작업' : kind === '집필' ? '신어 집필 작업' : kind === '집필 테스트' ? '신어 집필 작업(테스트)' : '신어 판별 작업'; }
function folder_(pathArr) {   // 중첩 폴더 get-or-create
  var f = DriveApp.getRootFolder();
  for (var i = 0; i < pathArr.length; i++) { var it = f.getFoldersByName(pathArr[i]); f = it.hasNext() ? it.next() : f.createFolder(pathArr[i]); }
  return f;
}
function projList_(kind) {
  var sh = projRegSheet_(), n = sh.getLastRow() - 1; if (n < 1) return [];
  var idx = headerIndex_(sh), rows = sh.getRange(2, 1, n, sh.getLastColumn()).getValues(), out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i]; if (kind && String(r[idx['유형']]).trim() !== kind) continue;
    out.push({ id: String(r[idx['프로젝트 ID']]).trim(), name: String(r[idx['이름']]).trim(), type: String(r[idx['유형']]).trim(),
      registered: fmtDate_(r[idx['등록일']]), due: fmtDate_(r[idx['마감일']]), status: String(r[idx['상태']]).trim(), fileId: String(r[idx['파일 ID']] || '').trim(),
      phase: (idx['단계'] != null && String(r[idx['단계']] || '').trim()) || '형태부', _row: i + 2 });
  }
  return out;
}
function projById_(id) { var a = projList_(null); for (var i = 0; i < a.length; i++) if (a[i].id === id) return a[i]; return null; }
function projItemSheet_(projectId) {   // 프로젝트 전용 스프레드시트의 '항목' 탭
  var p = projById_(projectId); if (!p || !p.fileId) return null;
  try { var pss = SpreadsheetApp.openById(p.fileId); return pss.getSheetByName('항목') || pss.getSheets()[0]; } catch (e) { return null; }
}
function projItemsCount_(p) { var sh = projItemSheet_(p.id); return sh ? Math.max(0, sh.getLastRow() - 1) : 0; }
function projSheetOfRow_(rowId) { return projItemSheet_(String(rowId).split('::')[0]); }
function projSetStatus_(projectId, status) { var p = projById_(projectId); if (!p) return; var reg = projRegSheet_(), idx = headerIndex_(reg); reg.getRange(p._row, idx['상태'] + 1).setValue(status); }

function getProjects(token, kind) {
  me_(token);
  return projList_(kind).map(function (p) { return { id: p.id, name: p.name, type: p.type, registered: p.registered, due: p.due, status: p.status, phase: p.phase, items: projItemsCount_(p) }; });
}
// 집필 프로젝트 단계 전환(관리자): 형태부 ↔ 의미부(되돌리기 허용)
function setProjectPhase(token, projectId, phase) {
  assertManager_(me_(token));
  if (WRITE_PARTS.indexOf(phase) < 0) throw new Error('단계 값 오류: ' + phase);
  var p = projById_(projectId); if (!p) throw new Error('프로젝트 없음');
  if (!isWriteKind_(p.type)) throw new Error('집필 프로젝트가 아닙니다.');
  var reg = projRegSheet_(), idx = headerIndex_(reg);
  reg.getRange(p._row, idx['단계'] + 1).setValue(phase);
  return { ok: true, phase: phase };
}
function createProject(token, kind, name, csvText, regDate, dueDate) {
  assertManager_(me_(token));
  name = String(name || '').trim() || '새 프로젝트';
  var H = headersFor_(kind);
  var pid = (kind === '일치도' ? 'ag' : kind === '집필' ? 'wr' : kind === '집필 테스트' ? 'wt' : 'rl') + '-' + Date.now().toString(36);
  try { folder_([DRIVE_ROOT, '원본 업로드', kindLabel_(kind)]).createFile(Utilities.newBlob(String(csvText || ''), 'text/csv', name + '.csv')); } catch (e) {}   // 원본 보관(재현성)
  var pss = SpreadsheetApp.create(name), fileId = pss.getId();   // 프로젝트 전용 스프레드시트
  try { var file = DriveApp.getFileById(fileId); folder_([DRIVE_ROOT, '프로젝트', kindLabel_(kind)]).addFile(file); DriveApp.getRootFolder().removeFile(file); } catch (e) {}
  var sh = pss.getSheets()[0]; sh.setName('항목');
  sh.getRange(1, 1, 1, H.length).setValues([H]);
  var t = Utilities.parseCsv(String(csvText || '').replace(/^﻿/, ''));
  var hdr = (t[0] || []).map(function (h) { return String(h).replace(/^﻿/, '').trim(); });
  var col = {}; hdr.forEach(function (h, i) { if (H.indexOf(h) >= 0) col[h] = i; });
  var CLEAR = ['작업자', '검수자', '배정 주차', '1차 판별', '1차 메모', '1차 일시', '1차 고유명 여부', '1차 혐오 표현 여부', '2차 판별', '2차 메모', '2차 일시', '2차 고유명 여부', '2차 혐오 표현 여부'];   // 판별: 배정·판별 추적 초기화
  if (isWriteKind_(kind)) {   // 집필: 판별 참조(1차/2차 판별·메모·고유명·혐오)와 CSV의 작업자·검수자·주차는 유지, 집필 입력·상태·일시만 초기화
    CLEAR = ['형태부 상태', '의미부 상태', '형태부 1차 일시', '형태부 2차 일시', '의미부 1차 일시', '의미부 2차 일시'];
    [1, 2].forEach(function (st) { WRITE_PARTS.forEach(function (pt) { CLEAR = CLEAR.concat(writeStageCols_(pt, st)); }); });
  }
  var out = [];
  for (var r = 1; r < t.length; r++) {
    var row = t[r]; if (!row || row.join('') === '') continue;
    var cand = col['신어 후보'] != null ? String(row[col['신어 후보']] || '').trim() : ''; if (!cand) continue;
    var o = [];
    for (var c = 0; c < H.length; c++) {
      var key = H[c], ci = col[key], v = (ci != null && row[ci] != null) ? row[ci] : '';
      if (key === 'ID') v = pid + '::' + (v || r);
      else if (key === '작업 구분') v = kind;
      else if (key === '상태' || key === '형태부 상태' || key === '의미부 상태') v = STATUS.NONE;
      else if (CLEAR.indexOf(key) >= 0) v = '';
      o.push(v);
    }
    out.push(o);
  }
  if (out.length) sh.getRange(2, 1, out.length, H.length).setValues(out);
  styleHeader_(sh, H.length);
  var reg = fmtDate_(regDate) || Utilities.formatDate(new Date(), TZ, 'yyyy.MM.dd');   // 미지정 시 오늘
  var regSh = projRegSheet_(), ridx = headerIndex_(regSh), rrow = [];
  PROJ_HEADERS.forEach(function (h) { rrow[ridx[h]] = ({ '프로젝트 ID': pid, '이름': name, '유형': kind, '등록일': reg, '마감일': fmtDate_(dueDate), '상태': '미배분', '파일 ID': fileId, '단계': isWriteKind_(kind) ? '형태부' : '' })[h]; });
  regSh.appendRow(rrow);
  return { id: pid, name: name, items: out.length };
}
function updateProject(token, kind, id, name, regDate, dueDate) {
  assertManager_(me_(token));
  var p = projById_(id); if (!p) throw new Error('프로젝트 없음');
  var reg = projRegSheet_(), idx = headerIndex_(reg);
  name = String(name || '').trim();
  if (name) {
    reg.getRange(p._row, idx['이름'] + 1).setValue(name);
    if (p.fileId) { try { DriveApp.getFileById(p.fileId).setName(name); } catch (e) {} }   // 드라이브 파일명 동기화
  }
  reg.getRange(p._row, idx['등록일'] + 1).setValue(fmtDate_(regDate) || p.registered);   // 등록일 미지정 시 기존 유지
  reg.getRange(p._row, idx['마감일'] + 1).setValue(fmtDate_(dueDate));                    // 마감일 빈 값 허용
  return { ok: true };
}
function deleteProject(token, kind, id) {
  assertManager_(me_(token));
  var p = projById_(id); if (!p) return { ok: true };
  if (p.fileId) { try { DriveApp.getFileById(p.fileId).setTrashed(true); } catch (e) {} }   // 파일 → 휴지통(영구삭제 아님)
  projRegSheet_().deleteRow(p._row);
  return { ok: true };
}
function exportProject(token, kind, id) {
  assertManager_(me_(token));
  var sh = projItemSheet_(id); if (!sh) throw new Error('프로젝트 없음');
  return toCsv_(sh.getDataRange().getValues());
}
function getTemplate(token, kind) {
  me_(token);
  if (isWriteKind_(kind)) return ['ID', '신어 후보', '작업자', '검수자', '형태부 주차', '의미부 주차'].concat(WRITE_REF, WRITE_LLM).join(',');
  return '작업 구분,작업자,검수자,배정 주차,상태,1차 판별,1차 메모,1차 일시,1차 고유명 여부,1차 혐오 표현 여부,2차 판별,2차 메모,2차 일시,2차 고유명 여부,2차 혐오 표현 여부,ID,신어 후보,출처,추출 시기,LLM 판단 결과,LLM 판단 기준,LLM 판단 근거,용례,용례 일자,용례 URL,검색 URL';
}

// ── 지침 ───────────────────────────────────────────────
function getGuide(token, which) {
  me_(token);
  var sh = ss_().getSheetByName(SHEET_GUIDE); if (!sh || sh.getLastRow() < 1) return '';
  var vals = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  for (var i = 0; i < vals.length; i++) if (String(vals[i][0]).trim() === which) return String(vals[i][1] || '');
  return '';
}
function setGuide(token, which, md) {
  assertManager_(me_(token));
  var sh = ss_().getSheetByName(SHEET_GUIDE) || ss_().insertSheet(SHEET_GUIDE);
  var vals = sh.getLastRow() ? sh.getRange(1, 1, sh.getLastRow(), 2).getValues() : [];
  for (var i = 0; i < vals.length; i++) if (String(vals[i][0]).trim() === which) { sh.getRange(i + 1, 2).setValue(String(md || '')); return { ok: true }; }
  sh.appendRow([which, String(md || '')]); return { ok: true };
}

// ── 진행률 ─────────────────────────────────────────────
function researcherOrder_() {
  var map = {}, sh = ss_().getSheetByName(SHEET_USERS);
  if (!sh || sh.getLastRow() < 2) return map;
  var names = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < names.length; i++) map[String(names[i][0]).trim()] = i;
  return map;
}
function getProgress(token, kind, projectId) {
  me_(token);
  if (isWriteKind_(kind)) return getProgressWrite_(projectId);
  var sh = projItemSheet_(projectId), idx = sh ? headerIndex_(sh) : {}, n = sh ? sh.getLastRow() - 1 : 0;
  var overall = { total: 0, 미작업: 0, '1차완료': 0, '2차완료': 0, done1: 0, done2: 0, weeks: {} }, groups = {};
  if (sh && n > 0) {
    var need = ['작업자', '검수자', '배정 주차', '상태', '1차 판별', '2차 판별'], cmin = Infinity, cmax = -1;
    for (var ni = 0; ni < need.length; ni++) { var ci = idx[need[ni]]; if (ci != null) { if (ci < cmin) cmin = ci; if (ci > cmax) cmax = ci; } }
    var data = sh.getRange(2, cmin + 1, n, cmax - cmin + 1).getValues();
    var dW = idx['작업자'] - cmin, dR = idx['검수자'] - cmin, dWk = idx['배정 주차'] - cmin, dS = idx['상태'] - cmin;
    var dV1 = idx['1차 판별'] != null ? idx['1차 판별'] - cmin : -1, dV2 = idx['2차 판별'] != null ? idx['2차 판별'] - cmin : -1;
    for (var r = 0; r < data.length; r++) {
      var st = String(data[r][dS]).trim() || '미작업';
      var f1 = dV1 >= 0 && !!String(data[r][dV1]).trim(), f2 = dV2 >= 0 && !!String(data[r][dV2]).trim();   // 1차·2차 각각 독립 완료 여부
      overall.total++; if (overall[st] !== undefined) overall[st]++;
      if (f1) overall.done1++; if (f2) overall.done2++;
      var wk = String(data[r][dWk]).trim(); if (wk) overall.weeks[wk] = (overall.weeks[wk] || 0) + 1;
      var w = String(data[r][dW]).trim(), rv = String(data[r][dR]).trim();
      var key = (kind === '일치도') ? (w || '(미배정)') : ((w || '?') + ' / ' + (rv || '?'));
      if (!groups[key]) groups[key] = { label: key, worker: w, reviewer: rv, total: 0, 미작업: 0, '1차완료': 0, '2차완료': 0, done1: 0, done2: 0, weeks: {} };
      groups[key].total++; if (groups[key][st] !== undefined) groups[key][st]++;
      if (f1) groups[key].done1++; if (f2) groups[key].done2++;
      if (wk) groups[key].weeks[wk] = (groups[key].weeks[wk] || 0) + 1;
    }
  }
  var order = researcherOrder_(), arr = Object.keys(groups).map(function (k) { return groups[k]; });
  arr.sort(function (a, b) { var ia = order[a.worker], ib = order[b.worker]; ia = (ia == null ? 9999 : ia); ib = (ib == null ? 9999 : ib); return ia !== ib ? ia - ib : a.label.localeCompare(b.label, 'ko'); });
  var weekList = Object.keys(overall.weeks).sort(function (a, b) { return (parseInt(a, 10) || 0) - (parseInt(b, 10) || 0); });
  return { overall: overall, groups: arr, weeks: weekList };
}

// 집필 진행률: 형태부/의미부 각각 1차·2차 완료 수. done1/done2는 현재 단계의 부 기준(기존 화면 호환).
function getProgressWrite_(projectId) {
  var p = projById_(projectId), phase = (p && p.phase) || '형태부';
  var sh = projItemSheet_(projectId), idx = sh ? headerIndex_(sh) : {}, n = sh ? sh.getLastRow() - 1 : 0;
  function mk() { return { total: 0, 미작업: 0, '1차완료': 0, '2차완료': 0, done1: 0, done2: 0, weeks: {}, parts: { '형태부': { done1: 0, done2: 0 }, '의미부': { done1: 0, done2: 0 } } }; }
  var overall = mk(), groups = {};
  if (sh && n > 0) {
    var data = sh.getRange(2, 1, n, sh.getLastColumn()).getValues();
    function g(r, k) { return idx[k] != null ? String(data[r][idx[k]] || '').trim() : ''; }
    for (var r = 0; r < data.length; r++) {
      var w = g(r, '작업자'), rv = g(r, '검수자'), key = (w || '?') + ' / ' + (rv || '?');
      if (!groups[key]) { groups[key] = mk(); groups[key].label = key; groups[key].worker = w; groups[key].reviewer = rv; }
      [overall, groups[key]].forEach(function (o) {
        o.total++;
        WRITE_PARTS.forEach(function (pt) { var st = g(r, pt + ' 상태'); if (st === STATUS.FIRST || st === STATUS.SECOND) o.parts[pt].done1++; if (st === STATUS.SECOND) o.parts[pt].done2++; });
        var cur = g(r, phase + ' 상태') || STATUS.NONE; if (o[cur] !== undefined) o[cur]++;
        o.done1 = o.parts[phase].done1; o.done2 = o.parts[phase].done2;
        var wk = g(r, phase + ' 주차'); if (wk) o.weeks[wk] = (o.weeks[wk] || 0) + 1;
      });
    }
  }
  var order = researcherOrder_(), arr = Object.keys(groups).map(function (k) { return groups[k]; });
  arr.sort(function (a, b) { var ia = order[a.worker], ib = order[b.worker]; ia = (ia == null ? 9999 : ia); ib = (ib == null ? 9999 : ib); return ia !== ib ? ia - ib : a.label.localeCompare(b.label, 'ko'); });
  var weekList = Object.keys(overall.weeks).sort(function (a, b) { return (parseInt(a, 10) || 0) - (parseInt(b, 10) || 0); });
  return { overall: overall, groups: arr, weeks: weekList, phase: phase };
}

// ── 항목 조회 ──────────────────────────────────────────
// 일치도 작업자→행범위 인덱스(genAgree가 작업자별 연속 정렬). 재배분 시 무효화.
function agreeIndex_(sh) {
  var props = PropertiesService.getScriptProperties(), key = 'idx:' + sh.getParent().getId(), raw = props.getProperty(key);
  if (raw) { try { return JSON.parse(raw); } catch (e) {} }
  var idx = headerIndex_(sh), n = sh.getLastRow() - 1; if (n <= 0) return {};
  var ws = sh.getRange(2, idx['작업자'] + 1, n, 1).getValues(), ks = sh.getRange(2, idx['작업 구분'] + 1, n, 1).getValues(), map = {};
  for (var i = 0; i < n; i++) {
    if (String(ks[i][0]).trim() !== '일치도') continue;
    var w = String(ws[i][0]).trim(); if (!w) continue; var row = i + 2;
    if (!map[w]) map[w] = [row, row]; else { if (row < map[w][0]) map[w][0] = row; if (row > map[w][1]) map[w][1] = row; }
  }
  try { props.setProperty(key, JSON.stringify(map)); } catch (e) {}
  return map;
}
var LIST_FIELDS = ['ID', '신어 후보', '출처', '추출 시기', '작업 구분', '작업자', '검수자', '배정 주차', '상태', '1차 판별', '2차 판별', '1차 고유명 여부', '1차 혐오 표현 여부', '2차 고유명 여부', '2차 혐오 표현 여부'];
var LIST_FIELDS_WRITE = ['ID', '신어 후보', '출처', '추출 시기', '작업 구분', '작업자', '검수자', '형태부 주차', '의미부 주차', '형태부 상태', '의미부 상태'];
function getItems(token, opts) {
  var me = me_(token); opts = opts || {};
  if (opts.kind === '일치도') opts.onlyMine = true;
  var sh = projItemSheet_(opts.projectId); if (!sh) return [];
  var n = sh.getLastRow() - 1; if (n <= 0) return [];
  var idx = headerIndex_(sh), lastCol = sh.getLastColumn();
  var startRow = 2, numRows = n;
  if (opts.kind === '일치도' && !me.isManager && me.name) {
    try { var rng = agreeIndex_(sh)[me.name];
      if (rng && rng[0] >= 2 && rng[1] >= rng[0]) { startRow = rng[0]; numRows = Math.min(rng[1], n + 1) - rng[0] + 1;
        if (numRows < 1 || startRow > n + 1) { startRow = 2; numRows = n; } }
    } catch (e) { startRow = 2; numRows = n; }
  }
  var isW = isWriteKind_(opts.kind), LF = isW ? LIST_FIELDS_WRITE : LIST_FIELDS, phase = '형태부';
  if (isW) { var pj = projById_(opts.projectId); phase = (pj && pj.phase) || '형태부'; }
  var stKey = isW ? phase + ' 상태' : '상태', wkKey = isW ? phase + ' 주차' : '배정 주차';
  var seeAll = me.isManager || (opts.kind === '집필 테스트' && WRITE_TEST_BOTH.indexOf(me.name) >= 0);
  var full = !!opts.full, blkCols = lastCol;
  if (!full) { var maxc = 0; for (var li = 0; li < LF.length; li++) { var ci = idx[LF[li]]; if (ci != null && ci > maxc) maxc = ci; } blkCols = Math.min(maxc + 1, lastCol); }
  var BLK = sh.getRange(startRow, 1, numRows, blkCols).getValues();
  function g(r, key) { var c = idx[key]; if (c == null || c >= blkCols) return ''; return String(BLK[r][c] || '').trim(); }
  var qy = opts.q ? String(opts.q).toLowerCase() : '', out = [];
  for (var r = 0; r < numRows; r++) {
    var k = g(r, '작업 구분'), w = g(r, '작업자'), rv = g(r, '검수자'), st = g(r, stKey), wk = g(r, wkKey);
    if (opts.kind && k !== opts.kind) continue;
    if (opts.worker && w !== opts.worker) continue;
    if (opts.week && wk !== String(opts.week)) continue;
    if (opts.status && (st || STATUS.NONE) !== opts.status) continue;
    if (opts.onlyMine && !seeAll && me.name !== w && me.name !== rv) continue;
    if (qy && g(r, '신어 후보').toLowerCase().indexOf(qy) === -1) continue;
    var obj = {}, fields = full ? Object.keys(idx) : LF;   // full=시트 실제 컬럼(판별·집필 공용)
    for (var h = 0; h < fields.length; h++) obj[fields[h]] = g(r, fields[h]);
    out.push(obj);
  }
  return out;
}
function getItem(token, rowId) {
  me_(token);
  var sh = projSheetOfRow_(rowId); if (!sh) throw new Error('프로젝트 없음: ' + rowId);
  var idx = headerIndex_(sh), rownum = findRow_(sh, idx, rowId);
  if (rownum < 0) throw new Error('행 없음: ' + rowId);
  var row = sh.getRange(rownum, 1, 1, sh.getLastColumn()).getValues()[0], obj = {};
  for (var key in idx) obj[key] = String(row[idx[key]]);   // 시트 실제 컬럼 전부(판별·집필 공용)
  return obj;
}
function getAssignees(kind, projectId) {   // 프론트가 (kind, pid)로 호출(토큰 없음)
  var sh = projItemSheet_(projectId); if (!sh) return [];
  var idx = headerIndex_(sh), n = sh.getLastRow() - 1; if (n <= 0) return [];
  var need = ['작업자', '검수자'], cmin = Infinity, cmax = -1;
  for (var ni = 0; ni < need.length; ni++) { var ci = idx[need[ni]]; if (ci != null) { if (ci < cmin) cmin = ci; if (ci > cmax) cmax = ci; } }
  var data = sh.getRange(2, cmin + 1, n, cmax - cmin + 1).getValues(), dW = idx['작업자'] - cmin, dR = idx['검수자'] - cmin, seen = {}, out = [];
  for (var r = 0; r < data.length; r++) {
    var w = String(data[r][dW]).trim(), rv = String(data[r][dR]).trim();
    if (!w || seen[w]) continue; seen[w] = true;
    out.push({ worker: w, reviewer: rv, label: kind === '일치도' ? w : (w + ' - ' + rv) });
  }
  return out;   // 정렬 안 함 — 시트 등장 순서(=genReal 팀 블록/배분 순서) 유지
}

// ── 배분(프로젝트 시트 재구성) ─────────────────────────
function genAgree(token, projectId, names) {
  assertManager_(me_(token));
  if (!names || !names.length) throw new Error('참여자를 선택하세요.');
  var p = projById_(projectId); if (!p) throw new Error('프로젝트 없음');
  var sh = projItemSheet_(projectId); if (!sh) throw new Error('시트 없음');
  var idx = ensureCheckCols_(sh), lastCol = sh.getLastColumn(), n = sh.getLastRow() - 1, C = {};
  HEADERS.forEach(function (h) { C[h] = idx[h]; });
  var base = {};
  if (n > 0) { var data = sh.getRange(2, 1, n, lastCol).getValues();
    for (var r = 0; r < data.length; r++) { var b = String(data[r][C['ID']]).trim().split('#')[0]; if (!base[b]) base[b] = data[r]; } }
  var out = [];
  Object.keys(base).forEach(function (b) { var src = base[b];
    names.forEach(function (nm) { var o = src.slice();
      o[C['ID']] = b + '#' + nm; o[C['작업 구분']] = '일치도'; o[C['작업자']] = nm; o[C['검수자']] = '';
      o[C['배정 주차']] = ''; o[C['상태']] = STATUS.NONE;
      o[C['1차 판별']] = ''; o[C['1차 메모']] = ''; o[C['1차 일시']] = ''; o[C['2차 판별']] = ''; o[C['2차 메모']] = ''; o[C['2차 일시']] = '';
      o[C['1차 고유명 여부']] = ''; o[C['1차 혐오 표현 여부']] = ''; o[C['2차 고유명 여부']] = ''; o[C['2차 혐오 표현 여부']] = '';
      out.push(o); }); });
  out.sort(function (a, b) { var aw = String(a[C['작업자']]), bw = String(b[C['작업자']]); return aw !== bw ? aw.localeCompare(bw, 'ko') : String(a[C['ID']]).localeCompare(String(b[C['ID']])); });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).clearContent();
  if (out.length) sh.getRange(2, 1, out.length, lastCol).setValues(out);
  styleHeader_(sh, HEADERS.length);
  try { PropertiesService.getScriptProperties().deleteProperty('idx:' + sh.getParent().getId()); } catch (e) {}
  projSetStatus_(projectId, '배분완료');
  return { ok: true, rows: out.length };
}
function genReal(token, projectId, cfg) {
  assertManager_(me_(token));
  cfg = cfg || {}; var pairs = cfg.pairs || [], weeks = Math.max(1, parseInt(cfg.weeks, 10) || 4);
  if (!pairs.length) throw new Error('팀을 지정하세요.');
  var p = projById_(projectId); if (!p) throw new Error('프로젝트 없음');
  var sh = projItemSheet_(projectId); if (!sh) throw new Error('시트 없음');
  var isW = isWriteKind_(p.type);
  var idx = isW ? ensureCols_(sh, HEADERS_WRITE) : ensureCheckCols_(sh), lastCol = sh.getLastColumn(), n = sh.getLastRow() - 1, C = {};
  headersFor_(p.type).forEach(function (h) { C[h] = idx[h]; });
  var base = {};
  if (n > 0) { var data = sh.getRange(2, 1, n, lastCol).getValues();
    for (var r = 0; r < data.length; r++) { var b = String(data[r][C['ID']]).trim().split('#')[0]; if (!base[b]) base[b] = data[r].slice(); } }
  var rows = Object.keys(base).map(function (k) { return base[k]; });
  rows.sort(function (a, b) { return String(a[C['신어 후보']]).localeCompare(String(b[C['신어 후보']]), 'ko'); });
  var P = pairs.length, N = rows.length, per = Math.floor(N / P), rem = N % P, pos = 0;
  for (var pi = 0; pi < P; pi++) {
    var cnt = per + (pi < rem ? 1 : 0), block = rows.slice(pos, pos + cnt); pos += cnt;
    var m = block.length, wb = Math.floor(m / weeks), wr = m % weeks, bp = 0;
    for (var wk = 0; wk < weeks; wk++) { var take = wb + (wk < wr ? 1 : 0);
      for (var bi = 0; bi < take; bi++) { var rr = block[bp + bi];
        rr[C['ID']] = String(rr[C['ID']]).split('#')[0];
        rr[C['작업 구분']] = p.type; rr[C['작업자']] = pairs[pi][0]; rr[C['검수자']] = pairs[pi][1];
        if (isW) { rr[C['형태부 주차']] = String(wk + 1); rr[C['의미부 주차']] = String(wk + 1 + weeks); rr[C['형태부 상태']] = STATUS.NONE; rr[C['의미부 상태']] = STATUS.NONE; }
        else { rr[C['배정 주차']] = String(wk + 1); rr[C['상태']] = STATUS.NONE; } }
      bp += take; }
  }
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, lastCol).setValues(rows);
  styleHeader_(sh, lastCol);
  projSetStatus_(projectId, '배분완료');
  return { ok: true, rows: rows.length };
}

// ── 저장 ───────────────────────────────────────────────
function findRow_(sh, idx, rowId) {
  var n = sh.getLastRow() - 1; if (n <= 0) return -1;
  var ids = sh.getRange(2, idx['ID'] + 1, n, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === String(rowId).trim()) return i + 2;
  return -1;
}
function setCell_(sh, rownum, idx, header, value) { sh.getRange(rownum, idx[header] + 1).setValue(value); }
function saveFirst(token, payload) {
  var me = me_(token);
  if (VERDICTS.indexOf(payload.verdict) === -1) throw new Error('판별 값 오류');
  if (!chkOk_(payload.proper) || !chkOk_(payload.hate)) throw new Error('고유명·혐오 표현 여부를 모두 선택해야 저장됩니다.');
  var lock = LockService.getDocumentLock(); lock.waitLock(20000);
  try {
    if (opSeen_(payload.op_id)) return { ok: true, dup: true };
    var sh = projSheetOfRow_(payload.row_id); if (!sh) throw new Error('프로젝트 없음');
    var idx = ensureCheckCols_(sh), rownum = findRow_(sh, idx, payload.row_id);
    if (rownum < 0) throw new Error('행 없음: ' + payload.row_id);
    assertCanEditStage_(me, sh.getRange(rownum, idx['작업자'] + 1).getValue(), sh.getRange(rownum, idx['검수자'] + 1).getValue(), 1);
    var cand = String(sh.getRange(rownum, idx['신어 후보'] + 1).getValue());
    var prev = String(sh.getRange(rownum, idx['상태'] + 1).getValue()).trim();
    var prevVerdict = String(sh.getRange(rownum, idx['1차 판별'] + 1).getValue()).trim();
    var prevTs = String(sh.getRange(rownum, idx['1차 일시'] + 1).getValue()).trim();
    setCell_(sh, rownum, idx, '1차 판별', payload.verdict);
    setCell_(sh, rownum, idx, '1차 메모', payload.memo || '');
    setCell_(sh, rownum, idx, '1차 고유명 여부', payload.proper);
    setCell_(sh, rownum, idx, '1차 혐오 표현 여부', payload.hate);
    if (!(prevVerdict === payload.verdict && prevTs)) setCell_(sh, rownum, idx, '1차 일시', now_());   // 같은 판별 재저장(체크 추가)이면 원래 일시 보존
    var ns = String(sh.getRange(rownum, idx['2차 판별'] + 1).getValue()).trim() ? STATUS.SECOND : STATUS.FIRST;
    setCell_(sh, rownum, idx, '상태', ns);
    appendLog_(me, payload.row_id, cand, '1차', payload.verdict + ' (고유명' + payload.proper + '/혐오' + payload.hate + ')', payload.memo || '', prev, ns);
    SpreadsheetApp.flush(); opMark_(payload.op_id);
    return { ok: true };
  } finally { lock.releaseLock(); }
}
function saveSecond(token, payload) {
  var me = me_(token);
  if (VERDICTS.indexOf(payload.verdict) === -1) throw new Error('판별 값 오류');
  if (!chkOk_(payload.proper) || !chkOk_(payload.hate)) throw new Error('고유명·혐오 표현 여부를 모두 선택해야 저장됩니다.');
  var lock = LockService.getDocumentLock(); lock.waitLock(20000);
  try {
    if (opSeen_(payload.op_id)) return { ok: true, dup: true };
    var sh = projSheetOfRow_(payload.row_id); if (!sh) throw new Error('프로젝트 없음');
    var idx = ensureCheckCols_(sh), rownum = findRow_(sh, idx, payload.row_id);
    if (rownum < 0) throw new Error('행 없음: ' + payload.row_id);
    assertCanEditStage_(me, sh.getRange(rownum, idx['작업자'] + 1).getValue(), sh.getRange(rownum, idx['검수자'] + 1).getValue(), 2);
    var cand = String(sh.getRange(rownum, idx['신어 후보'] + 1).getValue());
    var prev = String(sh.getRange(rownum, idx['상태'] + 1).getValue()).trim();
    var prevVerdict = String(sh.getRange(rownum, idx['2차 판별'] + 1).getValue()).trim();
    var prevTs = String(sh.getRange(rownum, idx['2차 일시'] + 1).getValue()).trim();
    setCell_(sh, rownum, idx, '2차 판별', payload.verdict);
    setCell_(sh, rownum, idx, '2차 메모', payload.memo || '');
    setCell_(sh, rownum, idx, '2차 고유명 여부', payload.proper);
    setCell_(sh, rownum, idx, '2차 혐오 표현 여부', payload.hate);
    if (!(prevVerdict === payload.verdict && prevTs)) setCell_(sh, rownum, idx, '2차 일시', now_());   // 같은 판별 재저장(체크 추가)이면 원래 일시 보존
    setCell_(sh, rownum, idx, '상태', STATUS.SECOND);
    appendLog_(me, payload.row_id, cand, '2차', payload.verdict + ' (고유명' + payload.proper + '/혐오' + payload.hate + ')', payload.memo || '', prev, STATUS.SECOND);
    SpreadsheetApp.flush(); opMark_(payload.op_id);
    return { ok: true };
  } finally { lock.releaseLock(); }
}
// ── 집필 저장 점검(프론트와 동일 규칙, 전부 오류=저장 차단). 어깨 번호 동형어 점검은 목록이 필요해 프론트에서만 수행. ──
var W_FIELD_67 = ['교육','문학','민속','언어','역사','철학','인문 일반','경영','경제','군사','매체','법률','복지','심리','정치','행정','사회 일반','동물','물리','생명','수학','식물','지구','지리','천문','천연자원','해양','화학','환경','자연 일반','공업','광업','농업','서비스업','수산업','임업','산업 일반','수의','식품','약학','의학','한의','보건 일반','건설','교통','기계','재료','전기ㆍ전자','정보ㆍ통신','공학 일반','체육','연기','영상','무용','음악','미술','복식','공예','예체능 일반','가톨릭','기독교','불교','종교 일반','인명','지명','책명','고유명 일반'];
var W_FIELD_15 = ['인간','삶','식생활','의생활','주생활','사회생활','경제생활','교육','종교','문화','정치와 행정','자연','동식물','개념','보건ㆍ의학','정보ㆍ기술'];
var W_POS = ['명','부','동','형','감','대','수','관','의명','보동','보형','관ㆍ명','수ㆍ관','명ㆍ부','감ㆍ명','대ㆍ부','대ㆍ감','동ㆍ형','관ㆍ감','부ㆍ감','수ㆍ관ㆍ명','대ㆍ관'];   // 대표 음절 표기
var W_LANG = ['고','한','그','네','노','독','라','러','루','말','몽','베','불','산','세','와','스','아','영','에','이','인','일','중','체','타','터','페','포','폴','프','헝','히','힌'];
var HATE_NOTE = '차별 및 비하하는 의미가 포함되어 있으므로 사용에 주의가 필요하다.';
var W_TPL = { '원어': '()', '어원': '【】', '용례': '¶', '참고 용례': '¶' };
function wBlank_(f, v) { v = String(v == null ? '' : v).trim(); return !v || v === (W_TPL[f] || ' '); }
function validateWrite_(part, v, hateY) {
  var E = []; function has(s, ch) { return String(s || '').indexOf(ch) >= 0; } function S(k) { return String(v[k] == null ? '' : v[k]).trim(); }
  if (part === '형태부') {
    var wg = S('단어/구'), wf = S('조어법'), pos = S('품사'), reg = S('일상어/전문어'), fld = S('전문 분야'), head = S('등재표제어'), native = S('고유어') === 'Y';
    var orig = wBlank_('원어', v['원어']) ? '' : S('원어'), et = S('어종 표시'), ety = wBlank_('어원', v['어원']) ? '' : S('어원');
    var ix = S('색인표제어');
    if (!ix) E.push('색인표제어가 비어 있습니다');
    else if (/[^가-힣ㆍ0-9]/.test(ix)) E.push('색인표제어에 쓸 수 없는 문자가 있습니다 (한글, ㆍ, 어깨 번호만 가능)');
    if (!head) E.push('등재표제어가 비어 있습니다');
    else {
      if (/[^가-힣ㄱ-ㅎㅏ-ㅣ\-\^ㆍ\s0-9]/.test(head)) E.push('등재표제어에 쓸 수 없는 문자가 있습니다 (한글, -, ^, ㆍ, 공백만 가능)');
      if (has(head, '^') && /\s/.test(head)) E.push('등재표제어에 ^과 공백을 함께 썼습니다');
      if ((head.match(/-/g) || []).length >= 2) E.push('등재표제어에 -이 두 개 이상입니다');
      if (wg === '단어' && (/\s/.test(head) || has(head, '^'))) E.push('단어인데 등재표제어에 공백이나 ^이 있습니다');
      if (wg === '구' && has(head, '-')) E.push('구인데 등재표제어에 -이 있습니다');
      if (wg === '구' && !/\s/.test(head) && !has(head, '^')) E.push('구인데 등재표제어에 공백이나 ^이 없습니다');
      if ((wf === '혼성' || wf === '축약') && has(head, '-')) E.push('조어법이 ' + wf + '인데 등재표제어에 -이 있습니다');
    }
    if (wg === '단어' && !pos) E.push('단어인데 품사가 비어 있습니다');
    if (reg === '전문어' && !fld) E.push('전문어인데 전문 분야가 비어 있습니다');
    if (wf === '차용') { if (has(orig, '▼')) E.push('차용어에는 원어에 ▼를 쓸 수 없습니다'); if (!ety) E.push('차용어인데 어원이 비어 있습니다'); }
    if (!native) {
      if (!orig) E.push('원어가 비어 있습니다 (고유어면 고유어를 체크)');
      else { if (!/^\(.+\)$/.test(orig)) E.push('원어를 ( ) 안에 쓰지 않았습니다');
        if (/[^가-힣ㄱ-ㅎㅏ-ㅣ一-鿿㐀-䶿A-Za-z←▼▽<>\[\]()\/\s0-9\-]/.test(orig)) E.push('원어에 쓸 수 없는 문자가 있습니다 (^, + 등)');
        if (/(^|[^A-Za-z])-|-([^A-Za-z]|$)/.test(orig)) E.push('원어의 -는 로마자 사이에만 쓸 수 있습니다 (on-line)'); }
    }
    if (!et) E.push('어종 표시가 비어 있습니다');
    else {
      et.replace(/[+_^()\s]/g, '|').split('|').filter(function (x) { return x; }).forEach(function (t) { if (!/^[가-힣]$/.test(t)) E.push('어종 표시에 쓸 수 없는 값이 있습니다: ' + t); });   // 약어는 한글 1음절(직접 입력 언어 포함)
      // 성분이 하나인 단어(혼성·축약 등)는 어종이 형태소를 +로 나열하므로 대조하지 않음. 여럿이면 괄호 안을 뺀 상위 기호만 대조
      var seqH = (head || '').replace(/\d+$/, '').replace(/[^-\s^]/g, '').replace(/-/g, '+').replace(/\s/g, '_'), seqE = et.replace(/\([^)]*\)/g, 'X').replace(/[^+_^]/g, '');
      if (head && seqH && seqH !== seqE) E.push('어종 표시의 기호(+ _ ^)가 등재표제어의 분절(- 공백 ^)과 맞지 않습니다');
    }
  } else {
    var def = S('뜻풀이'), ex = wBlank_('용례', v['용례']) ? '' : S('용례'), xy = S('X년 Y월 신어');
    if (!def) E.push('뜻풀이가 비어 있습니다');
    if (xy && !/^【\d{4}년 \d{1,2}월 신어】$/.test(xy)) E.push('X년 Y월 신어가 【OOOO년 OO월 신어】 형식이 아닙니다');
    function romanOut(t) { return /[A-Za-z]/.test(String(t || '').replace(/\([^)]*\)/g, '')); }
    if (romanOut(def)) E.push('뜻풀이에 괄호 밖 로마자가 있습니다');
    if (romanOut(ex)) E.push('용례에 괄호 밖 로마자가 있습니다');
    if (hateY && def.indexOf(HATE_NOTE) < 0) E.push('혐오 표현 항목인데 뜻풀이 끝에 주의 문구가 없습니다');
  }
  return E;
}
// 집필 저장: payload={row_id, phase('형태부'|'의미부'), stage(1|2), fields:{형태부:{…}, 의미부:{…}}, op_id}
//  - 프로젝트 단계와 payload.phase 일치 필수. 형태부 단계에서는 형태부만, 의미부 단계에서는 형태부+의미부를 한 번에 저장.
//  - 각 부의 값은 '{stage}차 {필드}' 컬럼에 기록. 상태·일시는 현재 단계의 부만 갱신.
function saveWrite(token, payload) {
  var me = me_(token);
  var lock = LockService.getDocumentLock(); lock.waitLock(20000);
  try {
    if (opSeen_(payload.op_id)) return { ok: true, dup: true };
    var pid = String(payload.row_id || '').split('::')[0], p = projById_(pid);
    if (!p || !isWriteKind_(p.type)) throw new Error('집필 프로젝트가 아닙니다.');
    var phase = String(payload.phase || '').trim();
    if (phase !== p.phase) throw new Error('프로젝트 단계가 ' + p.phase + '(으)로 바뀌었습니다. 화면을 새로고침하세요.');
    var parts = phase === '형태부' ? ['형태부'] : WRITE_PARTS;   // 점검·상태 대상
    var sh = projItemSheet_(pid); if (!sh) throw new Error('프로젝트 없음');
    var idx = ensureCols_(sh, HEADERS_WRITE), rownum = findRow_(sh, idx, payload.row_id);
    if (rownum < 0) throw new Error('행 없음: ' + payload.row_id);
    var stage = (parseInt(payload.stage, 10) === 2) ? 2 : 1;
    var row = sh.getRange(rownum, 1, 1, sh.getLastColumn()).getValues()[0];
    function cur(k) { return idx[k] != null ? String(row[idx[k]] == null ? '' : row[idx[k]]).trim() : ''; }
    if (!canEditWrite_(me, p.type, cur('작업자'), cur('검수자'), stage)) throw new Error('권한 없음: 배정된 담당자만 입력할 수 있습니다.');
    var fields = payload.fields || {}, errs = [];
    parts.forEach(function (pt) { validateWrite_(pt, fields[pt] || {}, cur('2차 혐오 표현 여부') === 'Y').forEach(function (e) { errs.push((parts.length > 1 ? pt + ' · ' : '') + e); }); });
    if (errs.length) throw new Error('형식 오류 ' + errs.length + '건 존재. ' + errs.join(' / '));
    WRITE_PARTS.forEach(function (pt) {   // 단계 외 부(형태부 단계의 의미부 복사값)도 값이 오면 기록. 점검·상태는 위에서 단계 부만
      if (!fields[pt]) return;
      var cols = writeStageCols_(pt, stage), fv = fields[pt] || {};
      cols.forEach(function (c) { var f = c.slice(3); if (f in fv) setCell_(sh, rownum, idx, c, fv[f] == null ? '' : fv[f]); });
    });
    var stKey = phase + ' 상태', prev = cur(stKey) || STATUS.NONE, ns;
    if (stage === 1) ns = (prev === STATUS.SECOND) ? STATUS.SECOND : STATUS.FIRST; else ns = STATUS.SECOND;
    setCell_(sh, rownum, idx, stKey, ns); setCell_(sh, rownum, idx, phase + ' ' + stage + '차 일시', now_());
    var summ = phase === '형태부' ? String((fields['형태부'] || {})['등재표제어'] || '') : String((fields['의미부'] || {})['뜻풀이'] || '');
    appendLog_(me, payload.row_id, cur('신어 후보'), '집필 ' + phase + ' ' + stage + '차', summ.slice(0, 40), '', prev, ns);
    SpreadsheetApp.flush(); opMark_(payload.op_id);
    return { ok: true };
  } finally { lock.releaseLock(); }
}
// 집필 새 항목 추가(표제어만). 작업자=본인(관리자는 비움)
function addWriteItem(token, projectId, cand) {
  var me = me_(token);
  cand = String(cand || '').trim(); if (!cand) throw new Error('색인표제어를 입력해 주세요.');
  var p = projById_(projectId); if (!p || !isWriteKind_(p.type)) throw new Error('집필 프로젝트가 아닙니다.');
  var sh = projItemSheet_(projectId); if (!sh) throw new Error('시트 없음');
  var idx = ensureCols_(sh, HEADERS_WRITE), rid = projectId + '::new-' + Date.now().toString(36), row = [];
  for (var k in idx) row[idx[k]] = '';
  row[idx['ID']] = rid; row[idx['신어 후보']] = cand; row[idx['작업 구분']] = p.type;
  row[idx['형태부 상태']] = STATUS.NONE; row[idx['의미부 상태']] = STATUS.NONE;
  row[idx['작업자']] = me.isManager ? '' : me.name;
  sh.appendRow(row); SpreadsheetApp.flush();
  return { id: rid, cand: cand };
}
// 집필 새 항목 삭제 — 새로 추가한 항목(::new-)만 허용, 배분된 항목은 불가
function deleteWriteItem(token, rowId) {
  me_(token);
  if (String(rowId || '').indexOf('::new-') < 0) throw new Error('배분된 항목은 삭제할 수 없습니다.');
  var sh = projSheetOfRow_(rowId); if (!sh) throw new Error('프로젝트 없음');
  var idx = headerIndex_(sh), rownum = findRow_(sh, idx, rowId);
  if (rownum < 0) throw new Error('행 없음: ' + rowId);
  sh.deleteRow(rownum); SpreadsheetApp.flush();
  return { ok: true };
}

// ── 로그 ───────────────────────────────────────────────
function logSheet_() {
  var sh = ss_().getSheetByName(SHEET_LOG);
  if (!sh) { sh = ss_().insertSheet(SHEET_LOG); sh.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]); styleHeader_(sh, LOG_HEADERS.length); }
  return sh;
}
function appendLog_(me, rowId, cand, stage, verdict, memo, prevStatus, newStatus) {
  var row = [now_(), rowId, cand, stage, me.name, verdict, memo, prevStatus, newStatus];
  try { logSheet_().appendRow(row); return; } catch (e) {}
  // 폴백: 관리 시트 셀 한계로 append 실패 → 별도 로그 파일로 롤오버(그 파일도 차면 다음 번호). 로그 유실 방지.
  for (var k = 0; k < 3; k++) { try { overflowLogSheet_(k > 0).appendRow(row); return; } catch (e2) {} }
}
function overflowLogSheet_(roll) {
  var props = PropertiesService.getScriptProperties(), n = parseInt(props.getProperty('overflow_log_n') || '1', 10);
  if (roll) { n++; props.setProperty('overflow_log_n', String(n)); }
  var key = 'overflow_log_id_' + n, id = props.getProperty(key), lss = null;
  if (id) { try { lss = SpreadsheetApp.openById(id); } catch (e) { lss = null; } }
  if (!lss) {
    lss = SpreadsheetApp.create('KNO_변경로그_오버플로_' + n);
    try { var f = DriveApp.getFileById(lss.getId()); folder_([DRIVE_ROOT]).addFile(f); DriveApp.getRootFolder().removeFile(f); } catch (e) {}
    var s0 = lss.getSheets()[0]; s0.setName(SHEET_LOG); s0.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]); styleHeader_(s0, LOG_HEADERS.length);
    props.setProperty(key, lss.getId());
  }
  return lss.getSheets()[0];
}
function failLogSheet_() {
  var sh = ss_().getSheetByName(SHEET_FAILLOG);
  if (!sh) { sh = ss_().insertSheet(SHEET_FAILLOG); sh.getRange(1, 1, 1, FAILLOG_HEADERS.length).setValues([FAILLOG_HEADERS]); styleHeader_(sh, FAILLOG_HEADERS.length); }
  return sh;
}
// 클라이언트가 저장 후 행을 다시 읽어 '확정 미저장'으로 판정하면 호출 → 실패 로그 1행.
function logClientFail(token, info) {
  var name = '(미확인)';
  try { var who = whoByToken_(token); if (who.known) name = who.name || who.email; } catch (e) {}
  try { info = info || {};
    failLogSheet_().appendRow([now_(), name, (SAVE_FNS[info.fn] || info.fn || '') + '(클라)', String(info.row_id || ''),
      String(info.verdict || ''), String(info.memo || ''), String(info.reason || '클라이언트 확정 미저장'), token ? String(token).slice(0, 6) + '…' : '']);
  } catch (e) {}
  return { ok: true };
}

// ── CSV 헬퍼(exportProject 백업 다운로드용) ────────────
function toCsv_(rows) { return rows.map(function (r) { return r.map(csvCell_).join(','); }).join('\r\n'); }
function csvCell_(v) { if (v instanceof Date) v = Utilities.formatDate(v, TZ, 'yyyyMMdd HHmmss'); v = String(v == null ? '' : v); return /[",\r\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }

// ── 연구원 명단(관리자) ────────────────────────────────
function getResearchers(token) {
  assertManager_(me_(token));
  var sh = sheet_(SHEET_USERS), n = sh.getLastRow() - 1; if (n <= 0) return [];
  var lc = sh.getLastColumn(), data = sh.getRange(2, 1, n, lc).getValues(), url = getAppUrl_(), out = [];
  for (var i = 0; i < data.length; i++) {
    var tok = String(data[i][6] || '').trim(), email = String(data[i][0]).trim();
    out.push({ email: email, name: String(data[i][1]).trim(), role: String(data[i][2]).trim(),
      '소속': lc > 10 ? String(data[i][10] || '').trim() : '', '성별': lc > 11 ? String(data[i][11] || '').trim() : '',
      token: tok, link: personalLink_(url, tok, email), id: String(data[i][8] || '').trim() });
  }
  return out;
}
// 목록으로 시트 재작성. 계정정보(token/링크/아이디/비번)는 이메일 기준 보존.
function saveResearchers(token, list) {
  assertManager_(me_(token));
  list = list || [];
  var sh = sheet_(SHEET_USERS), prev = {}, n = sh.getLastRow() - 1;
  if (n > 0) { var lc = sh.getLastColumn(), old = sh.getRange(2, 1, n, lc).getValues();
    for (var i = 0; i < old.length; i++) { var em = String(old[i][0]).trim().toLowerCase();
      if (em) prev[em] = { token: old[i][6] || '', link: old[i][7] || '', id: old[i][8] || '', pw: old[i][9] || '' }; } }
  var rows = [];
  list.forEach(function (r) { var em = String(r.email || '').trim(), pv = prev[em.toLowerCase()] || {};
    rows.push([em, String(r.name || '').trim(), String(r.role || '').trim(), '', '', '', pv.token || '', pv.link || '', pv.id || '', pv.pw || '', String(r['소속'] || '').trim(), String(r['성별'] || '').trim()]); });
  sh.clearContents();
  sh.getRange(1, 1, 1, USER_HDR.length).setValues([USER_HDR]);
  if (rows.length) sh.getRange(2, 1, rows.length, USER_HDR.length).setValues(rows);
  styleHeader_(sh, USER_HDR.length);
  cacheDel_('presence:roster');
  return { ok: true, count: rows.length };
}
// ── 편집기 전용: 연구원 명단 15명 고정 시드 ────────────
// 편집기에서 setupInit 실행 → 연구원 시트를 이 명단으로 세팅.
// 기존 계정정보(token/링크/아이디/비번)는 이메일 기준 보존하므로 재실행해도 안전.
var SEED_ROSTER = [
  { email: 'nki@yonsei.ac.kr', name: '남길임', role: '검수자', '소속': '연세대학교', '성별': '여성' },
  { email: 'songhj@knu.ac.kr', name: '송현주', role: '검수자', '소속': '경북대학교', '성별': '여성' },
  { email: 'cjuni2000@gmail.com', name: '최준', role: '검수자', '소속': '전남대학교', '성별': '남성' },
  { email: 'fbih02@gmail.com', name: '현영희', role: '검수자', '소속': '경북대학교', '성별': '여성' },
  { email: 'sjmano27@gmail.com', name: '이수진', role: '검수자', '소속': '경북대학교', '성별': '여성' },
  { email: 'bmg0128@gmail.com', name: '백미경', role: '검수자', '소속': '경북대학교', '성별': '여성' },
  { email: 'chunghaeyun1006@gmail.com', name: '정해윤', role: '검수자', '소속': '연세대학교', '성별': '여성' },
  { email: 'leejun0624@gmail.com', name: '이준', role: '작업자', '소속': '연세대학교', '성별': '남성' },
  { email: 'a01082406803@gmail.com', name: '김유정', role: '작업자', '소속': '전남대학교', '성별': '여성' },
  { email: 'saenu@yonsei.ac.kr', name: '김선우', role: '작업자', '소속': '연세대학교', '성별': '여성' },
  { email: 'goyelin08@gmail.com', name: '고예린', role: '작업자', '소속': '전남대학교', '성별': '여성' },
  { email: 'qhal7041@gmail.com', name: '김보미', role: '작업자', '소속': '전남대학교', '성별': '여성' },
  { email: 'sul010907@gmail.com', name: '남궁설', role: '작업자', '소속': '연세대학교', '성별': '여성' },
  { email: 'siveking@gmail.com', name: '안진산', role: '검수자', '소속': '경북대학교', '성별': '남성' },
  { email: 'koreanneology@gmail.com', name: '관리자', role: '관리자', '소속': '경북대학교', '성별': '남성' }
];
function setupInit() {
  var sh = ss_().getSheetByName(SHEET_USERS) || ss_().insertSheet(SHEET_USERS);
  var prev = {}, n = sh.getLastRow() - 1;
  if (n > 0) { var lc = sh.getLastColumn(), old = sh.getRange(2, 1, n, lc).getValues();
    for (var i = 0; i < old.length; i++) { var em = String(old[i][0]).trim().toLowerCase();
      if (em) prev[em] = { token: old[i][6] || '', link: old[i][7] || '', id: old[i][8] || '', pw: old[i][9] || '' }; } }
  var url = getAppUrl_(), rows = SEED_ROSTER.map(function (r) {
    var pv = prev[r.email.toLowerCase()] || {};
    var token = pv.token || Utilities.getUuid().replace(/-/g, '').slice(0, 12);
    var link = pv.link || personalLink_(url, token, r.email);
    return [r.email, r.name, r.role, '', '', '', token, link, pv.id || '', pv.pw || '', r['소속'], r['성별']];
  });
  sh.clearContents();
  sh.getRange(1, 1, 1, USER_HDR.length).setValues([USER_HDR]);
  sh.getRange(2, 1, rows.length, USER_HDR.length).setValues(rows);
  styleHeader_(sh, USER_HDR.length);
  cacheDel_('presence:roster');
  SpreadsheetApp.flush();
  for (var j = 0; j < rows.length; j++) Logger.log(rows[j][1] + ' (' + rows[j][2] + '): ' + rows[j][7]);
  return SEED_ROSTER.length + '명 시드 완료. 위 로그에서 각자 개인 링크 확인(관리자=koreanneology).';
}
