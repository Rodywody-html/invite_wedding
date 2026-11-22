/**
 * 모바일 청첩장 백엔드 (Google Apps Script)
 * - 구글 시트에 연결된 스크립트(시트 > 확장 프로그램 > Apps Script)에 붙여 넣어 사용
 * - 저장: 참석 응답(RSVP), 방명록(Guestbook), 추가 사진(Photos), 숨긴 기본사진(Hidden)
 * - 관리자 비밀번호: 코드에 적지 말고 [프로젝트 설정 > 스크립트 속성]에 ADMIN_PW 로 직접 입력
 * - 배포: 배포 > 새 배포 > 유형 "웹 앱" > 실행 사용자 "나" / 액세스 "모든 사용자"
 */

const LIMIT = { name: 20, msg: 300, companion: 100 };
const TOKEN_TTL = 6 * 60 * 60; // 관리자 로그인 유지 시간(초), 최대 6시간

function doGet(e) {
  try {
    const a = (e.parameter || {}).action;
    if (a === 'guestbook_list') return out({ ok: true, items: guestbookList() });
    if (a === 'photos_list') return out({ ok: true, added: photosList(), hidden: hiddenList() });
    return out({ ok: true, msg: 'wedding api' });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const b = JSON.parse(e.postData.contents);
    switch (b.action) {
      // ---- 하객 공개 기능 ----
      case 'rsvp_add': return out(rsvpAdd(b));
      case 'guestbook_add': return out(guestbookAdd(b));
      // ---- 관리자 ----
      case 'admin_login': return out(adminLogin(b));
      case 'rsvp_list': requireAdmin(b); return out({ ok: true, items: rsvpList() });
      case 'guestbook_delete': requireAdmin(b); return out(deleteRowById('Guestbook', b.id));
      case 'photo_add': requireAdmin(b); return out(photoAdd(b));
      case 'photo_remove': requireAdmin(b); return out(photoRemove(b));
      case 'default_hide': requireAdmin(b); return out(defaultSetHidden(b.key, true));
      case 'default_show': requireAdmin(b); return out(defaultSetHidden(b.key, false));
      default: return out({ ok: false, error: 'unknown action' });
    }
  } catch (err) {
    return out({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* ---------- 공통 ---------- */
function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function sheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(headers); sh.setFrozenRows(1); }
  return sh;
}
function clip(s, n) { return String(s == null ? '' : s).trim().slice(0, n); }
function now() { return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'); }
function sanitizeCell(s) { return /^[=+\-@]/.test(s) ? "'" + s : s; } // 시트 수식 주입 방지

/* ---------- 참석 응답 ---------- */
const RSVP_H = ['id', '시간', '구분', '성함', '참석여부', '참석인원', '동행인', '식사여부'];
function rsvpAdd(b) {
  const side = b.side === '신부측' ? '신부측' : '신랑측';
  const name = clip(b.name, LIMIT.name);
  if (!name) return { ok: false, error: '성함을 입력해 주세요.' };
  const attend = b.attend === '불참' ? '불참' : '참석';
  let count = Math.max(0, Math.min(20, parseInt(b.count, 10) || 0));
  if (attend === '불참') count = 0;
  const meal = ['식사함', '식사 안 함', '미정'].indexOf(b.meal) >= 0 ? b.meal : (attend === '불참' ? '-' : '미정');
  sheet('RSVP', RSVP_H).appendRow([
    Utilities.getUuid(), now(), side, sanitizeCell(name), attend, count,
    sanitizeCell(clip(b.companion, LIMIT.companion)), attend === '불참' ? '-' : meal
  ]);
  return { ok: true };
}
function rsvpList() {
  const v = sheet('RSVP', RSVP_H).getDataRange().getValues();
  return v.slice(1).map(r => ({ id: r[0], time: r[1], side: r[2], name: r[3], attend: r[4], count: r[5], companion: r[6], meal: r[7] }));
}

/* ---------- 방명록 ---------- */
const GB_H = ['id', '시간', '이름', '메시지'];
function guestbookAdd(b) {
  const name = clip(b.name, LIMIT.name), msg = clip(b.msg, LIMIT.msg);
  if (!name || !msg) return { ok: false, error: '이름과 메시지를 입력해 주세요.' };
  const id = Utilities.getUuid();
  sheet('Guestbook', GB_H).appendRow([id, now(), sanitizeCell(name), sanitizeCell(msg)]);
  return { ok: true, id: id };
}
function guestbookList() {
  const v = sheet('Guestbook', GB_H).getDataRange().getValues();
  return v.slice(1).map(r => ({ id: r[0], time: r[1], name: r[2], msg: r[3] }));
}
function deleteRowById(sheetName, id) {
  const sh = sheet(sheetName, ['id']);
  const v = sh.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) if (v[i][0] === id) { sh.deleteRow(i + 1); return { ok: true }; }
  return { ok: false, error: '항목을 찾을 수 없어요.' };
}

/* ---------- 관리자 인증 ---------- */
function adminLogin(b) {
  const cache = CacheService.getScriptCache();
  if (Number(cache.get('fails') || 0) >= 5) return { ok: false, error: '시도 횟수 초과. 5분 뒤 다시 시도해 주세요.' };
  const pw = PropertiesService.getScriptProperties().getProperty('ADMIN_PW');
  if (!pw) return { ok: false, error: '스크립트 속성에 ADMIN_PW 가 설정되지 않았어요.' };
  if (String(b.password || '') !== pw) {
    cache.put('fails', String(Number(cache.get('fails') || 0) + 1), 300);
    return { ok: false, error: '비밀번호가 맞지 않아요.' };
  }
  cache.remove('fails');
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('t_' + token, '1', TOKEN_TTL);
  return { ok: true, token: token };
}
function requireAdmin(b) {
  if (!b.token || !CacheService.getScriptCache().get('t_' + b.token)) throw new Error('관리자 로그인이 필요하거나 만료됐어요.');
}

/* ---------- 사진 (추가 사진은 드라이브 폴더에 저장) ---------- */
const PH_H = ['id', '시간', '파일ID'];
function photoFolder() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('PHOTO_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder('wedding_invite_photos');
  props.setProperty('PHOTO_FOLDER_ID', f.getId());
  return f;
}
function photoAdd(b) {
  const m = /^data:(image\/(?:jpeg|png|webp));base64,(.+)$/.exec(b.data || '');
  if (!m) return { ok: false, error: '이미지 형식이 올바르지 않아요.' };
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'photo_' + Date.now() + '.jpg');
  const file = photoFolder().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const id = Utilities.getUuid();
  sheet('Photos', PH_H).appendRow([id, now(), file.getId()]);
  return { ok: true, id: id, fileId: file.getId() };
}
function photosList() {
  const v = sheet('Photos', PH_H).getDataRange().getValues();
  return v.slice(1).map(r => ({ id: r[0], fileId: r[2] }));
}
function photoRemove(b) {
  const sh = sheet('Photos', PH_H);
  const v = sh.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) {
    if (v[i][0] === b.id) {
      try { DriveApp.getFileById(v[i][2]).setTrashed(true); } catch (e) {}
      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: '사진을 찾을 수 없어요.' };
}
/* 기본(저장소에 포함된) 사진 숨기기/되살리기 */
function hiddenList() {
  const v = sheet('Hidden', ['key']).getDataRange().getValues();
  return v.slice(1).map(r => r[0]);
}
function defaultSetHidden(key, hide) {
  key = clip(key, 20);
  if (!/^p\d{2}$/.test(key)) return { ok: false, error: 'invalid key' };
  const sh = sheet('Hidden', ['key']);
  const v = sh.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < v.length; i++) if (v[i][0] === key) row = i + 1;
  if (hide && row < 0) sh.appendRow([key]);
  if (!hide && row > 0) sh.deleteRow(row);
  return { ok: true };
}
