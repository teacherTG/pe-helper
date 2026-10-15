/**
 * 체육 업무 도우미 — 구글 시트 백엔드 (Google Apps Script 웹 앱 API)
 *
 * 프론트엔드(GitHub Pages)가 이 웹 앱 주소로 POST 요청을 보내고, 데이터는 이 스크립트가 붙은 구글 시트에 저장됩니다.
 *
 * 설치
 *   1) 구글 시트 → [확장 프로그램] → [Apps Script] 에 이 파일 내용을 붙여넣기
 *   2) 아래 API_KEY 를 나만 아는 긴 문자열로 바꾸기 (저장소에 올리지 마세요)
 *   3) [배포] → [새 배포] → 웹 앱 / 실행 사용자 "나" / 액세스 권한 "모든 사용자" → 배포
 *      (GitHub 페이지에서 부르려면 "모든 사용자"여야 합니다. 대신 API_KEY 가 없으면 아무것도 읽거나 쓸 수 없습니다.)
 *   4) 웹 앱 주소(…/exec)와 API_KEY 를 앱의 [시트 연결] 창에 넣기
 *
 * 저장 구조
 *   _data 시트(숨김) : 앱 원본 데이터(JSON). 한 줄에 반 하나, 업무 목록 하나.
 *   "1학년 3반" 같은 시트 : 사람이 보는 점수표 + 관찰 메모 + 교과세특 (저장할 때마다 새로 씀)
 *   "업무관리" 시트 : 업무 목록
 */

const API_KEY = 'CHANGE_ME';   // ← 반드시 바꾸세요

const DATA_SHEET = '_data';
const CHUNK = 40000;          // 셀 하나 최대 50,000자 → 4만 자씩 나눠 저장
const PREFIX = '~';           // 수식(=)이나 숫자로 바뀌지 않게 붙이는 앞글자

function doGet(e) {
  return json_({ ok: true, service: '체육 업무 도우미 API', hint: 'POST {key, action} 로 호출하세요.' });
}

function doPost(e) {
  let p = {};
  try { p = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return json_({ ok: false, error: 'bad json' }); }
  if (API_KEY === 'CHANGE_ME' || p.key !== API_KEY) return json_({ ok: false, error: 'unauthorized' });
  try {
    switch (p.action) {
      case 'ping': return json_({ ok: true, data: 'pong' });
      case 'loadAll': return json_({ ok: true, data: loadAll_() });
      case 'saveDoc': saveDoc_(String(p.id), String(p.json || ''), p.view); return json_({ ok: true });
      case 'deleteDoc': deleteDoc_(String(p.id), p.viewName); return json_({ ok: true });
      default: return json_({ ok: false, error: 'unknown action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function dataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(DATA_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DATA_SHEET);
    sh.getRange(1, 1, 1, 3).setValues([['id', 'updated', 'chunks']]);
    sh.hideSheet();
  }
  return sh;
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return 0;
}

function loadAll_() {
  const sh = dataSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const width = sh.getLastColumn();
  const vals = sh.getRange(2, 1, last - 1, width).getDisplayValues();
  return vals.filter(r => r[0]).map(r => {
    const n = Number(r[2] || 0);
    const json = r.slice(3, 3 + n).map(s => String(s).replace(/^~/, '')).join('');
    return { id: String(r[0]), json };
  });
}

function saveDoc_(id, json, view) {
  if (!id || !/^[\w.\-]+$/.test(id)) throw new Error('bad id');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = dataSheet_();
    const chunks = [];
    for (let i = 0; i < json.length; i += CHUNK) chunks.push(PREFIX + json.slice(i, i + CHUNK));
    const row = [id, new Date(), chunks.length].concat(chunks);
    if (sh.getMaxColumns() < row.length) sh.insertColumnsAfter(sh.getMaxColumns(), row.length - sh.getMaxColumns());
    let r = findRow_(sh, id);
    if (!r) r = sh.getLastRow() + 1;
    else sh.getRange(r, 1, 1, sh.getLastColumn()).clearContent();
    if (chunks.length) sh.getRange(r, 4, 1, chunks.length).setNumberFormat('@');
    sh.getRange(r, 1, 1, row.length).setValues([row]);
    if (view && view.name && Array.isArray(view.rows) && view.rows.length) writeView_(view);
  } finally {
    lock.releaseLock();
  }
}

function deleteDoc_(id, viewName) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = dataSheet_();
    const r = findRow_(sh, id);
    if (r) sh.deleteRow(r);
    const name = safeName_(viewName);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const v = name && ss.getSheetByName(name);
    if (v && ss.getSheets().length > 1) ss.deleteSheet(v);
  } finally {
    lock.releaseLock();
  }
}

function safeName_(name) {
  return String(name || '').replace(/[\[\]\*\?\/\\:]/g, ' ').trim().slice(0, 90);
}

function writeView_(view) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = safeName_(view.name);
  if (!name || name === DATA_SHEET || name === '안내') return;
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  const width = Math.max.apply(null, view.rows.map(r => r.length));
  const rows = view.rows.map(r => {
    const out = r.map(v => v == null ? '' : String(v));
    while (out.length < width) out.push('');
    return out;
  });
  sh.clearContents();
  if (sh.getMaxColumns() < width) sh.insertColumnsAfter(sh.getMaxColumns(), width - sh.getMaxColumns());
  const range = sh.getRange(1, 1, rows.length, width);
  range.setNumberFormat('@');           // 1:04.5 같은 기록이 시간·날짜로 바뀌지 않게
  range.setValues(rows);
  sh.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('#e8eef5');
  sh.setFrozenRows(1);
  const col = rows[0].indexOf('교과세특') + 1;
  if (col > 0) {
    sh.setColumnWidth(col, 520);
    sh.getRange(2, col, Math.max(rows.length - 1, 1), 1).setWrap(true);
  }
  if (name === '업무관리') sh.setColumnWidth(3, 420);
}
