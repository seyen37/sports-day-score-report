/**
 * 體育表演會 成績回報系統（後端）
 * 設計者：臺北市福林國小 許士彥 2026.11.21.
 * 使用方式：綁定在一份 Google 試算表上，先執行一次 setup()，再部署為網頁應用程式。
 */
const VERSION = 'v1.22.1'; // 後端版本，改版時與 Index.html 的 APP_VERSION 一起更新
const TZ = 'Asia/Taipei';
const SH = { CASE: '案件', ROSTER: '名單', ROLE: '角色', CONF: '設定', MERGE: '獎狀合併列印', BACKUP: '案件備份' };
const CASE_H = ['案件編號', '送出時間', '類別', '項目', '比賽日期', '比賽時間', '檢錄員', '填報帳號', '急件',
  '成績摘要', '狀態', '負責人', '處理結果', '完成時間', '播報時間', '稽核', '稽核意見', '稽核時間',
  '照片檔案ID', '成績資料', '處理紀錄', '列印時間'];
const COL = {}; CASE_H.forEach((h, i) => { COL[h] = i; });
const STATUS = ['待處理', '處理中', '已完成'];
const ROLES = ['行政', '司儀', '校長'];
const GRADE_NUM = { '一': '1', '二': '2', '三': '3', '四': '4', '五': '5', '六': '6' };
const ROSTER_H = ['類別', '項目', '班級', '座號', '姓名'];


/* ───────── 網頁進入點 ───────── */
function doGet() {
  let title = CONF_DEFAULT['系統名稱'];
  try { title = conf_()['系統名稱'] || title; } catch (e) { /* 尚未初始化時用預設名稱 */ }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ───────── 初始化（只需執行一次，可重複執行，不會覆蓋既有資料） ───────── */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SH.CASE);
  if (!sh) { sh = ss.insertSheet(SH.CASE); }
  if (sh.getLastRow() === 0) { writeRows_(sh, 1, [CASE_H]); sh.setFrozenRows(1); }

  sh = ss.getSheetByName(SH.ROSTER);
  if (!sh) { sh = ss.insertSheet(SH.ROSTER); }
  if (sh.getLastRow() === 0) {
    // 名冊不內建在程式裡：這裡只建立標題列，選手名單請到網站的「名冊管理」匯入，比賽項目可在「系統設定」新增
    const rows = [ROSTER_H];
    writeRows_(sh, 1, rows); sh.setFrozenRows(1);
  }

  sh = ss.getSheetByName(SH.ROLE);
  if (!sh) { sh = ss.insertSheet(SH.ROLE); }
  if (sh.getLastRow() === 0) {
    writeRows_(sh, 1, [
      ['身分', '姓名', 'Email（選填，同網域登入時自動辨識）', '通行碼'],
      ['行政', '獎狀列印', '', pin_()],
      ['司儀', '司儀', '', pin_()],
      ['校長', '校長', '', pin_()]
    ]);
    sh.setFrozenRows(1);
  }

  sh = ss.getSheetByName(SH.CONF);
  if (!sh) { sh = ss.insertSheet(SH.CONF); }
  if (sh.getLastRow() === 0) {
    const folder = DriveApp.createFolder('體表會成績回報_現場照片');
    writeRows_(sh, 1, [
      ['設定項目', '值'],
      ['系統名稱', '體表會成績回報系統'],
      ['學校名稱', CONF_DEFAULT['學校名稱']],
      ['學年度', CONF_DEFAULT['學年度']],
      ['活動名稱', CONF_DEFAULT['活動名稱']],
      ['活動日期', TW_NOW_.toISOString().slice(0, 10)],
      ['字號前綴', CONF_DEFAULT['字號前綴']],
      ['獎狀字號', ''],
      ['編號方式', '同一字號'],
      ['逾時分鐘', '15'],
      ['照片資料夾ID', folder.getId()]
    ]);
    sh.setFrozenRows(1);
  }
  return '初始化完成';
}

/* ───────── 小工具 ───────── */
function pin_() { return String(Math.floor(100000 + Math.random() * 900000)); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function safe_(v) { // 一律存成文字，並避免被當成公式
  const s = String(v === null || v === undefined ? '' : v);
  return /^[=+\-@]/.test(s) ? `'${s}` : s;
}
function writeRows_(sh, row, rows) {
  const out = rows.map(r => r.map(safe_));
  sh.getRange(row, 1, out.length, out[0].length).setNumberFormat('@').setValues(out);
}
function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) { throw new Error(`找不到工作表「${name}」，請先執行 setup()`); }
  return sh;
}
function table_(name) { // 回傳不含標題列的文字資料
  const sh = sheet_(name); const n = sh.getLastRow();
  if (n < 2) { return []; }
  return sh.getRange(2, 1, n - 1, sh.getLastColumn()).getDisplayValues();
}
/* 預設值一律用中性文字，各校請到網站的「系統設定」改成自己的學校名稱、活動名稱與獎狀字號 */
const TW_NOW_ = new Date(Date.now() + 8 * 3600000); // 臺灣時間（以 UTC 欄位讀取）
const SCHOOL_YEAR_ = String(TW_NOW_.getUTCFullYear() - 1911 - (TW_NOW_.getUTCMonth() < 7 ? 1 : 0)); // 8 月起算新學年度
const CONF_DEFAULT = { '系統名稱': '體表會成績回報系統', '學校名稱': '○○國小', '學年度': SCHOOL_YEAR_, '活動名稱': '體育表演會',
  '活動日期': '', '字號前綴': '○縣市○獎字第', '獎狀字號': '', '編號方式': '同一字號', '逾時分鐘': '15',
  '獎狀上邊界微調': '0', '獎狀下邊界微調': '0', '獎狀左邊界微調': '0', '獎狀右邊界微調': '0',
  '列印頁垂直位移': '0', '列印頁水平位移': '0', '頒獎日期': '', '頒獎日期格式': '中華民國', '檢錄員可更正': '開啟', '檢錄通行碼': '' };
const DATE_MODES = ['中華民國', '西元'];
const SHIFT_KEYS = ['列印頁垂直位移', '列印頁水平位移'];
const MARGIN_KEYS = ['獎狀上邊界微調', '獎狀下邊界微調', '獎狀左邊界微調', '獎狀右邊界微調'];
const NO_MODES = ['同一字號', '流水號'];
let CONF_CACHE_ = null; // 同一次請求內只讀一次設定工作表（每次請求都會重新載入程式，不會跨請求殘留）
function conf_() { // 舊版試算表沒有的設定項目，會以預設值補上
  if (!CONF_CACHE_) {
    const o = {}; Object.keys(CONF_DEFAULT).forEach(k => { o[k] = CONF_DEFAULT[k]; });
    table_(SH.CONF).forEach(r => { if (r[0]) { o[r[0]] = r[1]; } }); CONF_CACHE_ = o;
  }
  const copy = {}; Object.keys(CONF_CACHE_).forEach(k => { copy[k] = CONF_CACHE_[k]; }); return copy;
}
function str_(v, max, label, required) {
  const s = String(v === null || v === undefined ? '' : v).trim();
  if (required && !s) { throw new Error(`請填寫${label}`); }
  if (s.length > max) { throw new Error(`${label}過長`); }
  return s;
}
function parseTime_(s) { return new Date(`${String(s).replace(' ', 'T')}+08:00`).getTime(); }

/* ───────── 身分 ───────── */
function who_(pin) {
  let email = '';
  try { email = String(Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) { email = ''; }
  const p = String(pin || '').trim();
  const roles = []; let name = '';
  table_(SH.ROLE).forEach(r => {
    const role = String(r[0]).trim();
    if (!ROLES.includes(role)) { return; }
    const em = String(r[2]).trim().toLowerCase(); const pn = String(r[3]).trim();
    if ((email && em && em === email) || (p && pn && pn === p)) {
      if (!roles.includes(role)) { roles.push(role); }
      if (!name) { name = String(r[1]).trim(); }
    }
  });
  return { email: email, roles: roles, name: name };
}
function need_(pin, allowed) {
  const w = who_(pin);
  const ok = w.roles.some(r => allowed.includes(r));
  if (!ok) { throw new Error(`沒有權限：此功能限${allowed.join('、')}使用`); }
  return w;
}
/* 檢錄通行碼：系統設定有填才檢查。行政、司儀、校長登入後不必另外輸入 */
function gateCode_() { return String(conf_()['檢錄通行碼'] || '').trim(); }
function gate_(p) {
  const code = gateCode_(); if (!code) { return; }
  if (String(p.code === null || p.code === undefined ? '' : p.code).trim() === code) { return; }
  if (p.pin && who_(p.pin).roles.length) { return; }
  throw new Error('檢錄通行碼不正確，請向行政人員確認');
}

/* ───────── 項目與名單 ───────── */
function events_() {
  const ind = [], relay = [], idx = {};
  table_(SH.ROSTER).forEach(r => {
    const type = String(r[0]).trim(), ev = String(r[1]).trim(), cls = String(r[2]).trim();
    const no = String(r[3]).trim(), nm = String(r[4]).trim();
    if (!ev || !cls) { return; }
    const key = `${type}|${ev}`;
    if (!idx[key]) {
      idx[key] = { name: ev, grade: ev.charAt(0), finalists: [], classes: [], roster: {} };
      if (type === '個人') { ind.push(idx[key]); } else if (type === '接力') { relay.push(idx[key]); } else { delete idx[key]; return; }
    }
    const e = idx[key];
    if (type === '個人') {
      if (nm) { e.finalists.push({ cls: cls, name: nm }); }
    } else {
      if (!e.classes.includes(cls)) { e.classes.push(cls); e.roster[cls] = []; }
      if (nm) { e.roster[cls].push({ no: no, name: nm }); }
    }
  });
  return { individual: ind, relay: relay };
}

/* ───────── 給前端的初始資料 ───────── */
function getBootstrap(pin) {
  const ev = events_(); const c = conf_(); const w = who_(pin);
  return {
    version: VERSION, needCode: !!gateCode_(), allowRevise: c['檢錄員可更正'] !== '關閉', title: c['活動名稱'] || '體育表演會', appName: c['系統名稱'], school: c['學校名稱'], year: c['學年度'], eventDate: c['活動日期'],
    overdueMin: Number(c['逾時分鐘']) || 15,
    individual: ev.individual.map(e => ({
      name: e.name,
      grade: e.grade,
      finalists: e.finalists
    })),
    relay: ev.relay.map(e => {
      const counts = {}; e.classes.forEach(c => { counts[c] = e.roster[c].length; });
      return { name: e.name, grade: e.grade, classes: e.classes, counts: counts };
    }),
    who: w,
    assignees: table_(SH.ROLE).filter(r => String(r[0]).trim() === '行政')
      .map(r => String(r[1]).trim()).filter(String)
  };
}

/* ───────── 案件讀取 ───────── */
function rowToCase_(r, overdueMin, nowMs) {
  let places = [], log = [];
  try { places = JSON.parse(r[COL['成績資料']] || '[]'); } catch (e) { places = []; }
  try { log = JSON.parse(r[COL['處理紀錄']] || '[]'); } catch (e) { log = []; }
  const status = r[COL['狀態']];
  const waited = Math.floor((nowMs - parseTime_(r[COL['送出時間']])) / 60000);
  return {
    id: r[COL['案件編號']], at: r[COL['送出時間']], type: r[COL['類別']], event: r[COL['項目']],
    date: r[COL['比賽日期']], time: r[COL['比賽時間']], recorder: r[COL['檢錄員']],
    urgent: r[COL['急件']] === '是', places: places, status: status,
    assignee: r[COL['負責人']], result: r[COL['處理結果']], doneAt: r[COL['完成時間']],
    announcedAt: r[COL['播報時間']], audit: r[COL['稽核']], auditNote: r[COL['稽核意見']],
    auditAt: r[COL['稽核時間']], hasPhoto: !!r[COL['照片檔案ID']], log: log, printedAt: r[COL['列印時間']] || '',
    corrected: log.some(l => String(l.what).startsWith('更正成績')),
    waited: waited, overdue: status !== '已完成' && waited >= overdueMin
  };
}
function listCases(pin) {
  const c = conf_(); const om = Number(c['逾時分鐘']) || 15; const nowMs = Date.now();
  const cases = table_(SH.CASE).filter(r => r[0])
    .map(r => rowToCase_(r, om, nowMs));
  return { cases: cases, who: who_(pin), now: now_(), needCode: !!gateCode_(), allowRevise: c['檢錄員可更正'] !== '關閉', mcOrder: String(c['播報順序'] || '').split(',').filter(String) };
}
function findRow_(id) {
  const sh = sheet_(SH.CASE); const n = sh.getLastRow();
  if (n >= 2) {
    const ids = sh.getRange(2, 1, n - 1, 1).getDisplayValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === String(id)) {
        return { sh: sh, row: i + 2, values: sh.getRange(i + 2, 1, 1, CASE_H.length).getDisplayValues()[0] };
      }
    }
  }
  throw new Error(`找不到案件 ${id}`);
}
function saveRow_(f, who, action) {
  let log = [];
  try { log = JSON.parse(f.values[COL['處理紀錄']] || '[]'); } catch (e) { log = []; }
  log.push({ t: now_(), by: who, what: action });
  f.values[COL['處理紀錄']] = JSON.stringify(log);
  writeRows_(f.sh, f.row, [f.values]);
}
function getPhoto(pin, id) { // 現場照片可能拍到學生，限登入後查看
  need_(pin, ROLES);
  const f = findRow_(id); const fid = f.values[COL['照片檔案ID']];
  if (!fid) { return ''; }
  const blob = DriveApp.getFileById(fid).getBlob();
  return `data:${blob.getContentType()};base64,${Utilities.base64Encode(blob.getBytes())}`;
}

/* ───────── 名次資料檢查（送出與更正共用） ───────── */
function places_(type, e, input) {
  if (!Array.isArray(input) || input.length < 1 || input.length > 3) { throw new Error('請填寫名次'); }
  const seen = {}; const places = [];
  input.forEach(x => {
    x = x || {};
    const rank = Number(x.rank);
    if (![1, 2, 3].includes(rank) || seen[`r${rank}`]) { throw new Error('名次資料不正確'); }
    seen[`r${rank}`] = true;
    const cls = str_(x.cls, 3, `第${rank}名班級`, true);
    if (!/^[1-6]\d{2}$/.test(cls) || cls.charAt(0) !== GRADE_NUM[e.grade]) { throw new Error(`第${rank}名班級與項目年級不符`); }
    const noScore = type === '接力' && e.grade === '三'; // 三年級大隊接力為趣味賽，可不記成績
    const score = str_(x.score, 16, `第${rank}名成績`, !noScore);
    const o = { rank: rank, cls: cls, score: score };
    if (type === '個人') {
      o.name = str_(x.name, 12, `第${rank}名姓名`, true);
      if (seen[cls + o.name]) { throw new Error(`同一位選手重複出現：${o.name}`); }
      seen[cls + o.name] = true;
    } else {
      if (e.classes && !e.classes.includes(cls)) { throw new Error(`班級不在該項目名單：${cls}`); }
      if (seen[cls]) { throw new Error(`同一班級重複出現：${cls}`); }
      seen[cls] = true;
    }
    places.push(o);
  });
  places.sort((a, b) => a.rank - b.rank);
  if (places[0].rank !== 1) { throw new Error('請至少填寫第 1 名'); }
  return places;
}
function summary_(places) {
  return places.map(x => `第${x.rank}名 ${x.cls}${x.name ? ` ${x.name}` : ''} ${x.score || '不記成績'}`).join('／');
}

/* 案件編號取「目前最大號碼」與「設定裡記錄的流水號」較大者加 1，刪除案件後號碼不會重複使用 */
function setConf_(key, value) {
  const sh = sheet_(SH.CONF); const rows = table_(SH.CONF);
  if (CONF_CACHE_) { CONF_CACHE_[key] = value; }
  for (let i = 0; i < rows.length; i++) { if (rows[i][0] === key) { writeRows_(sh, i + 2, [[key, value]]); return; } }
  writeRows_(sh, sh.getLastRow() + 1, [[key, value]]);
}
function nextCaseId_(rows) {
  let max = Number(conf_()['案件流水號']) || 0;
  rows.forEach(r => { const m = /^R(\d+)$/.exec(r[0] || ''); if (m) { max = Math.max(max, Number(m[1])); } });
  const n = max + 1; setConf_('案件流水號', String(n));
  return `R${n < 1000 ? (`000${n}`).slice(-3) : String(n)}`;
}

/* ───────── 檢錄員送出 ───────── */
function submitReport(p, force) {
  p = p || {};
  gate_(p);
  const type = p.type === '接力' ? '接力' : '個人';
  const recorder = str_(p.recorder, 20, '檢錄員姓名', true);
  const date = str_(p.date, 10, '日期', true); const time = str_(p.time, 5, '時間', true);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !/^\d{2}:\d{2}$/.test(time)) { throw new Error('日期或時間格式不正確'); }
  const ev = events_(); const list = type === '接力' ? ev.relay : ev.individual;
  const evName = str_(p.event, 30, '項目', true);
  const e = list.filter(x => x.name === evName)[0];
  if (!e) { throw new Error(`項目不存在：${evName}`); }
  const places = places_(type, e, p.places);
  let photo = null;
  if (p.photo) {
    const m = /^data:(image\/(?:jpeg|png));base64,(.+)$/.exec(String(p.photo));
    if (!m || m[2].length > 8000000) { throw new Error('照片格式不正確或檔案過大'); }
    photo = { mime: m[1], b64: m[2] };
  }

  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = sheet_(SH.CASE); const rows = table_(SH.CASE);
    if (!force) {
      const dup = rows.filter(r => r[COL['項目']] === evName)[0];
      if (dup) { return { duplicate: dup[COL['案件編號']] }; }
    }
    const id = nextCaseId_(rows);
    let fileId = '';
    if (photo) {
      const blob = Utilities.newBlob(Utilities.base64Decode(photo.b64), photo.mime, `${id}_${evName}.jpg`);
      fileId = DriveApp.getFolderById(conf_()['照片資料夾ID']).createFile(blob).getId();
    }
    const summary = summary_(places);
    const w = who_('');
    const row = CASE_H.map(() => '');
    row[COL['案件編號']] = id; row[COL['送出時間']] = now_(); row[COL['類別']] = type; row[COL['項目']] = evName;
    row[COL['比賽日期']] = date; row[COL['比賽時間']] = time; row[COL['檢錄員']] = recorder;
    row[COL['填報帳號']] = w.email; row[COL['急件']] = p.urgent ? '是' : '否'; row[COL['成績摘要']] = summary;
    row[COL['狀態']] = '待處理'; row[COL['照片檔案ID']] = fileId; row[COL['成績資料']] = JSON.stringify(places);
    row[COL['處理紀錄']] = JSON.stringify([{ t: now_(), by: recorder, what: `送出成績${p.urgent ? '（急件）' : ''}` }]);
    writeRows_(sh, rows.length + 2, [row]);
    return { id: id, summary: summary, at: row[COL['送出時間']] };
  } finally { lock.releaseLock(); }
}

/* ───────── 檢錄員更正自己回報過的項目（填報頁帶入原資料後送出） ───────── */
function reviseReport(p) {
  p = p || {};
  gate_(p);
  const id = str_(p.id, 8, '案件編號', true);
  const recorder = str_(p.recorder, 20, '檢錄員姓名', true);
  const date = str_(p.date, 10, '日期', true); const time = str_(p.time, 5, '時間', true);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !/^\d{2}:\d{2}$/.test(time)) { throw new Error('日期或時間格式不正確'); }
  const why = str_(p.reason, 100, '更正原因', false);
  if (conf_()['檢錄員可更正'] === '關閉') { throw new Error('目前不開放檢錄員自行更正，請洽獎狀列印（行政）人員'); }
  let photo = null;
  if (p.photo) {
    const m = /^data:(image\/(?:jpeg|png));base64,(.+)$/.exec(String(p.photo));
    if (!m || m[2].length > 8000000) { throw new Error('照片格式不正確或檔案過大'); }
    photo = { mime: m[1], b64: m[2] };
  }
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const f = findRow_(id); const v = f.values;
    const type = v[COL['類別']]; const evName = v[COL['項目']];
    if (String(p.event || '') !== evName) { throw new Error('項目與原案件不符，請重新選擇項目'); }
    const ev = events_(); const list = type === '接力' ? ev.relay : ev.individual;
    const e = list.filter(x => x.name === evName)[0] || { name: evName, grade: evName.charAt(0), classes: null };
    const places = places_(type, e, p.places);
    const acts = []; const notes = [];
    const before = v[COL['成績摘要']]; const after = summary_(places);
    if (before !== after) {
      acts.push(`更正成績（檢錄員）：${before} → ${after}`);
      v[COL['成績資料']] = JSON.stringify(places); v[COL['成績摘要']] = after;
      if (v[COL['播報時間']]) { v[COL['播報時間']] = ''; notes.push('需重新播報'); }
      if (v[COL['狀態']] === '已完成') { v[COL['狀態']] = '處理中'; v[COL['完成時間']] = ''; notes.push('退回處理中，獎狀請重印'); }
      if (v[COL['稽核']] === '已核可') { v[COL['稽核']] = ''; v[COL['稽核意見']] = ''; v[COL['稽核時間']] = ''; notes.push('需重新稽核'); }
      if (v[COL['列印時間']]) { v[COL['列印時間']] = ''; notes.push('需重新列印'); }
    }
    if (date !== v[COL['比賽日期']] || time !== v[COL['比賽時間']]) {
      acts.push(`比賽時間：${v[COL['比賽日期']]} ${v[COL['比賽時間']]} → ${date} ${time}`);
      v[COL['比賽日期']] = date; v[COL['比賽時間']] = time;
    }
    const urgent = p.urgent ? '是' : '否';
    if (urgent !== v[COL['急件']]) { acts.push(urgent === '是' ? '標記為急件' : '取消急件'); v[COL['急件']] = urgent; }
    if (photo) {
      const blob = Utilities.newBlob(Utilities.base64Decode(photo.b64), photo.mime, `${id}_${evName}_更正.jpg`);
      v[COL['照片檔案ID']] = DriveApp.getFolderById(conf_()['照片資料夾ID']).createFile(blob).getId();
      acts.push('更換現場照片');
    }
    if (!acts.length) { return { id: id, changed: false }; }
    saveRow_(f, recorder, acts.join('；') + (why ? `（原因：${why}）` : '') + (notes.length ? `；${notes.join('、')}` : ''));
    return { id: id, changed: true, notes: notes, at: now_(), summary: after };
  } finally { lock.releaseLock(); }
}

/* ───────── 列印作業（原行政收件匣） ───────── */
function adminUpdate(pin, id, patch) {
  const w = need_(pin, ['行政']); patch = patch || {};
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const f = findRow_(id); const v = f.values; const acts = [];
    const assignee = str_(patch.assignee, 20, '負責人', false);
    const status = str_(patch.status, 5, '狀態', true);
    const result = str_(patch.result, 300, '處理結果', false);
    if (!STATUS.includes(status)) { throw new Error('狀態不正確'); }
    if (status === '已完成' && !result) { throw new Error('標記已完成前，請填寫實際處理結果'); }
    if (assignee !== v[COL['負責人']]) { acts.push(`分派負責人：${assignee || '（清空）'}`); v[COL['負責人']] = assignee; }
    if (status !== v[COL['狀態']]) {
      acts.push(`狀態：${v[COL['狀態']]} → ${status}`);
      v[COL['狀態']] = status;
      v[COL['完成時間']] = status === '已完成' ? now_() : '';
      if (status === '已完成') { v[COL['稽核']] = ''; v[COL['稽核意見']] = ''; v[COL['稽核時間']] = ''; }
    }
    if (result !== v[COL['處理結果']]) { acts.push(`處理結果：${result || '（清空）'}`); v[COL['處理結果']] = result; }
    if (acts.length === 0) { return { ok: true, changed: false }; }
    saveRow_(f, w.name || '行政', acts.join('；'));
    return { ok: true, changed: true };
  } finally { lock.releaseLock(); }
}
/* 更正已送出的成績（限行政，必填原因，前後內容留在處理紀錄） */
function correctReport(pin, id, input, reason) {
  const w = need_(pin, ['行政']);
  const why = str_(reason, 100, '更正原因', true);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const f = findRow_(id); const v = f.values;
    const type = v[COL['類別']]; const evName = v[COL['項目']];
    const ev = events_(); const list = type === '接力' ? ev.relay : ev.individual;
    const e = list.filter(x => x.name === evName)[0] || { name: evName, grade: evName.charAt(0), classes: null };
    const places = places_(type, e, input);
    const before = v[COL['成績摘要']]; const after = summary_(places);
    if (before === after) { return { ok: true, changed: false }; }
    v[COL['成績資料']] = JSON.stringify(places); v[COL['成績摘要']] = after;
    const notes = [];
    if (v[COL['播報時間']]) { v[COL['播報時間']] = ''; notes.push('需重新播報'); }
    if (v[COL['狀態']] === '已完成') { v[COL['狀態']] = '處理中'; v[COL['完成時間']] = ''; notes.push('退回處理中，獎狀請重印'); }
    if (v[COL['稽核']] === '已核可') { v[COL['稽核']] = ''; v[COL['稽核意見']] = ''; v[COL['稽核時間']] = ''; notes.push('需重新稽核'); }
    if (v[COL['列印時間']]) { v[COL['列印時間']] = ''; notes.push('需重新列印'); }
    saveRow_(f, w.name || '行政', `更正成績：${before} → ${after}（原因：${why}）${notes.length ? `；${notes.join('、')}` : ''}`);
    return { ok: true, changed: true, notes: notes };
  } finally { lock.releaseLock(); }
}
/* 司儀自訂播報順序：把待播報的案件上移、下移或置頂。順序存在設定工作表，所有司儀看到的一致 */
function moveAnnounce(pin, id, dir) {
  need_(pin, ['司儀', '行政']);
  if (!['up', 'down', 'top'].includes(dir)) { throw new Error('移動方向不正確'); }
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const saved = String(conf_()['播報順序'] || '').split(',').filter(String);
    const pos = x => { const i = saved.indexOf(x); return i < 0 ? 1e9 : i; };
    const ids = caseRows_().filter(r => r[0] && !r[COL['播報時間']])
      .map(r => ({
      id: r[0],
      at: r[COL['送出時間']]
    }))
      .sort((a, b) => pos(a.id) - pos(b.id) || (a.at < b.at ? -1 : a.at > b.at ? 1 : (a.id < b.id ? -1 : 1)))
      .map(x => x.id);
    const i = ids.indexOf(String(id));
    if (i < 0) { throw new Error(`找不到待播報的案件 ${id}`); }
    const j = dir === 'top' ? 0 : dir === 'up' ? i - 1 : i + 1;
    if (j >= 0 && j < ids.length && j !== i) { ids.splice(i, 1); ids.splice(j, 0, String(id)); }
    setConf_('播報順序', ids.join(','));
    return { ok: true, order: ids };
  } finally { lock.releaseLock(); }
}

const AUTO_DONE = '司儀已播報（自動完成）';
/* 司儀播報後，案件自動變成已完成；取消播報時，若是自動完成的就一併退回 */
function markAnnounced(pin, id, flag) {
  const w = need_(pin, ['司儀', '行政']);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const f = findRow_(id); const v = f.values; let act;
    if (flag) {
      v[COL['播報時間']] = now_(); act = '已播報';
      if (v[COL['狀態']] !== '已完成') {
        v[COL['狀態']] = '已完成'; v[COL['完成時間']] = now_();
        if (!v[COL['處理結果']]) { v[COL['處理結果']] = AUTO_DONE; }
        v[COL['稽核']] = ''; v[COL['稽核意見']] = ''; v[COL['稽核時間']] = '';
        act += '，案件自動完成';
      }
    } else {
      v[COL['播報時間']] = ''; act = '取消已播報';
      if (v[COL['狀態']] === '已完成' && v[COL['處理結果']] === AUTO_DONE) {
        v[COL['狀態']] = v[COL['列印時間']] ? '處理中' : '待處理'; v[COL['完成時間']] = ''; v[COL['處理結果']] = '';
        v[COL['稽核']] = ''; v[COL['稽核意見']] = ''; v[COL['稽核時間']] = '';
        act += `，案件退回${v[COL['狀態']]}`;
      }
    }
    saveRow_(f, w.name || '司儀', act);
    return { ok: true, status: v[COL['狀態']] };
  } finally { lock.releaseLock(); }
}

/* 標記獎狀已列印（限行政）。ids 可為單一編號或陣列；狀態真的改變才寫入紀錄 */
function fixCaseHeader_() { // 舊版建立的試算表沒有後來新增的欄位標題，補上
  const sh = sheet_(SH.CASE); const h = sh.getRange(1, 1, 1, CASE_H.length).getDisplayValues()[0];
  if (CASE_H.some((x, i) => h[i] !== x)) { writeRows_(sh, 1, [CASE_H]); }
}
function markPrinted(pin, ids, flag) {
  const w = need_(pin, ['行政']);
  const want = (Array.isArray(ids) ? ids : [ids]).map(String).filter(String);
  if (!want.length) { return { ok: true, changed: 0 }; }
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    fixCaseHeader_();
    const sh = sheet_(SH.CASE); const rows = caseRows_(); const t = now_(); let n = 0;
    rows.forEach((v, i) => {
      if (!v[0] || !want.includes(v[0]) || !!v[COL['列印時間']] === !!flag) { return; }
      v[COL['列印時間']] = flag ? t : '';
      // 行政收件匣的狀態跟著流程走：列印後由待處理進到處理中，取消列印且尚未播報則退回待處理
      if (flag && v[COL['狀態']] === '待處理') { v[COL['狀態']] = '處理中'; }
      else if (!flag && v[COL['狀態']] === '處理中' && !v[COL['播報時間']]) { v[COL['狀態']] = '待處理'; }
      saveRow_({ sh: sh, row: i + 2, values: v }, w.name || '行政', flag ? '標記已列印' : '取消已列印'); n++;
    });
    return { ok: true, changed: n };
  } finally { lock.releaseLock(); }
}

/* ───────── 校長稽核（v1.18.0 起介面已取消核可功能，分頁改為「頒發獎狀」；此函式僅為相容舊資料而保留，介面不會呼叫） ───────── */
function auditDecide(pin, id, decision, note) {
  const w = need_(pin, ['校長']);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const f = findRow_(id); const v = f.values;
    if (v[COL['狀態']] !== '已完成') { throw new Error('案件尚未完成，無法稽核'); }
    const n = str_(note, 300, '稽核意見', decision === '退回');
    if (decision === '核可') {
      v[COL['稽核']] = '已核可';
    } else if (decision === '退回') {
      v[COL['稽核']] = '退回補正'; v[COL['狀態']] = '處理中'; v[COL['完成時間']] = '';
    } else { throw new Error('稽核決定不正確'); }
    v[COL['稽核意見']] = n; v[COL['稽核時間']] = now_();
    saveRow_(f, w.name || '校長', `稽核${decision}${n ? `：${n}` : ''}`);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* 流水號：字號結尾是數字就從該數字往上加並維持位數（115001 → 115002）；否則在後面接三位序號 */
function seqNo_(base, i) {
  const m = /^(.*?)(\d+)$/.exec(String(base || ''));
  if (!m) { return String(base || '') + (`000${i + 1}`).slice(-3); }
  let s = String(Number(m[2]) + i);
  while (s.length < m[2].length) { s = `0${s}`; }
  return m[1] + s;
}

/* ───────── 匯出獎狀合併列印資料 ───────── */
/* 所有獎狀逐張列出（Excel 匯出與 Word 獎狀共用，編號因此一致）
   kind：個人＝個人單項；選手＝接力每位選手；班級＝接力以班級為對象。
   接力有選手名單時，每個名次先列一張班級獎狀，再列各選手；沒有名單時只有班級獎狀。 */
function certRows_() {
  const cf = conf_(); const base = cf['獎狀字號'] || ''; const serial = cf['編號方式'] === '流水號';
  const ev = events_(); const rosterOf = {};
  ev.relay.forEach(e => { rosterOf[e.name] = e.roster; });
  const rows = [];
  table_(SH.CASE).forEach(r => {
    if (!r[0]) { return; }
    let places = []; try { places = JSON.parse(r[COL['成績資料']] || '[]'); } catch (e) { places = []; }
    const event = r[COL['項目']]; const grade = event.charAt(0);
    places.forEach(x => {
      const add = (kind, name, inSheet) => {
        rows.push({ kind: kind, inSheet: inSheet, caseId: r[0], status: r[COL['狀態']], event: event, rank: String(x.rank),
          grade: grade, clsNum: String(Number(String(x.cls).slice(1))), name: name, score: x.score || '' });
      };
      if (r[COL['類別']] === '接力') {
        const mem = (rosterOf[event] || {})[x.cls] || []; const seen = {};
        const uniq = mem.filter(m => { const k = m.no + m.name; if (seen[k]) { return false; } seen[k] = true; return true; });
        add('班級', '', uniq.length === 0);
        uniq.forEach(m => { add('選手', m.name, true); });
      } else { add('個人', x.name, true); }
    });
  });
  rows.forEach((o, i) => {
    o.no = serial ? seqNo_(base, i) : base;
    o.fullNo = o.no ? `${(cf['字號前綴'] || '') + o.no}號` : '';
  });
  return rows;
}
/* 獎狀最下方的頒獎日期：先用設定的「頒獎日期」，沒填就用「活動日期」，兩者都沒填則用產生獎狀當天，確保不會空白 */
function awardDate_(cf) {
  const pick = s => /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || '').trim());
  const md = pick(cf['頒獎日期']) || pick(cf['活動日期']) || pick(String(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd')).slice(0, 10));
  const y = Number(md[1]), tail = `${Number(md[2])}月${Number(md[3])}日`;
  return cf['頒獎日期格式'] === '西元' ? `西元${y}年${tail}` : `中華民國${y - 1911}年${tail}`;
}
function exportMerge(pin) {
  need_(pin, ['行政']);
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const cf = conf_(); const awardDate = awardDate_(cf);
  const out = [['編號', '項目', '名次', '年', '班', '姓名', '成績', '案件編號', '案件狀態', '活動名稱', '頒發日期', '完整字號']];
  certRows_().forEach(o => {
    if (o.inSheet) { out.push([o.no, o.event, o.rank, o.grade, o.clsNum, o.name, o.score, o.caseId, o.status, cf['活動名稱'], awardDate, o.fullNo]); }
  });
  let sh = ss.getSheetByName(SH.MERGE);
  if (!sh) { sh = ss.insertSheet(SH.MERGE); }
  sh.clearContents();
  writeRows_(sh, 1, out);
  return {
    count: out.length - 1,
    url: `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=xlsx&gid=${sh.getSheetId()}`,
    view: `${ss.getUrl()}#gid=${sh.getSheetId()}`
  };
}

/* ───────── Word 獎狀（限行政）：後端提供版型與資料，由瀏覽器組成 .docx ───────── */
const CERT_TEMPLATE = {"parts":[["[Content_Types].xml","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/><Override PartName=\"/customXml/itemProps1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.customXmlProperties+xml\"/><Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/><Override PartName=\"/word/settings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml\"/><Override PartName=\"/word/webSettings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml\"/><Override PartName=\"/word/footnotes.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml\"/><Override PartName=\"/word/endnotes.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml\"/><Override PartName=\"/word/fontTable.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml\"/><Override PartName=\"/word/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/><Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/><Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/></Types>"],["_rels/.rels","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>"],["word/_rels/document.xml.rels","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId8\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\"/><Relationship Id=\"rId7\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable\" Target=\"fontTable.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml\" Target=\"../customXml/item1.xml\"/><Relationship Id=\"rId6\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes\" Target=\"endnotes.xml\"/><Relationship Id=\"rId5\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes\" Target=\"footnotes.xml\"/><Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings\" Target=\"webSettings.xml\"/></Relationships>"],["word/styles.xml","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:styles xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" mc:Ignorable=\"w14 w15 w16se w16cid w16 w16cex w16sdtdh\"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme=\"minorHAnsi\" w:eastAsiaTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorHAnsi\" w:cstheme=\"minorBidi\"/><w:kern w:val=\"2\"/><w:sz w:val=\"24\"/><w:szCs w:val=\"22\"/><w:lang w:val=\"en-US\" w:eastAsia=\"zh-TW\" w:bidi=\"ar-SA\"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults><w:latentStyles w:defLockedState=\"0\" w:defUIPriority=\"99\" w:defSemiHidden=\"0\" w:defUnhideWhenUsed=\"0\" w:defQFormat=\"0\" w:count=\"376\"><w:lsdException w:name=\"Normal\" w:uiPriority=\"0\" w:qFormat=\"1\"/><w:lsdException w:name=\"heading 1\" w:uiPriority=\"9\" w:qFormat=\"1\"/><w:lsdException w:name=\"heading 2\" w:semiHidden=\"1\" w:uiPriority=\"9\" w:unhideWhenUsed=\"1\" w:qFormat=\"1\"/><w:lsdException w:name=\"heading 3\" w:semiHidden=\"1\" w:uiPriority=\"9\" w:unhideWhenUsed=\"1\" w:qFormat=\"1\"/><w:lsdException w:name=\"heading 4\" w:semiHidden=\"1\" w:uiPriority=\"9\" w:unhideWhenUsed=\"1\" w:qFormat=\"1\"/><w:lsdException w:name=\"heading 5\" w:semiHidden=\"1\" w:uiPriority=\"9\" w:unhideWhenUsed=\"1\" w:qFormat=\"1\"/><w:lsdException w:name=\"heading 6\" w:semiHidden=\"1\" w:uiPriority=\"9\" w:unhideWhenUsed=\"1\" w:qFormat=\"1\"/><w:lsdException w:name=\"heading 7\" w:semiHidden=\"1\" w:uiPriority=\"9\" w:unhideWhenUsed=\"1\" w:qFormat=\"1\"/><w:lsdException w:name=\"heading 8\" w:semiHidden=\"1\" w:uiPriority=\"9\" w:unhideWhenUsed=\"1\" w:qFormat=\"1\"/><w:lsdException w:name=\"heading 9\" w:semiHidden=\"1\" w:uiPriority=\"9\" w:unhideWhenUsed=\"1\" w:qFormat=\"1\"/><w:lsdException w:name=\"index 1\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"index 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"index 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"index 4\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"index 5\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"index 6\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"index 7\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"index 8\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"index 9\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"toc 1\" w:semiHidden=\"1\" w:uiPriority=\"39\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"toc 2\" w:semiHidden=\"1\" w:uiPriority=\"39\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"toc 3\" w:semiHidden=\"1\" w:uiPriority=\"39\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"toc 4\" w:semiHidden=\"1\" w:uiPriority=\"39\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"toc 5\" w:semiHidden=\"1\" w:uiPriority=\"39\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"toc 6\" w:semiHidden=\"1\" w:uiPriority=\"39\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"toc 7\" w:semiHidden=\"1\" w:uiPriority=\"39\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"toc 8\" w:semiHidden=\"1\" w:uiPriority=\"39\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"toc 9\" w:semiHidden=\"1\" w:uiPriority=\"39\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Normal Indent\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"footnote text\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"annotation text\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"header\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"footer\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"index heading\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"caption\" w:semiHidden=\"1\" w:uiPriority=\"35\" w:unhideWhenUsed=\"1\" w:qFormat=\"1\"/><w:lsdException w:name=\"table of figures\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"envelope address\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"envelope return\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"footnote reference\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"annotation reference\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"line number\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"page number\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"endnote reference\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"endnote text\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"table of authorities\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"macro\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"toa heading\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Bullet\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Number\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List 4\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List 5\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Bullet 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Bullet 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Bullet 4\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Bullet 5\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Number 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Number 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Number 4\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Number 5\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Title\" w:uiPriority=\"10\" w:qFormat=\"1\"/><w:lsdException w:name=\"Closing\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Signature\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Default Paragraph Font\" w:semiHidden=\"1\" w:uiPriority=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Body Text\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Body Text Indent\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Continue\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Continue 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Continue 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Continue 4\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"List Continue 5\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Message Header\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Subtitle\" w:uiPriority=\"11\" w:qFormat=\"1\"/><w:lsdException w:name=\"Salutation\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Date\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Body Text First Indent\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Body Text First Indent 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Note Heading\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Body Text 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Body Text 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Body Text Indent 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Body Text Indent 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Block Text\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Hyperlink\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"FollowedHyperlink\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Strong\" w:uiPriority=\"22\" w:qFormat=\"1\"/><w:lsdException w:name=\"Emphasis\" w:uiPriority=\"20\" w:qFormat=\"1\"/><w:lsdException w:name=\"Document Map\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Plain Text\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"E-mail Signature\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"HTML Top of Form\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"HTML Bottom of Form\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Normal (Web)\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"HTML Acronym\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"HTML Address\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"HTML Cite\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"HTML Code\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"HTML Definition\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"HTML Keyboard\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"HTML Preformatted\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"HTML Sample\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"HTML Typewriter\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"HTML Variable\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Normal Table\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"annotation subject\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"No List\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Outline List 1\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Outline List 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Outline List 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Simple 1\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Simple 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Simple 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Classic 1\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Classic 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Classic 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Classic 4\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Colorful 1\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Colorful 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Colorful 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Columns 1\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Columns 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Columns 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Columns 4\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Columns 5\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Grid 1\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Grid 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Grid 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Grid 4\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Grid 5\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Grid 6\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Grid 7\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Grid 8\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table List 1\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table List 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table List 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table List 4\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table List 5\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table List 6\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table List 7\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table List 8\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table 3D effects 1\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table 3D effects 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table 3D effects 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Contemporary\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Elegant\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Professional\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Subtle 1\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Subtle 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Web 1\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Web 2\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Web 3\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Balloon Text\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Table Grid\" w:uiPriority=\"59\"/><w:lsdException w:name=\"Table Theme\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Placeholder Text\" w:semiHidden=\"1\"/><w:lsdException w:name=\"No Spacing\" w:uiPriority=\"1\" w:qFormat=\"1\"/><w:lsdException w:name=\"Light Shading\" w:uiPriority=\"60\"/><w:lsdException w:name=\"Light List\" w:uiPriority=\"61\"/><w:lsdException w:name=\"Light Grid\" w:uiPriority=\"62\"/><w:lsdException w:name=\"Medium Shading 1\" w:uiPriority=\"63\"/><w:lsdException w:name=\"Medium Shading 2\" w:uiPriority=\"64\"/><w:lsdException w:name=\"Medium List 1\" w:uiPriority=\"65\"/><w:lsdException w:name=\"Medium List 2\" w:uiPriority=\"66\"/><w:lsdException w:name=\"Medium Grid 1\" w:uiPriority=\"67\"/><w:lsdException w:name=\"Medium Grid 2\" w:uiPriority=\"68\"/><w:lsdException w:name=\"Medium Grid 3\" w:uiPriority=\"69\"/><w:lsdException w:name=\"Dark List\" w:uiPriority=\"70\"/><w:lsdException w:name=\"Colorful Shading\" w:uiPriority=\"71\"/><w:lsdException w:name=\"Colorful List\" w:uiPriority=\"72\"/><w:lsdException w:name=\"Colorful Grid\" w:uiPriority=\"73\"/><w:lsdException w:name=\"Light Shading Accent 1\" w:uiPriority=\"60\"/><w:lsdException w:name=\"Light List Accent 1\" w:uiPriority=\"61\"/><w:lsdException w:name=\"Light Grid Accent 1\" w:uiPriority=\"62\"/><w:lsdException w:name=\"Medium Shading 1 Accent 1\" w:uiPriority=\"63\"/><w:lsdException w:name=\"Medium Shading 2 Accent 1\" w:uiPriority=\"64\"/><w:lsdException w:name=\"Medium List 1 Accent 1\" w:uiPriority=\"65\"/><w:lsdException w:name=\"Revision\" w:semiHidden=\"1\"/><w:lsdException w:name=\"List Paragraph\" w:uiPriority=\"34\" w:qFormat=\"1\"/><w:lsdException w:name=\"Quote\" w:uiPriority=\"29\" w:qFormat=\"1\"/><w:lsdException w:name=\"Intense Quote\" w:uiPriority=\"30\" w:qFormat=\"1\"/><w:lsdException w:name=\"Medium List 2 Accent 1\" w:uiPriority=\"66\"/><w:lsdException w:name=\"Medium Grid 1 Accent 1\" w:uiPriority=\"67\"/><w:lsdException w:name=\"Medium Grid 2 Accent 1\" w:uiPriority=\"68\"/><w:lsdException w:name=\"Medium Grid 3 Accent 1\" w:uiPriority=\"69\"/><w:lsdException w:name=\"Dark List Accent 1\" w:uiPriority=\"70\"/><w:lsdException w:name=\"Colorful Shading Accent 1\" w:uiPriority=\"71\"/><w:lsdException w:name=\"Colorful List Accent 1\" w:uiPriority=\"72\"/><w:lsdException w:name=\"Colorful Grid Accent 1\" w:uiPriority=\"73\"/><w:lsdException w:name=\"Light Shading Accent 2\" w:uiPriority=\"60\"/><w:lsdException w:name=\"Light List Accent 2\" w:uiPriority=\"61\"/><w:lsdException w:name=\"Light Grid Accent 2\" w:uiPriority=\"62\"/><w:lsdException w:name=\"Medium Shading 1 Accent 2\" w:uiPriority=\"63\"/><w:lsdException w:name=\"Medium Shading 2 Accent 2\" w:uiPriority=\"64\"/><w:lsdException w:name=\"Medium List 1 Accent 2\" w:uiPriority=\"65\"/><w:lsdException w:name=\"Medium List 2 Accent 2\" w:uiPriority=\"66\"/><w:lsdException w:name=\"Medium Grid 1 Accent 2\" w:uiPriority=\"67\"/><w:lsdException w:name=\"Medium Grid 2 Accent 2\" w:uiPriority=\"68\"/><w:lsdException w:name=\"Medium Grid 3 Accent 2\" w:uiPriority=\"69\"/><w:lsdException w:name=\"Dark List Accent 2\" w:uiPriority=\"70\"/><w:lsdException w:name=\"Colorful Shading Accent 2\" w:uiPriority=\"71\"/><w:lsdException w:name=\"Colorful List Accent 2\" w:uiPriority=\"72\"/><w:lsdException w:name=\"Colorful Grid Accent 2\" w:uiPriority=\"73\"/><w:lsdException w:name=\"Light Shading Accent 3\" w:uiPriority=\"60\"/><w:lsdException w:name=\"Light List Accent 3\" w:uiPriority=\"61\"/><w:lsdException w:name=\"Light Grid Accent 3\" w:uiPriority=\"62\"/><w:lsdException w:name=\"Medium Shading 1 Accent 3\" w:uiPriority=\"63\"/><w:lsdException w:name=\"Medium Shading 2 Accent 3\" w:uiPriority=\"64\"/><w:lsdException w:name=\"Medium List 1 Accent 3\" w:uiPriority=\"65\"/><w:lsdException w:name=\"Medium List 2 Accent 3\" w:uiPriority=\"66\"/><w:lsdException w:name=\"Medium Grid 1 Accent 3\" w:uiPriority=\"67\"/><w:lsdException w:name=\"Medium Grid 2 Accent 3\" w:uiPriority=\"68\"/><w:lsdException w:name=\"Medium Grid 3 Accent 3\" w:uiPriority=\"69\"/><w:lsdException w:name=\"Dark List Accent 3\" w:uiPriority=\"70\"/><w:lsdException w:name=\"Colorful Shading Accent 3\" w:uiPriority=\"71\"/><w:lsdException w:name=\"Colorful List Accent 3\" w:uiPriority=\"72\"/><w:lsdException w:name=\"Colorful Grid Accent 3\" w:uiPriority=\"73\"/><w:lsdException w:name=\"Light Shading Accent 4\" w:uiPriority=\"60\"/><w:lsdException w:name=\"Light List Accent 4\" w:uiPriority=\"61\"/><w:lsdException w:name=\"Light Grid Accent 4\" w:uiPriority=\"62\"/><w:lsdException w:name=\"Medium Shading 1 Accent 4\" w:uiPriority=\"63\"/><w:lsdException w:name=\"Medium Shading 2 Accent 4\" w:uiPriority=\"64\"/><w:lsdException w:name=\"Medium List 1 Accent 4\" w:uiPriority=\"65\"/><w:lsdException w:name=\"Medium List 2 Accent 4\" w:uiPriority=\"66\"/><w:lsdException w:name=\"Medium Grid 1 Accent 4\" w:uiPriority=\"67\"/><w:lsdException w:name=\"Medium Grid 2 Accent 4\" w:uiPriority=\"68\"/><w:lsdException w:name=\"Medium Grid 3 Accent 4\" w:uiPriority=\"69\"/><w:lsdException w:name=\"Dark List Accent 4\" w:uiPriority=\"70\"/><w:lsdException w:name=\"Colorful Shading Accent 4\" w:uiPriority=\"71\"/><w:lsdException w:name=\"Colorful List Accent 4\" w:uiPriority=\"72\"/><w:lsdException w:name=\"Colorful Grid Accent 4\" w:uiPriority=\"73\"/><w:lsdException w:name=\"Light Shading Accent 5\" w:uiPriority=\"60\"/><w:lsdException w:name=\"Light List Accent 5\" w:uiPriority=\"61\"/><w:lsdException w:name=\"Light Grid Accent 5\" w:uiPriority=\"62\"/><w:lsdException w:name=\"Medium Shading 1 Accent 5\" w:uiPriority=\"63\"/><w:lsdException w:name=\"Medium Shading 2 Accent 5\" w:uiPriority=\"64\"/><w:lsdException w:name=\"Medium List 1 Accent 5\" w:uiPriority=\"65\"/><w:lsdException w:name=\"Medium List 2 Accent 5\" w:uiPriority=\"66\"/><w:lsdException w:name=\"Medium Grid 1 Accent 5\" w:uiPriority=\"67\"/><w:lsdException w:name=\"Medium Grid 2 Accent 5\" w:uiPriority=\"68\"/><w:lsdException w:name=\"Medium Grid 3 Accent 5\" w:uiPriority=\"69\"/><w:lsdException w:name=\"Dark List Accent 5\" w:uiPriority=\"70\"/><w:lsdException w:name=\"Colorful Shading Accent 5\" w:uiPriority=\"71\"/><w:lsdException w:name=\"Colorful List Accent 5\" w:uiPriority=\"72\"/><w:lsdException w:name=\"Colorful Grid Accent 5\" w:uiPriority=\"73\"/><w:lsdException w:name=\"Light Shading Accent 6\" w:uiPriority=\"60\"/><w:lsdException w:name=\"Light List Accent 6\" w:uiPriority=\"61\"/><w:lsdException w:name=\"Light Grid Accent 6\" w:uiPriority=\"62\"/><w:lsdException w:name=\"Medium Shading 1 Accent 6\" w:uiPriority=\"63\"/><w:lsdException w:name=\"Medium Shading 2 Accent 6\" w:uiPriority=\"64\"/><w:lsdException w:name=\"Medium List 1 Accent 6\" w:uiPriority=\"65\"/><w:lsdException w:name=\"Medium List 2 Accent 6\" w:uiPriority=\"66\"/><w:lsdException w:name=\"Medium Grid 1 Accent 6\" w:uiPriority=\"67\"/><w:lsdException w:name=\"Medium Grid 2 Accent 6\" w:uiPriority=\"68\"/><w:lsdException w:name=\"Medium Grid 3 Accent 6\" w:uiPriority=\"69\"/><w:lsdException w:name=\"Dark List Accent 6\" w:uiPriority=\"70\"/><w:lsdException w:name=\"Colorful Shading Accent 6\" w:uiPriority=\"71\"/><w:lsdException w:name=\"Colorful List Accent 6\" w:uiPriority=\"72\"/><w:lsdException w:name=\"Colorful Grid Accent 6\" w:uiPriority=\"73\"/><w:lsdException w:name=\"Subtle Emphasis\" w:uiPriority=\"19\" w:qFormat=\"1\"/><w:lsdException w:name=\"Intense Emphasis\" w:uiPriority=\"21\" w:qFormat=\"1\"/><w:lsdException w:name=\"Subtle Reference\" w:uiPriority=\"31\" w:qFormat=\"1\"/><w:lsdException w:name=\"Intense Reference\" w:uiPriority=\"32\" w:qFormat=\"1\"/><w:lsdException w:name=\"Book Title\" w:uiPriority=\"33\" w:qFormat=\"1\"/><w:lsdException w:name=\"Bibliography\" w:semiHidden=\"1\" w:uiPriority=\"37\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"TOC Heading\" w:semiHidden=\"1\" w:uiPriority=\"39\" w:unhideWhenUsed=\"1\" w:qFormat=\"1\"/><w:lsdException w:name=\"Plain Table 1\" w:uiPriority=\"41\"/><w:lsdException w:name=\"Plain Table 2\" w:uiPriority=\"42\"/><w:lsdException w:name=\"Plain Table 3\" w:uiPriority=\"43\"/><w:lsdException w:name=\"Plain Table 4\" w:uiPriority=\"44\"/><w:lsdException w:name=\"Plain Table 5\" w:uiPriority=\"45\"/><w:lsdException w:name=\"Grid Table Light\" w:uiPriority=\"40\"/><w:lsdException w:name=\"Grid Table 1 Light\" w:uiPriority=\"46\"/><w:lsdException w:name=\"Grid Table 2\" w:uiPriority=\"47\"/><w:lsdException w:name=\"Grid Table 3\" w:uiPriority=\"48\"/><w:lsdException w:name=\"Grid Table 4\" w:uiPriority=\"49\"/><w:lsdException w:name=\"Grid Table 5 Dark\" w:uiPriority=\"50\"/><w:lsdException w:name=\"Grid Table 6 Colorful\" w:uiPriority=\"51\"/><w:lsdException w:name=\"Grid Table 7 Colorful\" w:uiPriority=\"52\"/><w:lsdException w:name=\"Grid Table 1 Light Accent 1\" w:uiPriority=\"46\"/><w:lsdException w:name=\"Grid Table 2 Accent 1\" w:uiPriority=\"47\"/><w:lsdException w:name=\"Grid Table 3 Accent 1\" w:uiPriority=\"48\"/><w:lsdException w:name=\"Grid Table 4 Accent 1\" w:uiPriority=\"49\"/><w:lsdException w:name=\"Grid Table 5 Dark Accent 1\" w:uiPriority=\"50\"/><w:lsdException w:name=\"Grid Table 6 Colorful Accent 1\" w:uiPriority=\"51\"/><w:lsdException w:name=\"Grid Table 7 Colorful Accent 1\" w:uiPriority=\"52\"/><w:lsdException w:name=\"Grid Table 1 Light Accent 2\" w:uiPriority=\"46\"/><w:lsdException w:name=\"Grid Table 2 Accent 2\" w:uiPriority=\"47\"/><w:lsdException w:name=\"Grid Table 3 Accent 2\" w:uiPriority=\"48\"/><w:lsdException w:name=\"Grid Table 4 Accent 2\" w:uiPriority=\"49\"/><w:lsdException w:name=\"Grid Table 5 Dark Accent 2\" w:uiPriority=\"50\"/><w:lsdException w:name=\"Grid Table 6 Colorful Accent 2\" w:uiPriority=\"51\"/><w:lsdException w:name=\"Grid Table 7 Colorful Accent 2\" w:uiPriority=\"52\"/><w:lsdException w:name=\"Grid Table 1 Light Accent 3\" w:uiPriority=\"46\"/><w:lsdException w:name=\"Grid Table 2 Accent 3\" w:uiPriority=\"47\"/><w:lsdException w:name=\"Grid Table 3 Accent 3\" w:uiPriority=\"48\"/><w:lsdException w:name=\"Grid Table 4 Accent 3\" w:uiPriority=\"49\"/><w:lsdException w:name=\"Grid Table 5 Dark Accent 3\" w:uiPriority=\"50\"/><w:lsdException w:name=\"Grid Table 6 Colorful Accent 3\" w:uiPriority=\"51\"/><w:lsdException w:name=\"Grid Table 7 Colorful Accent 3\" w:uiPriority=\"52\"/><w:lsdException w:name=\"Grid Table 1 Light Accent 4\" w:uiPriority=\"46\"/><w:lsdException w:name=\"Grid Table 2 Accent 4\" w:uiPriority=\"47\"/><w:lsdException w:name=\"Grid Table 3 Accent 4\" w:uiPriority=\"48\"/><w:lsdException w:name=\"Grid Table 4 Accent 4\" w:uiPriority=\"49\"/><w:lsdException w:name=\"Grid Table 5 Dark Accent 4\" w:uiPriority=\"50\"/><w:lsdException w:name=\"Grid Table 6 Colorful Accent 4\" w:uiPriority=\"51\"/><w:lsdException w:name=\"Grid Table 7 Colorful Accent 4\" w:uiPriority=\"52\"/><w:lsdException w:name=\"Grid Table 1 Light Accent 5\" w:uiPriority=\"46\"/><w:lsdException w:name=\"Grid Table 2 Accent 5\" w:uiPriority=\"47\"/><w:lsdException w:name=\"Grid Table 3 Accent 5\" w:uiPriority=\"48\"/><w:lsdException w:name=\"Grid Table 4 Accent 5\" w:uiPriority=\"49\"/><w:lsdException w:name=\"Grid Table 5 Dark Accent 5\" w:uiPriority=\"50\"/><w:lsdException w:name=\"Grid Table 6 Colorful Accent 5\" w:uiPriority=\"51\"/><w:lsdException w:name=\"Grid Table 7 Colorful Accent 5\" w:uiPriority=\"52\"/><w:lsdException w:name=\"Grid Table 1 Light Accent 6\" w:uiPriority=\"46\"/><w:lsdException w:name=\"Grid Table 2 Accent 6\" w:uiPriority=\"47\"/><w:lsdException w:name=\"Grid Table 3 Accent 6\" w:uiPriority=\"48\"/><w:lsdException w:name=\"Grid Table 4 Accent 6\" w:uiPriority=\"49\"/><w:lsdException w:name=\"Grid Table 5 Dark Accent 6\" w:uiPriority=\"50\"/><w:lsdException w:name=\"Grid Table 6 Colorful Accent 6\" w:uiPriority=\"51\"/><w:lsdException w:name=\"Grid Table 7 Colorful Accent 6\" w:uiPriority=\"52\"/><w:lsdException w:name=\"List Table 1 Light\" w:uiPriority=\"46\"/><w:lsdException w:name=\"List Table 2\" w:uiPriority=\"47\"/><w:lsdException w:name=\"List Table 3\" w:uiPriority=\"48\"/><w:lsdException w:name=\"List Table 4\" w:uiPriority=\"49\"/><w:lsdException w:name=\"List Table 5 Dark\" w:uiPriority=\"50\"/><w:lsdException w:name=\"List Table 6 Colorful\" w:uiPriority=\"51\"/><w:lsdException w:name=\"List Table 7 Colorful\" w:uiPriority=\"52\"/><w:lsdException w:name=\"List Table 1 Light Accent 1\" w:uiPriority=\"46\"/><w:lsdException w:name=\"List Table 2 Accent 1\" w:uiPriority=\"47\"/><w:lsdException w:name=\"List Table 3 Accent 1\" w:uiPriority=\"48\"/><w:lsdException w:name=\"List Table 4 Accent 1\" w:uiPriority=\"49\"/><w:lsdException w:name=\"List Table 5 Dark Accent 1\" w:uiPriority=\"50\"/><w:lsdException w:name=\"List Table 6 Colorful Accent 1\" w:uiPriority=\"51\"/><w:lsdException w:name=\"List Table 7 Colorful Accent 1\" w:uiPriority=\"52\"/><w:lsdException w:name=\"List Table 1 Light Accent 2\" w:uiPriority=\"46\"/><w:lsdException w:name=\"List Table 2 Accent 2\" w:uiPriority=\"47\"/><w:lsdException w:name=\"List Table 3 Accent 2\" w:uiPriority=\"48\"/><w:lsdException w:name=\"List Table 4 Accent 2\" w:uiPriority=\"49\"/><w:lsdException w:name=\"List Table 5 Dark Accent 2\" w:uiPriority=\"50\"/><w:lsdException w:name=\"List Table 6 Colorful Accent 2\" w:uiPriority=\"51\"/><w:lsdException w:name=\"List Table 7 Colorful Accent 2\" w:uiPriority=\"52\"/><w:lsdException w:name=\"List Table 1 Light Accent 3\" w:uiPriority=\"46\"/><w:lsdException w:name=\"List Table 2 Accent 3\" w:uiPriority=\"47\"/><w:lsdException w:name=\"List Table 3 Accent 3\" w:uiPriority=\"48\"/><w:lsdException w:name=\"List Table 4 Accent 3\" w:uiPriority=\"49\"/><w:lsdException w:name=\"List Table 5 Dark Accent 3\" w:uiPriority=\"50\"/><w:lsdException w:name=\"List Table 6 Colorful Accent 3\" w:uiPriority=\"51\"/><w:lsdException w:name=\"List Table 7 Colorful Accent 3\" w:uiPriority=\"52\"/><w:lsdException w:name=\"List Table 1 Light Accent 4\" w:uiPriority=\"46\"/><w:lsdException w:name=\"List Table 2 Accent 4\" w:uiPriority=\"47\"/><w:lsdException w:name=\"List Table 3 Accent 4\" w:uiPriority=\"48\"/><w:lsdException w:name=\"List Table 4 Accent 4\" w:uiPriority=\"49\"/><w:lsdException w:name=\"List Table 5 Dark Accent 4\" w:uiPriority=\"50\"/><w:lsdException w:name=\"List Table 6 Colorful Accent 4\" w:uiPriority=\"51\"/><w:lsdException w:name=\"List Table 7 Colorful Accent 4\" w:uiPriority=\"52\"/><w:lsdException w:name=\"List Table 1 Light Accent 5\" w:uiPriority=\"46\"/><w:lsdException w:name=\"List Table 2 Accent 5\" w:uiPriority=\"47\"/><w:lsdException w:name=\"List Table 3 Accent 5\" w:uiPriority=\"48\"/><w:lsdException w:name=\"List Table 4 Accent 5\" w:uiPriority=\"49\"/><w:lsdException w:name=\"List Table 5 Dark Accent 5\" w:uiPriority=\"50\"/><w:lsdException w:name=\"List Table 6 Colorful Accent 5\" w:uiPriority=\"51\"/><w:lsdException w:name=\"List Table 7 Colorful Accent 5\" w:uiPriority=\"52\"/><w:lsdException w:name=\"List Table 1 Light Accent 6\" w:uiPriority=\"46\"/><w:lsdException w:name=\"List Table 2 Accent 6\" w:uiPriority=\"47\"/><w:lsdException w:name=\"List Table 3 Accent 6\" w:uiPriority=\"48\"/><w:lsdException w:name=\"List Table 4 Accent 6\" w:uiPriority=\"49\"/><w:lsdException w:name=\"List Table 5 Dark Accent 6\" w:uiPriority=\"50\"/><w:lsdException w:name=\"List Table 6 Colorful Accent 6\" w:uiPriority=\"51\"/><w:lsdException w:name=\"List Table 7 Colorful Accent 6\" w:uiPriority=\"52\"/><w:lsdException w:name=\"Mention\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Smart Hyperlink\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Hashtag\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Unresolved Mention\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/><w:lsdException w:name=\"Smart Link\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\"/></w:latentStyles><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"a\"><w:name w:val=\"Normal\"/><w:qFormat/><w:rsid w:val=\"004D5293\"/><w:pPr><w:widowControl w:val=\"0\"/></w:pPr><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:eastAsia=\"新細明體\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/><w:szCs w:val=\"24\"/></w:rPr></w:style><w:style w:type=\"character\" w:default=\"1\" w:styleId=\"a0\"><w:name w:val=\"Default Paragraph Font\"/><w:uiPriority w:val=\"1\"/><w:semiHidden/><w:unhideWhenUsed/></w:style><w:style w:type=\"table\" w:default=\"1\" w:styleId=\"a1\"><w:name w:val=\"Normal Table\"/><w:uiPriority w:val=\"99\"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w=\"0\" w:type=\"dxa\"/><w:tblCellMar><w:top w:w=\"0\" w:type=\"dxa\"/><w:left w:w=\"108\" w:type=\"dxa\"/><w:bottom w:w=\"0\" w:type=\"dxa\"/><w:right w:w=\"108\" w:type=\"dxa\"/></w:tblCellMar></w:tblPr></w:style><w:style w:type=\"numbering\" w:default=\"1\" w:styleId=\"a2\"><w:name w:val=\"No List\"/><w:uiPriority w:val=\"99\"/><w:semiHidden/><w:unhideWhenUsed/></w:style><w:style w:type=\"paragraph\" w:styleId=\"a3\"><w:name w:val=\"header\"/><w:basedOn w:val=\"a\"/><w:link w:val=\"a4\"/><w:uiPriority w:val=\"99\"/><w:unhideWhenUsed/><w:rsid w:val=\"00D04FAC\"/><w:pPr><w:tabs><w:tab w:val=\"center\" w:pos=\"4153\"/><w:tab w:val=\"right\" w:pos=\"8306\"/></w:tabs><w:snapToGrid w:val=\"0\"/></w:pPr><w:rPr><w:sz w:val=\"20\"/><w:szCs w:val=\"20\"/></w:rPr></w:style><w:style w:type=\"character\" w:customStyle=\"1\" w:styleId=\"a4\"><w:name w:val=\"頁首 字元\"/><w:basedOn w:val=\"a0\"/><w:link w:val=\"a3\"/><w:uiPriority w:val=\"99\"/><w:rsid w:val=\"00D04FAC\"/><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:eastAsia=\"新細明體\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/><w:sz w:val=\"20\"/><w:szCs w:val=\"20\"/></w:rPr></w:style><w:style w:type=\"paragraph\" w:styleId=\"a5\"><w:name w:val=\"footer\"/><w:basedOn w:val=\"a\"/><w:link w:val=\"a6\"/><w:uiPriority w:val=\"99\"/><w:unhideWhenUsed/><w:rsid w:val=\"00D04FAC\"/><w:pPr><w:tabs><w:tab w:val=\"center\" w:pos=\"4153\"/><w:tab w:val=\"right\" w:pos=\"8306\"/></w:tabs><w:snapToGrid w:val=\"0\"/></w:pPr><w:rPr><w:sz w:val=\"20\"/><w:szCs w:val=\"20\"/></w:rPr></w:style><w:style w:type=\"character\" w:customStyle=\"1\" w:styleId=\"a6\"><w:name w:val=\"頁尾 字元\"/><w:basedOn w:val=\"a0\"/><w:link w:val=\"a5\"/><w:uiPriority w:val=\"99\"/><w:rsid w:val=\"00D04FAC\"/><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:eastAsia=\"新細明體\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/><w:sz w:val=\"20\"/><w:szCs w:val=\"20\"/></w:rPr></w:style><w:style w:type=\"paragraph\" w:styleId=\"a7\"><w:name w:val=\"Balloon Text\"/><w:basedOn w:val=\"a\"/><w:link w:val=\"a8\"/><w:uiPriority w:val=\"99\"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val=\"009A0416\"/><w:rPr><w:rFonts w:asciiTheme=\"majorHAnsi\" w:eastAsiaTheme=\"majorEastAsia\" w:hAnsiTheme=\"majorHAnsi\" w:cstheme=\"majorBidi\"/><w:sz w:val=\"18\"/><w:szCs w:val=\"18\"/></w:rPr></w:style><w:style w:type=\"character\" w:customStyle=\"1\" w:styleId=\"a8\"><w:name w:val=\"註解方塊文字 字元\"/><w:basedOn w:val=\"a0\"/><w:link w:val=\"a7\"/><w:uiPriority w:val=\"99\"/><w:semiHidden/><w:rsid w:val=\"009A0416\"/><w:rPr><w:rFonts w:asciiTheme=\"majorHAnsi\" w:eastAsiaTheme=\"majorEastAsia\" w:hAnsiTheme=\"majorHAnsi\" w:cstheme=\"majorBidi\"/><w:sz w:val=\"18\"/><w:szCs w:val=\"18\"/></w:rPr></w:style><w:style w:type=\"paragraph\" w:styleId=\"a9\"><w:name w:val=\"Date\"/><w:basedOn w:val=\"a\"/><w:next w:val=\"a\"/><w:link w:val=\"aa\"/><w:uiPriority w:val=\"99\"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val=\"000E0C05\"/><w:pPr><w:jc w:val=\"right\"/></w:pPr></w:style><w:style w:type=\"character\" w:customStyle=\"1\" w:styleId=\"aa\"><w:name w:val=\"日期 字元\"/><w:basedOn w:val=\"a0\"/><w:link w:val=\"a9\"/><w:uiPriority w:val=\"99\"/><w:semiHidden/><w:rsid w:val=\"000E0C05\"/><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:eastAsia=\"新細明體\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/><w:szCs w:val=\"24\"/></w:rPr></w:style></w:styles>"],["word/settings.xml","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:settings xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:sl=\"http://schemas.openxmlformats.org/schemaLibrary/2006/main\" mc:Ignorable=\"w14 w15 w16se w16cid w16 w16cex w16sdtdh\"><w:zoom w:percent=\"100\"/><w:bordersDoNotSurroundHeader/><w:bordersDoNotSurroundFooter/><w:proofState w:spelling=\"clean\"/><w:defaultTabStop w:val=\"480\"/><w:displayHorizontalDrawingGridEvery w:val=\"0\"/><w:displayVerticalDrawingGridEvery w:val=\"2\"/><w:characterSpacingControl w:val=\"compressPunctuation\"/><w:hdrShapeDefaults><o:shapedefaults v:ext=\"edit\" spidmax=\"34817\"/></w:hdrShapeDefaults><w:footnotePr><w:footnote w:id=\"-1\"/><w:footnote w:id=\"0\"/></w:footnotePr><w:endnotePr><w:endnote w:id=\"-1\"/><w:endnote w:id=\"0\"/></w:endnotePr><w:compat><w:spaceForUL/><w:balanceSingleByteDoubleByteWidth/><w:doNotLeaveBackslashAlone/><w:ulTrailSpace/><w:doNotExpandShiftReturn/><w:adjustLineHeightInTable/><w:useFELayout/><w:compatSetting w:name=\"compatibilityMode\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"15\"/><w:compatSetting w:name=\"overrideTableStyleFontSizeAndJustification\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"1\"/><w:compatSetting w:name=\"enableOpenTypeFeatures\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"1\"/><w:compatSetting w:name=\"doNotFlipMirrorIndents\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"1\"/><w:compatSetting w:name=\"differentiateMultirowTableHeaders\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"1\"/><w:compatSetting w:name=\"useWord2013TrackBottomHyphenation\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"1\"/></w:compat><m:mathPr><m:mathFont m:val=\"Cambria Math\"/><m:brkBin m:val=\"before\"/><m:brkBinSub m:val=\"--\"/><m:smallFrac m:val=\"0\"/><m:dispDef/><m:lMargin m:val=\"0\"/><m:rMargin m:val=\"0\"/><m:defJc m:val=\"centerGroup\"/><m:wrapIndent m:val=\"1440\"/><m:intLim m:val=\"subSup\"/><m:naryLim m:val=\"undOvr\"/></m:mathPr><w:themeFontLang w:val=\"en-US\" w:eastAsia=\"zh-TW\"/><w:clrSchemeMapping w:bg1=\"light1\" w:t1=\"dark1\" w:bg2=\"light2\" w:t2=\"dark2\" w:accent1=\"accent1\" w:accent2=\"accent2\" w:accent3=\"accent3\" w:accent4=\"accent4\" w:accent5=\"accent5\" w:accent6=\"accent6\" w:hyperlink=\"hyperlink\" w:followedHyperlink=\"followedHyperlink\"/><w:shapeDefaults><o:shapedefaults v:ext=\"edit\" spidmax=\"34817\"/><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout></w:shapeDefaults><w:decimalSymbol w:val=\".\"/><w:listSeparator w:val=\",\"/><w14:docId w14:val=\"432979C8\"/><w15:docId w15:val=\"{D7DC6BE9-03B9-493F-825D-4CB533B68D16}\"/></w:settings>"],["word/theme/theme1.xml","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office 佈景主題\"><a:themeElements><a:clrScheme name=\"Office\"><a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1><a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1><a:dk2><a:srgbClr val=\"1F497D\"/></a:dk2><a:lt2><a:srgbClr val=\"EEECE1\"/></a:lt2><a:accent1><a:srgbClr val=\"4F81BD\"/></a:accent1><a:accent2><a:srgbClr val=\"C0504D\"/></a:accent2><a:accent3><a:srgbClr val=\"9BBB59\"/></a:accent3><a:accent4><a:srgbClr val=\"8064A2\"/></a:accent4><a:accent5><a:srgbClr val=\"4BACC6\"/></a:accent5><a:accent6><a:srgbClr val=\"F79646\"/></a:accent6><a:hlink><a:srgbClr val=\"0000FF\"/></a:hlink><a:folHlink><a:srgbClr val=\"800080\"/></a:folHlink></a:clrScheme><a:fontScheme name=\"Office\"><a:majorFont><a:latin typeface=\"Cambria\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"ＭＳ ゴシック\"/><a:font script=\"Hang\" typeface=\"맑은 고딕\"/><a:font script=\"Hans\" typeface=\"宋体\"/><a:font script=\"Hant\" typeface=\"新細明體\"/><a:font script=\"Arab\" typeface=\"Times New Roman\"/><a:font script=\"Hebr\" typeface=\"Times New Roman\"/><a:font script=\"Thai\" typeface=\"Angsana New\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"MoolBoran\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Times New Roman\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/><a:font script=\"Geor\" typeface=\"Sylfaen\"/></a:majorFont><a:minorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"ＭＳ 明朝\"/><a:font script=\"Hang\" typeface=\"맑은 고딕\"/><a:font script=\"Hans\" typeface=\"宋体\"/><a:font script=\"Hant\" typeface=\"新細明體\"/><a:font script=\"Arab\" typeface=\"Arial\"/><a:font script=\"Hebr\" typeface=\"Arial\"/><a:font script=\"Thai\" typeface=\"Cordia New\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"DaunPenh\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Arial\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/><a:font script=\"Geor\" typeface=\"Sylfaen\"/></a:minorFont></a:fontScheme><a:fmtScheme name=\"Office\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"50000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs><a:gs pos=\"35000\"><a:schemeClr val=\"phClr\"><a:tint val=\"37000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:tint val=\"15000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"16200000\" scaled=\"1\"/></a:gradFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:shade val=\"51000\"/><a:satMod val=\"130000\"/></a:schemeClr></a:gs><a:gs pos=\"80000\"><a:schemeClr val=\"phClr\"><a:shade val=\"93000\"/><a:satMod val=\"130000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"94000\"/><a:satMod val=\"135000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"16200000\" scaled=\"0\"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"><a:shade val=\"95000\"/><a:satMod val=\"105000\"/></a:schemeClr></a:solidFill><a:prstDash val=\"solid\"/></a:ln><a:ln w=\"25400\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln><a:ln w=\"38100\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"20000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"38000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"23000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"35000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"23000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"35000\"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst=\"orthographicFront\"><a:rot lat=\"0\" lon=\"0\" rev=\"0\"/></a:camera><a:lightRig rig=\"threePt\" dir=\"t\"><a:rot lat=\"0\" lon=\"0\" rev=\"1200000\"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w=\"63500\" h=\"25400\"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"40000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs><a:gs pos=\"40000\"><a:schemeClr val=\"phClr\"><a:tint val=\"45000\"/><a:shade val=\"99000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"20000\"/><a:satMod val=\"255000\"/></a:schemeClr></a:gs></a:gsLst><a:path path=\"circle\"><a:fillToRect l=\"50000\" t=\"-80000\" r=\"50000\" b=\"180000\"/></a:path></a:gradFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"80000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"30000\"/><a:satMod val=\"200000\"/></a:schemeClr></a:gs></a:gsLst><a:path path=\"circle\"><a:fillToRect l=\"50000\" t=\"50000\" r=\"50000\" b=\"50000\"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>"],["word/fontTable.xml","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:fonts xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" mc:Ignorable=\"w14 w15 w16se w16cid w16 w16cex w16sdtdh\"><w:font w:name=\"Calibri\"><w:panose1 w:val=\"020F0502020204030204\"/><w:charset w:val=\"00\"/><w:family w:val=\"swiss\"/><w:pitch w:val=\"variable\"/><w:sig w:usb0=\"E4002EFF\" w:usb1=\"C000247B\" w:usb2=\"00000009\" w:usb3=\"00000000\" w:csb0=\"000001FF\" w:csb1=\"00000000\"/></w:font><w:font w:name=\"新細明體\"><w:altName w:val=\"PMingLiU\"/><w:panose1 w:val=\"02020500000000000000\"/><w:charset w:val=\"88\"/><w:family w:val=\"roman\"/><w:pitch w:val=\"variable\"/><w:sig w:usb0=\"A00002FF\" w:usb1=\"28CFFCFA\" w:usb2=\"00000016\" w:usb3=\"00000000\" w:csb0=\"00100001\" w:csb1=\"00000000\"/></w:font><w:font w:name=\"Times New Roman\"><w:panose1 w:val=\"02020603050405020304\"/><w:charset w:val=\"00\"/><w:family w:val=\"roman\"/><w:pitch w:val=\"variable\"/><w:sig w:usb0=\"E0002EFF\" w:usb1=\"C000785B\" w:usb2=\"00000009\" w:usb3=\"00000000\" w:csb0=\"000001FF\" w:csb1=\"00000000\"/></w:font><w:font w:name=\"Cambria\"><w:panose1 w:val=\"02040503050406030204\"/><w:charset w:val=\"00\"/><w:family w:val=\"roman\"/><w:pitch w:val=\"variable\"/><w:sig w:usb0=\"E00006FF\" w:usb1=\"420024FF\" w:usb2=\"02000000\" w:usb3=\"00000000\" w:csb0=\"0000019F\" w:csb1=\"00000000\"/></w:font><w:font w:name=\"金梅毛顏楷\"><w:altName w:val=\"細明體\"/><w:charset w:val=\"88\"/><w:family w:val=\"modern\"/><w:pitch w:val=\"fixed\"/><w:sig w:usb0=\"00000001\" w:usb1=\"08080000\" w:usb2=\"00000010\" w:usb3=\"00000000\" w:csb0=\"00100000\" w:csb1=\"00000000\"/></w:font><w:font w:name=\"標楷體\"><w:altName w:val=\"DF Kai Shu\"/><w:panose1 w:val=\"03000509000000000000\"/><w:charset w:val=\"88\"/><w:family w:val=\"script\"/><w:pitch w:val=\"fixed\"/><w:sig w:usb0=\"00000003\" w:usb1=\"080E0000\" w:usb2=\"00000016\" w:usb3=\"00000000\" w:csb0=\"00100001\" w:csb1=\"00000000\"/></w:font></w:fonts>"],["word/webSettings.xml","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:webSettings xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" mc:Ignorable=\"w14 w15 w16se w16cid w16 w16cex w16sdtdh\"><w:optimizeForBrowser/><w:relyOnVML/><w:allowPNG/></w:webSettings>"],["word/footnotes.xml","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:footnotes xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:aink=\"http://schemas.microsoft.com/office/drawing/2016/ink\" xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14\"><w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote></w:footnotes>"],["word/endnotes.xml","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:endnotes xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:aink=\"http://schemas.microsoft.com/office/drawing/2016/ink\" xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14\"><w:endnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote></w:endnotes>"],["docProps/core.xml","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dc:title></dc:title><dc:creator></dc:creator><cp:lastModifiedBy></cp:lastModifiedBy><cp:revision>26</cp:revision><cp:lastPrinted>2026-11-10T04:17:00Z</cp:lastPrinted><dcterms:created xsi:type=\"dcterms:W3CDTF\">2018-11-12T07:13:00Z</dcterms:created><dcterms:modified xsi:type=\"dcterms:W3CDTF\">2026-11-10T04:36:00Z</dcterms:modified></cp:coreProperties>"],["docProps/app.xml","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><Template>Normal</Template><TotalTime>304</TotalTime><Pages>1</Pages><Words>33</Words><Characters>189</Characters><Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity><Lines>1</Lines><Paragraphs>1</Paragraphs><ScaleCrop>false</ScaleCrop><Company></Company><LinksUpToDate>false</LinksUpToDate><CharactersWithSpaces>221</CharactersWithSpaces><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0000</AppVersion></Properties>"],["customXml/item1.xml","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?><b:Sources SelectedStyle=\"\\APA.XSL\" StyleName=\"APA\" xmlns:b=\"http://schemas.openxmlformats.org/officeDocument/2006/bibliography\" xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/bibliography\"></b:Sources>"],["customXml/itemProps1.xml","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?>\n<ds:datastoreItem ds:itemID=\"{E9F39D70-53A1-4A4B-8171-506AEA0BFD8C}\" xmlns:ds=\"http://schemas.openxmlformats.org/officeDocument/2006/customXml\"><ds:schemaRefs><ds:schemaRef ds:uri=\"http://schemas.openxmlformats.org/officeDocument/2006/bibliography\"/></ds:schemaRefs></ds:datastoreItem>"],["customXml/_rels/item1.xml.rels","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps\" Target=\"itemProps1.xml\"/></Relationships>"]],"pre":"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:aink=\"http://schemas.microsoft.com/office/drawing/2016/ink\" xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14\"><w:body>","body":"<w:p><w:pPr><w:jc w:val=\"center\"/><w:rPr><w:rFonts w:ascii=\"金梅毛顏楷\" w:eastAsia=\"金梅毛顏楷\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"16\"/><w:szCs w:val=\"16\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:jc w:val=\"center\"/><w:rPr><w:rFonts w:ascii=\"金梅毛顏楷\" w:eastAsia=\"金梅毛顏楷\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"16\"/><w:szCs w:val=\"16\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:jc w:val=\"center\"/><w:rPr><w:rFonts w:ascii=\"金梅毛顏楷\" w:eastAsia=\"金梅毛顏楷\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"6\"/><w:szCs w:val=\"6\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:jc w:val=\"right\"/><w:rPr><w:rFonts w:ascii=\"金梅毛顏楷\" w:eastAsia=\"金梅毛顏楷\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"16\"/><w:szCs w:val=\"16\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:noProof/><w:sz w:val=\"52\"/><w:szCs w:val=\"52\"/></w:rPr><mc:AlternateContent><mc:Choice Requires=\"wps\"><w:drawing><wp:anchor distT=\"45720\" distB=\"45720\" distL=\"114300\" distR=\"114300\" simplePos=\"0\" relativeHeight=\"251659264\" behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\" wp14:anchorId=\"481C6EE1\" wp14:editId=\"78B6F1BE\"><wp:simplePos x=\"0\" y=\"0\"/><wp:positionH relativeFrom=\"column\"><wp:posOffset>2009775</wp:posOffset></wp:positionH><wp:positionV relativeFrom=\"paragraph\"><wp:posOffset>228600</wp:posOffset></wp:positionV><wp:extent cx=\"1419225\" cy=\"819150\"/><wp:effectExtent l=\"0\" t=\"0\" r=\"9525\" b=\"0\"/><wp:wrapNone/><wp:docPr id=\"217\" name=\"文字方塊 2\"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/></wp:cNvGraphicFramePr><a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"><wps:wsp><wps:cNvSpPr txBox=\"1\"><a:spLocks noChangeArrowheads=\"1\"/></wps:cNvSpPr><wps:spPr bwMode=\"auto\"><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1419225\" cy=\"819150\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val=\"FFFFFF\"/></a:solidFill><a:ln w=\"9525\"><a:noFill/><a:miter lim=\"800000\"/><a:headEnd/><a:tailEnd/></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:pPr><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"90\"/><w:szCs w:val=\"90\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\" w:hint=\"eastAsia\"/><w:sz w:val=\"90\"/><w:szCs w:val=\"90\"/></w:rPr><w:t>獎狀</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr rot=\"0\" vert=\"horz\" wrap=\"square\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" anchor=\"t\" anchorCtr=\"0\"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic><wp14:sizeRelH relativeFrom=\"margin\"><wp14:pctWidth>0</wp14:pctWidth></wp14:sizeRelH><wp14:sizeRelV relativeFrom=\"margin\"><wp14:pctHeight>0</wp14:pctHeight></wp14:sizeRelV></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shapetype w14:anchorId=\"481C6EE1\" id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype><v:shape id=\"文字方塊 2\" o:spid=\"_x0000_s1026\" type=\"#_x0000_t202\" style=\"position:absolute;left:0;text-align:left;margin-left:158.25pt;margin-top:18pt;width:111.75pt;height:64.5pt;z-index:251659264;visibility:visible;mso-wrap-style:square;mso-width-percent:0;mso-height-percent:0;mso-wrap-distance-left:9pt;mso-wrap-distance-top:3.6pt;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:3.6pt;mso-position-horizontal:absolute;mso-position-horizontal-relative:text;mso-position-vertical:absolute;mso-position-vertical-relative:text;mso-width-percent:0;mso-height-percent:0;mso-width-relative:margin;mso-height-relative:margin;v-text-anchor:top\" o:gfxdata=\"UEsDBBQABgAIAAAAIQC2gziS/gAAAOEBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbJSRQU7DMBBF&#xA;90jcwfIWJU67QAgl6YK0S0CoHGBkTxKLZGx5TGhvj5O2G0SRWNoz/78nu9wcxkFMGNg6quQqL6RA&#xA;0s5Y6ir5vt9lD1JwBDIwOMJKHpHlpr69KfdHjyxSmriSfYz+USnWPY7AufNIadK6MEJMx9ApD/oD&#xA;OlTrorhX2lFEilmcO2RdNtjC5xDF9pCuTyYBB5bi6bQ4syoJ3g9WQ0ymaiLzg5KdCXlKLjvcW893&#xA;SUOqXwnz5DrgnHtJTxOsQfEKIT7DmDSUCaxw7Rqn8787ZsmRM9e2VmPeBN4uqYvTtW7jvijg9N/y&#xA;JsXecLq0q+WD6m8AAAD//wMAUEsDBBQABgAIAAAAIQA4/SH/1gAAAJQBAAALAAAAX3JlbHMvLnJl&#xA;bHOkkMFqwzAMhu+DvYPRfXGawxijTi+j0GvpHsDYimMaW0Yy2fr2M4PBMnrbUb/Q94l/f/hMi1qR&#xA;JVI2sOt6UJgd+ZiDgffL8ekFlFSbvV0oo4EbChzGx4f9GRdb25HMsYhqlCwG5lrLq9biZkxWOiqY&#xA;22YiTra2kYMu1l1tQD30/bPm3wwYN0x18gb45AdQl1tp5j/sFB2T0FQ7R0nTNEV3j6o9feQzro1i&#xA;OWA14Fm+Q8a1a8+Bvu/d/dMb2JY5uiPbhG/ktn4cqGU/er3pcvwCAAD//wMAUEsDBBQABgAIAAAA&#xA;IQCHGRmFNQIAACEEAAAOAAAAZHJzL2Uyb0RvYy54bWysU11u2zAMfh+wOwh6XxwbydoYcYouXYYB&#xA;3Q/Q7QCKLMfCJFGTlNjdBQb0AN3zDrAD7EDtOUbJaRp0b8P8IIgm+enjR3J+1mtFdsJ5Caai+WhM&#xA;iTAcamk2Ff38afXilBIfmKmZAiMqei08PVs8fzbvbCkKaEHVwhEEMb7sbEXbEGyZZZ63QjM/AisM&#xA;OhtwmgU03SarHesQXausGI9fZh242jrgwnv8ezE46SLhN43g4UPTeBGIqihyC+l06VzHM1vMWblx&#xA;zLaS72mwf2ChmTT46AHqggVGtk7+BaUld+ChCSMOOoOmkVykGrCafPykmquWWZFqQXG8Pcjk/x8s&#xA;f7/76IisK1rkJ5QYprFJ97ff7379uL/9fffzhhRRo876EkOvLAaH/hX02OtUr7eXwL94YmDZMrMR&#xA;585B1wpWI8c8ZmZHqQOOjyDr7h3U+BTbBkhAfeN0FBAlIYiOvbo+9Ef0gfD45CSfFcWUEo6+03yW&#xA;T1MDM1Y+ZFvnwxsBmsRLRR32P6Gz3aUPkQ0rH0LiYx6UrFdSqWS4zXqpHNkxnJVV+lIBT8KUIV1F&#xA;Z1PkEbMMxPw0RloGnGUlNZIbx2+YrqjGa1OnkMCkGu7IRJm9PFGRQZvQr3sMjJqtob5GoRwMM4s7&#xA;hpcW3DdKOpzXivqvW+YEJeqtQbFn+WQSBzwZk+lJgYY79qyPPcxwhKpooGS4LkNaiqGic2xKI5Ne&#xA;j0z2XHEOk4z7nYmDfmynqMfNXvwBAAD//wMAUEsDBBQABgAIAAAAIQBqWv5v3QAAAAoBAAAPAAAA&#xA;ZHJzL2Rvd25yZXYueG1sTI/BTsMwEETvSPyDtUhcELULjQshTgVIIK4t/QAn3iYR8TqK3Sb9e5YT&#xA;vc1on2Znis3se3HCMXaBDCwXCgRSHVxHjYH998f9E4iYLDnbB0IDZ4ywKa+vCpu7MNEWT7vUCA6h&#xA;mFsDbUpDLmWsW/Q2LsKAxLdDGL1NbMdGutFOHO57+aCUlt52xB9aO+B7i/XP7ugNHL6mu+x5qj7T&#xA;fr1d6TfbratwNub2Zn59AZFwTv8w/NXn6lBypyocyUXRG3hc6oxRFpo3MZCtFIuKSZ0pkGUhLyeU&#xA;vwAAAP//AwBQSwECLQAUAAYACAAAACEAtoM4kv4AAADhAQAAEwAAAAAAAAAAAAAAAAAAAAAAW0Nv&#xA;bnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQA4/SH/1gAAAJQBAAALAAAAAAAAAAAAAAAA&#xA;AC8BAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQCHGRmFNQIAACEEAAAOAAAAAAAAAAAAAAAA&#xA;AC4CAABkcnMvZTJvRG9jLnhtbFBLAQItABQABgAIAAAAIQBqWv5v3QAAAAoBAAAPAAAAAAAAAAAA&#xA;AAAAAI8EAABkcnMvZG93bnJldi54bWxQSwUGAAAAAAQABADzAAAAmQUAAAAA&#xA;\" stroked=\"f\"><v:textbox><w:txbxContent><w:p><w:pPr><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\" w:hint=\"eastAsia\"/><w:sz w:val=\"90\"/><w:szCs w:val=\"90\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\" w:hint=\"eastAsia\"/><w:sz w:val=\"90\"/><w:szCs w:val=\"90\"/></w:rPr><w:t>獎狀</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p><w:p><w:pPr><w:rPr><w:rFonts w:ascii=\"金梅毛顏楷\" w:eastAsia=\"金梅毛顏楷\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:jc w:val=\"right\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:jc w:val=\"right\"/><w:rPr><w:rFonts w:ascii=\"金梅毛顏楷\" w:eastAsia=\"金梅毛顏楷\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"20\"/><w:szCs w:val=\"20\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\" w:hint=\"eastAsia\"/><w:sz w:val=\"20\"/><w:szCs w:val=\"20\"/></w:rPr><w:t xml:space=\"preserve\">@@NO@@</w:t></w:r></w:p><w:p><w:pPr><w:jc w:val=\"center\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:jc w:val=\"center\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"52\"/><w:szCs w:val=\"52\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\" w:hint=\"eastAsia\"/><w:sz w:val=\"52\"/><w:szCs w:val=\"52\"/></w:rPr><w:t xml:space=\"preserve\">@@TITLE@@</w:t></w:r></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"10\"/><w:szCs w:val=\"10\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:ind w:firstLineChars=\"400\" w:firstLine=\"1760\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"44\"/><w:szCs w:val=\"44\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\" w:hint=\"eastAsia\"/><w:sz w:val=\"44\"/><w:szCs w:val=\"44\"/></w:rPr><w:t xml:space=\"preserve\">@@L1@@</w:t></w:r></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"44\"/><w:szCs w:val=\"44\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\" w:hint=\"eastAsia\"/><w:sz w:val=\"44\"/><w:szCs w:val=\"44\"/></w:rPr><w:t xml:space=\"preserve\">@@L2@@</w:t></w:r></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"44\"/><w:szCs w:val=\"44\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\" w:hint=\"eastAsia\"/><w:sz w:val=\"44\"/><w:szCs w:val=\"44\"/></w:rPr><w:t xml:space=\"preserve\">@@L3@@</w:t></w:r></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:ind w:firstLineChars=\"300\" w:firstLine=\"1320\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\" w:hint=\"eastAsia\"/><w:sz w:val=\"44\"/><w:szCs w:val=\"44\"/></w:rPr><w:t xml:space=\"preserve\">@@L4@@</w:t></w:r></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:ind w:firstLineChars=\"300\" w:firstLine=\"60\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:ind w:firstLineChars=\"300\" w:firstLine=\"60\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:ind w:firstLineChars=\"300\" w:firstLine=\"60\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:ind w:firstLineChars=\"300\" w:firstLine=\"60\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:ind w:firstLineChars=\"300\" w:firstLine=\"60\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:ind w:firstLineChars=\"300\" w:firstLine=\"60\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:ind w:firstLineChars=\"300\" w:firstLine=\"60\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:jc w:val=\"center\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\" w:hint=\"eastAsia\"/><w:sz w:val=\"48\"/><w:szCs w:val=\"48\"/></w:rPr><w:t xml:space=\"preserve\"> </w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\" w:hint=\"eastAsia\"/><w:sz w:val=\"48\"/><w:szCs w:val=\"48\"/></w:rPr><w:t>成績優異，值得嘉許。</w:t></w:r></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:jc w:val=\"center\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:jc w:val=\"center\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:jc w:val=\"center\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:jc w:val=\"center\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:jc w:val=\"center\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:jc w:val=\"center\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:ind w:firstLineChars=\"400\" w:firstLine=\"1920\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"48\"/><w:szCs w:val=\"48\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\" w:hint=\"eastAsia\"/><w:sz w:val=\"48\"/><w:szCs w:val=\"48\"/></w:rPr><w:t>特頒此狀，以茲鼓勵。</w:t></w:r></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:ind w:firstLineChars=\"400\" w:firstLine=\"1920\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"48\"/><w:szCs w:val=\"48\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:ind w:firstLineChars=\"400\" w:firstLine=\"80\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:jc w:val=\"center\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:jc w:val=\"center\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:jc w:val=\"center\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:jc w:val=\"center\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:jc w:val=\"center\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:jc w:val=\"center\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:snapToGrid w:val=\"0\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:jc w:val=\"center\"/><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\"/><w:spacing w:val=\"1\"/><w:w w:val=\"76\"/><w:kern w:val=\"0\"/><w:sz w:val=\"48\"/><w:szCs w:val=\"48\"/></w:rPr>@@SECT@@</w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\" w:hint=\"eastAsia\"/><w:spacing w:val=\"98\"/><w:kern w:val=\"0\"/><w:sz w:val=\"48\"/><w:szCs w:val=\"48\"/><w:fitText w:val=\"7920\" w:id=\"-597457408\"/></w:rPr><w:t xml:space=\"preserve\">@@DATE_A@@</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\" w:hint=\"eastAsia\"/><w:spacing w:val=\"2\"/><w:kern w:val=\"0\"/><w:sz w:val=\"48\"/><w:szCs w:val=\"48\"/><w:fitText w:val=\"7920\" w:id=\"-597457408\"/></w:rPr><w:t xml:space=\"preserve\">@@DATE_B@@</w:t></w:r></w:p>","sect":"<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1440\" w:right=\"1800\" w:bottom=\"1440\" w:left=\"1800\" w:header=\"851\" w:footer=\"992\" w:gutter=\"0\"/><w:cols w:space=\"425\"/><w:docGrid w:type=\"lines\" w:linePitch=\"360\"/></w:sectPr>","post":"</w:body></w:document>","rpr":{"TITLE":"<w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\" w:hint=\"eastAsia\"/><w:sz w:val=\"52\"/><w:szCs w:val=\"52\"/></w:rPr>","L1":"<w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\" w:hint=\"eastAsia\"/><w:sz w:val=\"44\"/><w:szCs w:val=\"44\"/></w:rPr>","L2":"<w:rPr><w:rFonts w:ascii=\"標楷體\" w:eastAsia=\"標楷體\" w:hAnsi=\"標楷體\" w:hint=\"eastAsia\"/><w:sz w:val=\"44\"/><w:szCs w:val=\"44\"/></w:rPr>"}};
function getCertTemplate(pin) { need_(pin, ['行政']); return CERT_TEMPLATE; }
function getCertificates(pin, caseId, kind) {
  need_(pin, ['行政']);
  if (!['個人', '選手', '班級'].includes(kind)) { throw new Error('獎狀種類不正確'); }
  const cf = conf_(); // caseId 可為單一編號、編號陣列，或空白（全部）
  const want = Array.isArray(caseId) ? caseId.map(String) : (String(caseId || '') ? [String(caseId)] : null);
  const records = certRows_().filter(o => o.kind === kind && (!want || want.includes(o.caseId)))
    .map(o => ({ no: o.fullNo, event: o.event, rank: o.rank, grade: o.grade, clsNum: o.clsNum, name: o.name, score: o.score, caseId: o.caseId }));
  return { records: records, title: cf['活動名稱'] || '', date: awardDate_(cf), layout: certLayout_(cf) };
}
function certLayout_(cf) {
  const n = k => { const x = Number(cf[k]); return isFinite(x) ? x : 0; };
  return { top: n('獎狀上邊界微調'), bottom: n('獎狀下邊界微調'), left: n('獎狀左邊界微調'), right: n('獎狀右邊界微調'),
    dy: n('列印頁垂直位移'), dx: n('列印頁水平位移') };
}
/* 版面測試用的三張獎狀：一般、長項目名稱、長英文姓名 */
function getCertSample(pin) {
  need_(pin, ['行政']); const cf = conf_();
  const no = cf['獎狀字號'] ? `${(cf['字號前綴'] || '') + cf['獎狀字號']}號` : '';
  const mk = (event, name, score, rank) => ({
    no: no,
    event: event,
    rank: rank,
    grade: '五',
    clsNum: '3',
    name: name,
    score: score,
    caseId: '測試'
  });
  return {
    records: [mk('五年級100公尺男生組', '王小明', '15"32', '1'),
      mk('五年級四百公尺混合接力暨趣味競賽男女混合組', '歐陽司馬大明', '1分15秒32', '2'),
      mk('五年級100公尺男生組', 'Alexander Christopher Wang', '15"32', '3')],
    title: cf['活動名稱'] || '', date: awardDate_(cf), layout: certLayout_(cf)
  };
}

/* ───────── 名冊管理（限行政） ───────── */
function rosterRows_() {
  return table_(SH.ROSTER).map(r => [0, 1, 2, 3, 4].map(i => String(r[i] || '').trim()))
    .filter(r => r[0] || r[1] || r[2]);
}
function getRoster(pin) {
  need_(pin, ['行政']);
  return { rows: rosterRows_() };
}
function rosterRow_(r, i) {
  try {
    if (!Array.isArray(r)) { throw new Error('資料格式不正確'); }
    const type = str_(r[0], 2, '類別', true);
    if (type !== '個人' && type !== '接力') { throw new Error('類別只能填「個人」或「接力」'); }
    const ev = str_(r[1], 30, '項目', true);
    if (!GRADE_NUM[ev.charAt(0)]) { throw new Error('項目名稱要以年級開頭'); }
    const cls = str_(r[2], 3, '班級', true);
    if (!/^[1-6]\d{2}$/.test(cls) || cls.charAt(0) !== GRADE_NUM[ev.charAt(0)]) { throw new Error('班級代碼要三碼且與項目年級相符'); }
    const no = str_(r[3], 2, '座號', false);
    if (no && !/^\d{1,2}$/.test(no)) { throw new Error('座號要填數字'); }
    const name = str_(r[4], 12, '姓名', false);
    return [type, ev, cls, no, name];
  } catch (e) { throw new Error(`第 ${i + 1} 列：${e.message}`); }
}
function importRoster(pin, rows, mode) {
  need_(pin, ['行政']);
  if (!Array.isArray(rows) || rows.length > 3000) { throw new Error('名冊資料不正確或超過 3000 列'); }
  if (mode !== 'replace' && mode !== 'append') { throw new Error('匯入方式不正確'); }
  const clean = rows.map(rosterRow_);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = sheet_(SH.ROSTER);
    if (mode === 'replace') {
      sh.clearContents();
      writeRows_(sh, 1, [ROSTER_H].concat(clean));
    } else if (clean.length) {
      if (sh.getLastRow() === 0) { writeRows_(sh, 1, [ROSTER_H]); }
      writeRows_(sh, sh.getLastRow() + 1, clean);
    }
    return { count: clean.length, rows: rosterRows_() };
  } finally { lock.releaseLock(); }
}

/* ───────── 系統設定（限行政） ───────── */
const CONF_EDIT = ['系統名稱', '學校名稱', '學年度', '活動名稱', '活動日期', '字號前綴', '獎狀字號', '編號方式', '逾時分鐘'].concat(MARGIN_KEYS, SHIFT_KEYS, ['頒獎日期', '頒獎日期格式', '檢錄員可更正', '檢錄通行碼']);
function getSettings(pin) {
  need_(pin, ['行政']); const c = conf_(); const o = {};
  CONF_EDIT.forEach(k => { o[k] = c[k]; });
  return { values: o };
}
function saveSettings(pin, input) {
  need_(pin, ['行政']); const given = input || {}; const cur = conf_();
  input = {}; // 沒有送來的項目維持原值，避免舊版頁面儲存時把新設定清空
  CONF_EDIT.forEach(k => { input[k] = given[k] === undefined ? cur[k] : given[k]; });
  const v = {
    '系統名稱': str_(input['系統名稱'], 30, '系統名稱', true),
    '學校名稱': str_(input['學校名稱'], 30, '學校名稱', true),
    '學年度': str_(input['學年度'], 3, '學年度', true),
    '活動名稱': str_(input['活動名稱'], 40, '活動名稱', true),
    '活動日期': str_(input['活動日期'], 10, '活動日期', false),
    '字號前綴': str_(input['字號前綴'], 20, '字號前綴', false),
    '獎狀字號': str_(input['獎狀字號'], 20, '獎狀字號', false),
    '編號方式': str_(input['編號方式'], 4, '編號方式', false) || '同一字號',
    '逾時分鐘': str_(input['逾時分鐘'], 3, '逾時分鐘', true)
  };
  MARGIN_KEYS.forEach(k => {
    const s = str_(input[k], 6, k, false) || '0';
    if (!/^-?\d{1,2}(\.\d)?$/.test(s) || Number(s) < -20 || Number(s) > 30) { throw new Error(`${k}請填 -20 到 30 之間的數字（單位公釐）`); }
    v[k] = String(Number(s));
  });
  v['頒獎日期'] = str_(input['頒獎日期'], 10, '頒獎日期', false);
  if (v['頒獎日期'] && !/^\d{4}-\d{2}-\d{2}$/.test(v['頒獎日期'])) { throw new Error('頒獎日期格式不正確'); }
  v['頒獎日期格式'] = str_(input['頒獎日期格式'], 4, '頒獎日期格式', false) || '中華民國';
  if (!DATE_MODES.includes(v['頒獎日期格式'])) { throw new Error('頒獎日期格式只能是「中華民國」或「西元」'); }
  v['檢錄員可更正'] = str_(input['檢錄員可更正'], 2, '檢錄員可更正', false) || '開啟';
  if (!['開啟', '關閉'].includes(v['檢錄員可更正'])) { throw new Error('檢錄員可更正只能是「開啟」或「關閉」'); }
  v['檢錄通行碼'] = str_(input['檢錄通行碼'], 20, '檢錄通行碼', false);
  if (/\s/.test(v['檢錄通行碼'])) { throw new Error('檢錄通行碼不能含空白'); }
  SHIFT_KEYS.forEach(k => {
    const s = str_(input[k], 6, k, false) || '0';
    if (!/^-?\d{1,2}(\.\d)?$/.test(s) || Math.abs(Number(s)) > 20) { throw new Error(`${k}請填 -20 到 20 之間的數字（單位公釐）`); }
    v[k] = String(Number(s));
  });
  if (!/^\d{2,3}$/.test(v['學年度'])) { throw new Error('學年度請填數字，例如 115'); }
  if (!NO_MODES.includes(v['編號方式'])) { throw new Error('編號方式只能是「同一字號」或「流水號」'); }
  if (v['活動日期'] && !/^\d{4}-\d{2}-\d{2}$/.test(v['活動日期'])) { throw new Error('活動日期格式不正確'); }
  if (!/^\d{1,3}$/.test(v['逾時分鐘']) || Number(v['逾時分鐘']) < 1 || Number(v['逾時分鐘']) > 240) { throw new Error('逾時分鐘請填 1 到 240'); }
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = sheet_(SH.CONF); const rows = table_(SH.CONF); const at = {};
    rows.forEach((r, i) => { at[r[0]] = i + 2; });
    let next = sh.getLastRow() + 1;
    CONF_EDIT.forEach(k => {
      if (at[k]) { writeRows_(sh, at[k], [[k, v[k]]]); } else { writeRows_(sh, next++, [[k, v[k]]]); }
    });
    CONF_CACHE_ = null;
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ───────── 比賽項目管理（限行政；項目由名冊決定） ───────── */
function evType_(type) { if (type !== '個人' && type !== '接力') { throw new Error('類別不正確'); } return type; }
function evName_(name) {
  const n = str_(name, 30, '項目名稱', true);
  if (!GRADE_NUM[n.charAt(0)]) { throw new Error('項目名稱要以年級開頭，例如「五年級100公尺男生組」'); }
  return n;
}
function writeRoster_(rows) { const sh = sheet_(SH.ROSTER); sh.clearContents(); writeRows_(sh, 1, [ROSTER_H].concat(rows)); }
function addEvent(pin, type, name, classes) {
  need_(pin, ['行政']); type = evType_(type); const n = evName_(name);
  if (!Array.isArray(classes) || classes.length < 1 || classes.length > 20) { throw new Error('請填寫參賽班級代碼'); }
  const seen = {}; const add = classes.map(c => {
    const cls = str_(c, 3, '班級', true);
    if (!/^[1-6]\d{2}$/.test(cls) || cls.charAt(0) !== GRADE_NUM[n.charAt(0)]) { throw new Error(`班級 ${cls} 與項目年級不符`); }
    if (seen[cls]) { throw new Error(`班級重複：${cls}`); } seen[cls] = true;
    return [type, n, cls, '', ''];
  });
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const rows = rosterRows_();
    if (rows.some(r => r[0] === type && r[1] === n)) { throw new Error(`已有同名項目：${n}`); }
    writeRoster_(rows.concat(add));
    return { ok: true };
  } finally { lock.releaseLock(); }
}
function renameEvent(pin, type, oldName, newName) {
  const w = need_(pin, ['行政']); type = evType_(type); const n = evName_(newName); const o = String(oldName || '');
  if (n === o) { return { ok: true, roster: 0, cases: 0 }; }
  if (n.charAt(0) !== o.charAt(0)) { throw new Error('改名不能更換年級，否則班級會對不上'); }
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const rows = rosterRows_(); let nr = 0, nc = 0;
    if (rows.some(r => r[0] === type && r[1] === n)) { throw new Error(`已有同名項目：${n}`); }
    rows.forEach(r => { if (r[0] === type && r[1] === o) { r[1] = n; nr++; } });
    if (!nr) { throw new Error(`找不到項目：${o}`); }
    writeRoster_(rows);
    const sh = sheet_(SH.CASE); const last = sh.getLastRow();
    if (last >= 2) {
      const all = sh.getRange(2, 1, last - 1, CASE_H.length).getDisplayValues();
      all.forEach((v, i) => {
        if (v[0] && v[COL['類別']] === type && v[COL['項目']] === o) {
          v[COL['項目']] = n; nc++;
          saveRow_({ sh: sh, row: i + 2, values: v }, w.name || '行政', `項目改名：${o} → ${n}`);
        }
      });
    }
    return { ok: true, roster: nr, cases: nc };
  } finally { lock.releaseLock(); }
}
function deleteEvent(pin, type, name) {
  need_(pin, ['行政']); type = evType_(type);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const rows = rosterRows_(); const keep = rows.filter(r => !(r[0] === type && r[1] === String(name)));
    if (keep.length === rows.length) { throw new Error(`找不到項目：${name}`); }
    writeRoster_(keep);
    return { ok: true, removed: rows.length - keep.length };
  } finally { lock.releaseLock(); }
}
function moveEvent(pin, type, name, dir) {
  need_(pin, ['行政']); type = evType_(type);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const rows = rosterRows_(); const order = []; const group = {};
    rows.forEach(r => { const k = `${r[0]}|${r[1]}`; if (!group[k]) { group[k] = []; order.push(k); } group[k].push(r); });
    const same = order.filter(k => k.startsWith(`${type}|`));
    const i = same.indexOf(`${type}|${name}`); const j = i + (dir === 'up' ? -1 : 1);
    if (i < 0) { throw new Error(`找不到項目：${name}`); }
    if (j < 0 || j >= same.length) { return { ok: true, moved: false }; }
    const a = order.indexOf(same[i]), b = order.indexOf(same[j]);
    order[a] = same[j]; order[b] = same[i];
    let out = []; order.forEach(k => { out = out.concat(group[k]); });
    writeRoster_(out);
    return { ok: true, moved: true };
  } finally { lock.releaseLock(); }
}

/* ───────── 項目的班級增減（限行政） ───────── */
function addEventClass(pin, type, name, cls) {
  need_(pin, ['行政']); type = evType_(type); const n = String(name || '');
  const c = str_(cls, 3, '班級代碼', true);
  if (!/^[1-6]\d{2}$/.test(c) || c.charAt(0) !== GRADE_NUM[n.charAt(0)]) { throw new Error(`班級 ${c} 與項目年級不符`); }
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const rows = rosterRows_(); let last = -1;
    rows.forEach((r, i) => {
      if (r[0] === type && r[1] === n) { last = i; if (r[2] === c) { throw new Error(`這個項目已經有 ${c} 班`); } }
    });
    if (last < 0) { throw new Error(`找不到項目：${n}`); }
    rows.splice(last + 1, 0, [type, n, c, '', '']); // 接在該項目最後一列之後，維持項目順序
    writeRoster_(rows);
    return { ok: true };
  } finally { lock.releaseLock(); }
}
function deleteEventClass(pin, type, name, cls) {
  need_(pin, ['行政']); type = evType_(type); const n = String(name || ''); const c = String(cls || '');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const rows = rosterRows_();
    const mine = rows.filter(r => r[0] === type && r[1] === n);
    const keep = rows.filter(r => !(r[0] === type && r[1] === n && r[2] === c));
    const removed = rows.length - keep.length;
    if (!removed) { throw new Error(`這個項目沒有 ${c} 班`); }
    if (removed === mine.length) { throw new Error('這是此項目最後一個班級。要整個移除請改用「刪除」項目'); }
    writeRoster_(keep);
    return { ok: true, removed: removed };
  } finally { lock.releaseLock(); }
}

/* ───────── 刪除案件（限行政）：刪除前一律先抄到「案件備份」工作表 ───────── */
function backupCases_(rows, who, how) {
  if (!rows.length) { return; }
  const ss = SpreadsheetApp.getActiveSpreadsheet(); let sh = ss.getSheetByName(SH.BACKUP);
  if (!sh) { sh = ss.insertSheet(SH.BACKUP); }
  if (sh.getLastRow() === 0) { writeRows_(sh, 1, [['刪除時間', '刪除者', '方式'].concat(CASE_H)]); sh.setFrozenRows(1); }
  const t = now_();
  writeRows_(sh, sh.getLastRow() + 1, rows.map(r => [t, who, how].concat(r.slice(0, CASE_H.length))));
}
function rewriteCases_(keep, total) { // 以「清除內容再寫回」的方式刪列，避免刪到凍結列以外的最後一列時出錯
  const sh = sheet_(SH.CASE);
  if (total > 0) { sh.getRange(2, 1, total, CASE_H.length).clearContent(); }
  if (keep.length) { writeRows_(sh, 2, keep); }
}
function caseRows_() {
  const sh = sheet_(SH.CASE); const n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, CASE_H.length).getDisplayValues();
}
function deleteCase(pin, id) {
  const w = need_(pin, ['行政']); const target = String(id || '');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const all = caseRows_(); const gone = all.filter(r => r[0] === target);
    if (!target || !gone.length) { throw new Error(`找不到案件 ${target}`); }
    backupCases_(gone, w.name || '行政', '單筆刪除');
    rewriteCases_(all.filter(r => r[0] && r[0] !== target), all.length);
    return { ok: true, removed: gone.length };
  } finally { lock.releaseLock(); }
}
function clearCases(pin, expected) {
  const w = need_(pin, ['行政']);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const all = caseRows_(); const real = all.filter(r => r[0]);
    if (Number(expected) !== real.length) { throw new Error(`案件數量已變動（目前 ${real.length} 筆），請重新下載資料後再清空`); }
    backupCases_(real, w.name || '行政', '全部清空');
    rewriteCases_([], all.length);
    setConf_('案件流水號', '0'); // 清空後編號從 R001 重新開始
    setConf_('播報順序', '');
    return { ok: true, removed: real.length };
  } finally { lock.releaseLock(); }
}

/* ───────── 匯入案件資料（限行政）：格式與下載的 CSV 相同 ───────── */
function pad2_(v) { return (`0${Number(v)}`).slice(-2); }
function normDate_(s) { const m = /^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/.exec(String(s || '').trim()); return m ? `${m[1]}-${pad2_(m[2])}-${pad2_(m[3])}` : null; }
function normTime_(s) { const m = /^(\d{1,2}):(\d{2})(?::\d{2})?$/.exec(String(s || '').trim()); return m && Number(m[1]) < 24 && Number(m[2]) < 60 ? `${pad2_(m[1])}:${m[2]}` : null; }
function normStamp_(s, label) { // 空白回傳空字串；格式不對則丟出錯誤
  const v = String(s || '').trim(); if (!v) { return ''; }
  const m = /^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(v);
  if (!m) { throw new Error(`${label}格式不正確（應為 2026-11-21 09:30:00）`); }
  return `${m[1]}-${pad2_(m[2])}-${pad2_(m[3])} ${pad2_(m[4])}:${m[5]}:${pad2_(m[6] || 0)}`;
}
function importCases(pin, items, dryRun) {
  const w = need_(pin, ['行政']);
  if (!Array.isArray(items) || !items.length) { throw new Error('檔案裡沒有可匯入的資料'); }
  if (items.length > 500) { throw new Error('一次最多匯入 500 筆'); }
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ev = events_(); const all = caseRows_(); const existing = all.filter(r => r[0]);
    const used = {}; let max = Number(conf_()['案件流水號']) || 0;
    existing.forEach(r => { used[r[0]] = true; const m = /^R(\d+)$/.exec(r[0]); if (m) { max = Math.max(max, Number(m[1])); } });
    const errors = []; const rows = []; const stamp = now_();
    items.forEach((it, i) => {
      try {
        it = it || {};
        const t0 = String(it.type || '').trim();
        const type = ['個人', '個人單項', '個人單項競賽'].includes(t0) ? '個人' : ['接力', '大隊接力'].includes(t0) ? '接力' : null;
        if (!type) { throw new Error('類別要填「個人單項」或「大隊接力」'); }
        const evName = str_(it.event, 30, '項目', true);
        const e = (type === '接力' ? ev.relay : ev.individual).filter(x => x.name === evName)[0];
        if (!e) { throw new Error(`項目「${evName}」不在目前的比賽項目裡`); }
        const date = normDate_(it.date); const time = normTime_(it.time);
        if (!date) { throw new Error('比賽日期格式不正確（應為 2026-11-21）'); }
        if (!time) { throw new Error('比賽時間格式不正確（應為 09:30）'); }
        const recorder = str_(it.recorder, 20, '檢錄員', true);
        const places = places_(type, e, it.places);
        const id = str_(it.id, 8, '案件編號', false);
        if (id) {
          if (!/^R\d{3,6}$/.test(id)) { throw new Error('案件編號格式不正確（例如 R001，或留空由系統編號）'); }
          if (used[id]) { throw new Error(`案件編號 ${id} 已存在或在檔案中重複`); }
          used[id] = true;
        }
        const status = str_(it.status, 5, '狀態', false) || '待處理';
        if (!STATUS.includes(status)) { throw new Error('狀態只能是待處理、處理中、已完成'); }
        const audit = str_(it.audit, 5, '稽核', false);
        if (!['', '已核可', '退回補正'].includes(audit)) { throw new Error('稽核只能是空白、已核可、退回補正'); }
        const log = (Array.isArray(it.log) ? it.log : []).slice(0, 200).map(l => ({
          t: String((l || {}).t || '').slice(0, 19),
          by: String((l || {}).by || '').slice(0, 20),
          what: String((l || {}).what || '').slice(0, 600)
        })).filter(l => l.what);
        log.push({ t: stamp, by: w.name || '行政', what: '由檔案匯入' });
        const row = CASE_H.map(() => '');
        row[COL['案件編號']] = id; row[COL['送出時間']] = normStamp_(it.at, '送出時間') || stamp;
        row[COL['類別']] = type; row[COL['項目']] = evName; row[COL['比賽日期']] = date; row[COL['比賽時間']] = time;
        row[COL['檢錄員']] = recorder; row[COL['急件']] = String(it.urgent || '').trim() === '是' ? '是' : '否';
        row[COL['成績摘要']] = summary_(places); row[COL['狀態']] = status;
        row[COL['負責人']] = str_(it.assignee, 20, '負責人', false); row[COL['處理結果']] = str_(it.result, 300, '處理結果', false);
        row[COL['完成時間']] = normStamp_(it.doneAt, '完成時間'); row[COL['播報時間']] = normStamp_(it.announcedAt, '播報時間');
        row[COL['稽核']] = audit; row[COL['稽核意見']] = str_(it.auditNote, 300, '稽核意見', false); row[COL['稽核時間']] = normStamp_(it.auditAt, '稽核時間');
        row[COL['成績資料']] = JSON.stringify(places); row[COL['處理紀錄']] = JSON.stringify(log);
        row[COL['列印時間']] = normStamp_(it.printedAt, '列印時間');
        rows.push(row);
      } catch (err) { errors.push(`第 ${i + 2} 列：${err.message}`); }
    });
    if (errors.length || dryRun) { return { ok: !errors.length, count: rows.length, errorCount: errors.length, errors: errors.slice(0, 30) }; }
    rows.forEach(r => { const m = /^R(\d+)$/.exec(r[0]); if (m) { max = Math.max(max, Number(m[1])); } });
    rows.forEach(r => { if (!r[0]) { max++; r[0] = `R${max < 1000 ? (`000${max}`).slice(-3) : String(max)}`; } });
    rewriteCases_(existing.concat(rows), all.length);
    setConf_('案件流水號', String(max));
    return { ok: true, count: rows.length, errorCount: 0, errors: [] };
  } finally { lock.releaseLock(); }
}
