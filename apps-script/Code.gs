/**
 * ห้องเรียน Climate Change 282221 — ระบบหลังบ้าน (Google Apps Script)
 *
 * เก็บข้อมูลใน Google Sheet "ฐานข้อมูลห้องเรียน_282221" และเก็บไฟล์ลงโฟลเดอร์
 * climate_change_topic 282121 ใน Google Drive ของอาจารย์
 *
 * ตั้งค่าครั้งแรก (ทำครั้งเดียว):
 *   1) แก้ข้อความในบรรทัด const key = '...' ของฟังก์ชัน setup() ให้เป็นรหัสผ่านโหมดอาจารย์ (แก้บรรทัดเดียว)
 *   2) เลือกฟังก์ชัน setup แล้วกด Run (อนุญาตสิทธิ์ Drive/Sheets เมื่อถูกถาม)
 *   3) ลบรหัสที่พิมพ์ไว้ใน setup() ออกได้เลย รหัสถูกเก็บใน Script Properties แล้ว
 *   4) Deploy > New deployment > Web app: Execute as = Me, Who has access = Anyone
 */

const SHEET_ID = '1Zxv9mbEBykVpRS_tKhlHL3BfilXoW4qLCxq-lwXR2sI';
const FOLDERS = {
  slides:   '1toIBa_Rn78lykCTOzC2TQofSnVbVlVcg', // 01_สไลด์บรรยาย
  work:     '1ZHAyiHql9IcxOZGMtvHpFBJC24v0AlCd', // 02_งานกลุ่ม
  records:  '1DNQtJXZvMKL8G4m_chqWMAa0zNU2zVbz', // 03_บันทึกการเรียน
  homework: '1EG1VEBnZaPY2SVStedW9hILOKvVGTpcT'  // 04_การบ้าน
};
const SCEN = [
  'เปิดเครื่องปรับอากาศในหอพักช่วงเมษายน', 'ไฟฟ้าในอาคารเรียนและห้องแล็บของคณะ',
  'ชาร์จและเปิดอุปกรณ์ไฟฟ้าทิ้งไว้ในหอพัก', 'เศษอาหารจากโรงอาหารส่งฝังกลบ',
  'ขยะในหอพักที่ไม่แยกประเภท', 'ร้านอาหารรอบมหาวิทยาลัยใช้ก๊าซหุงต้ม',
  'พิมพ์รายงานและเอกสารลงกระดาษ', 'เครื่องสูบน้ำดีเซลช่วงภัยแล้ง', 'เผาตอซังข้าวโพดก่อนฤดูปลูก'
];
const EXITQ = [
  'อธิบายปรากฏการณ์เรือนกระจก และยกก๊าซเรือนกระจกหลัก 2 ชนิด พร้อมแหล่งที่มา',
  'ภาคใดปล่อยก๊าซเรือนกระจกมากที่สุดของไทย ราวร้อยละเท่าไร และสถานการณ์ของกลุ่มนับอยู่ในภาคใด',
  'มาตรการที่กลุ่มเสนอเป็น Mitigation หรือ Adaptation เพราะเหตุใด และจะทำได้จริงอย่างไร'
];
const RATING = { pass: 'ผ่าน', partial: 'ผ่านบางส่วน', fail: 'ยังไม่ผ่าน' };
const MAX_IMAGE_BYTES = 8 * 1024 * 1024;

/* ---------- setup ---------- */
function setup() {
  const key = 'ใส่รหัสอาจารย์ที่นี่';
  if (!key || key.indexOf('ใส่รหัส') === 0) throw new Error('แก้รหัสอาจารย์ในฟังก์ชัน setup() ก่อนกด Run');
  PropertiesService.getScriptProperties().setProperty('TEACHER_KEY', key);
  kvSheet_();
  DriveApp.getFolderById(FOLDERS.work); // ขอสิทธิ์ Drive ตั้งแต่ตอนนี้
  Logger.log('ตั้งค่าเรียบร้อย นักเรียน ' + roster_().length + ' คน');
}

/* ---------- entry points ---------- */
function doGet(e) { return handle_((e && e.parameter) || {}); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'คำขอไม่ถูกต้อง' }); }
  return handle_(p);
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function handle_(p) {
  try {
    const teacher = isTeacher_(p.key);
    switch (p.action) {
      case 'state':          return out_({ ok: true, teacher: teacher, state: state_(teacher) });
      case 'checkKey':       return out_({ ok: teacher });
      case 'saveGroups':     need_(teacher); return out_(saveGroups_(p));
      case 'unlockGroups':   need_(teacher); return out_(withLock_(() => { const g = kvGet_('groups') || {}; g.lockedAt = null; kvSet_('groups', g); return { ok: true }; }));
      case 'setPresent':     need_(teacher); return out_(setGroupDoc_('present', p.group, p.data));
      case 'setExit':        need_(teacher); return out_(setGroupDoc_('exit', p.group, p.data));
      case 'setSettings':    need_(teacher); return out_(withLock_(() => { kvSet_('settings', Object.assign(kvGet_('settings') || {}, p.data || {})); return { ok: true }; }));
      case 'uploadSlides':   need_(teacher); return out_(uploadSlides_(p));
      case 'setSlidesLink':  need_(teacher); return out_(setSlidesLink_(p));
      case 'clearSlides':    need_(teacher); return out_(withLock_(() => { kvSet_('slides', null); return { ok: true }; }));
      case 'uploadWork':     return out_(uploadWork_(p));
      case 'saveWorkText':   return out_(saveWorkText_(p));
      case 'submitHomework': return out_(submitHomework_(p));
      case 'checkHomework':  return out_(checkHomework_(p));
      case 'homeworkList':   need_(teacher); return out_({ ok: true, list: homeworkList_() });
      case 'resetAll':       need_(teacher); return out_(resetAll_(p.parts || []));
      default:               return out_({ ok: false, error: 'ไม่รู้จักคำสั่ง' });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ---------- auth ---------- */
function isTeacher_(key) {
  const k = PropertiesService.getScriptProperties().getProperty('TEACHER_KEY');
  return !!k && !!key && String(key) === k;
}
function need_(t) { if (!t) throw new Error('ต้องใช้รหัสอาจารย์'); }

/* ---------- storage: roster + key/value sheet ---------- */
function ss_() { return SpreadsheetApp.openById(SHEET_ID); }
function roster_() {
  const c = CacheService.getScriptCache().get('roster');
  if (c) return JSON.parse(c);
  const sh = ss_().getSheets()[0];
  const rows = sh.getDataRange().getValues().slice(1).filter(r => r[0]);
  const list = rows.map(r => ({ id: String(r[0]).trim(), name: String(r[1]).trim(), prog: String(r[2] || '').trim() }));
  CacheService.getScriptCache().put('roster', JSON.stringify(list), 600);
  return list;
}
function kvSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName('_data');
  if (!sh) { sh = ss.insertSheet('_data'); sh.getRange(1, 1, 1, 3).setValues([['key', 'json', 'updated']]); sh.hideSheet(); }
  return sh;
}
function kvAll_() {
  const v = kvSheet_().getDataRange().getValues();
  const m = {};
  for (let i = 1; i < v.length; i++) if (v[i][0]) { try { m[v[i][0]] = JSON.parse(v[i][1]); } catch (e) {} }
  return m;
}
function kvGet_(key) { return kvAll_()[key]; }
function kvSet_(key, val) {
  const sh = kvSheet_();
  const keys = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), 1).getValues().map(r => r[0]);
  const row = keys.indexOf(key) + 1;
  const vals = [[key, JSON.stringify(val), new Date()]];
  if (row > 0) sh.getRange(row, 1, 1, 3).setValues(vals);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, 3).setValues(vals);
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------- names ---------- */
function stripTitle_(n) { return String(n || '').replace(/^(นางสาว|นาย|นาง)/, '').trim(); }
function shortName_(n) {
  const p = stripTitle_(n).split(/\s+/);
  return p.length > 1 ? p[0] + ' ' + p[1].charAt(0) + '.' : p[0];
}
function person_(id, teacher, byId) {
  if (!id) return null;
  const r = byId[id];
  if (!r) return { id: teacher ? id : undefined, name: teacher ? id : '?' };
  return teacher ? { id: r.id, name: r.name, prog: r.prog } : { name: shortName_(r.name) };
}

/* ---------- public / teacher state ---------- */
function state_(teacher) {
  const kv = kvAll_();
  const ro = roster_();
  const byId = {}; ro.forEach(r => byId[r.id] = r);
  const P = id => person_(id, teacher, byId);
  const g = kv.groups || null;
  const groups = g && g.groups ? g.groups.map(m => m.map(P)) : null;
  const work = {}, present = {}, exit = {};
  for (let i = 1; i <= 9; i++) {
    const w = kv['work/g' + i]; if (w) work[i] = w;
    const pr = kv['present/g' + i];
    if (pr) present[i] = { presenter: P(pr.presenter), asker: P(pr.asker), done: !!pr.done };
    const ex = kv['exit/g' + i];
    if (ex) exit[i] = {
      done: !!ex.done,
      q: (ex.q || []).map(a => ({ answerer: P(a.answerer), rating: a.rating || null, note: teacher ? (a.note || '') : undefined, tries: (a.tries || []).length }))
    };
  }
  const st = {
    groups: groups, lockedAt: g ? g.lockedAt || null : null,
    work: work, present: present, exit: exit,
    slides: kv.slides || null, settings: kv.settings || {},
    rosterCount: ro.length
  };
  if (teacher) {
    st.roster = ro;
    st.groupIds = g && g.groups ? g.groups : null;
    st.exitRaw = {}; for (let i = 1; i <= 9; i++) if (kv['exit/g' + i]) st.exitRaw[i] = kv['exit/g' + i];
    st.homework = homeworkSummary_(kv);
  }
  return st;
}

/* ---------- groups ---------- */
function saveGroups_(p) {
  const ids = {}; roster_().forEach(r => ids[r.id] = 1);
  const groups = p.groups;
  if (!Array.isArray(groups) || groups.length !== 9) throw new Error('ต้องมี 9 กลุ่ม');
  const flat = [].concat.apply([], groups);
  if (flat.some(id => !ids[id])) throw new Error('มีรหัสนิสิตที่ไม่อยู่ในรายชื่อ');
  return withLock_(() => {
    kvSet_('groups', { groups: groups, lockedAt: Date.now() });
    writeReports_();
    return { ok: true };
  });
}

/* ---------- present / exit ---------- */
function setGroupDoc_(kind, group, data) {
  group = Number(group);
  if (!(group >= 1 && group <= 9)) throw new Error('หมายเลขกลุ่มไม่ถูกต้อง');
  return withLock_(() => {
    const key = kind + '/g' + group;
    const cur = kvGet_(key) || {};
    kvSet_(key, Object.assign(cur, data || {}, { at: Date.now() }));
    writeReports_();
    return { ok: true };
  });
}

/* ---------- files ---------- */
function blobFrom_(b64, mime, name) {
  if (!b64) throw new Error('ไม่มีไฟล์');
  const bytes = Utilities.base64Decode(String(b64).replace(/^data:[^,]+,/, ''));
  if (bytes.length > MAX_IMAGE_BYTES * 4) throw new Error('ไฟล์ใหญ่เกินไป');
  return Utilities.newBlob(bytes, mime, name);
}
function isImage_(mime) { return /^image\/(jpeg|png|webp)$/.test(String(mime)); }
function stamp_() { return Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd_HHmm'); }
function share_(file) { try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {} }

function uploadSlides_(p) {
  const okType = /pdf|presentation/.test(String(p.mime));
  if (!okType) throw new Error('รองรับเฉพาะ PDF หรือ PowerPoint');
  const f = DriveApp.getFolderById(FOLDERS.slides).createFile(blobFrom_(p.data, p.mime, p.name || 'slides.pdf'));
  share_(f);
  return withLock_(() => { const s = { fileId: f.getId(), name: f.getName(), at: Date.now() }; kvSet_('slides', s); return { ok: true, slides: s }; });
}
function setSlidesLink_(p) {
  const m = String(p.link || '').match(/[-\w]{25,}/);
  if (!m) throw new Error('ลิงก์ Google Drive ไม่ถูกต้อง');
  const f = DriveApp.getFileById(m[0]);
  share_(f);
  return withLock_(() => { const s = { fileId: f.getId(), name: f.getName(), at: Date.now() }; kvSet_('slides', s); return { ok: true, slides: s }; });
}

function uploadWork_(p) {
  const g = Number(p.group);
  if (!(g >= 1 && g <= 9)) throw new Error('หมายเลขกลุ่มไม่ถูกต้อง');
  if (!isImage_(p.mime)) throw new Error('รองรับเฉพาะรูป JPG หรือ PNG');
  const kind = p.kind === 'slide' ? 'slide' : 'photo';
  const label = kind === 'slide' ? 'ภาพนำเสนอ' : 'แผ่นงาน';
  const ext = p.mime === 'image/png' ? 'png' : 'jpg';
  const f = DriveApp.getFolderById(FOLDERS.work).createFile(blobFrom_(p.data, p.mime, 'กลุ่ม' + g + '_' + label + '_' + stamp_() + '.' + ext));
  share_(f); // ให้ภาพแสดงบนจอนำเสนอได้
  return withLock_(() => {
    const key = 'work/g' + g;
    const cur = kvGet_(key) || {};
    cur[kind] = { fileId: f.getId(), at: Date.now(), source: p.source || (kind === 'slide' ? 'upload' : 'photo') };
    if (p.ex) cur.ex = p.ex;
    cur.at = Date.now();
    kvSet_(key, cur);
    return { ok: true, fileId: f.getId() };
  });
}
function saveWorkText_(p) {
  const g = Number(p.group);
  if (!(g >= 1 && g <= 9)) throw new Error('หมายเลขกลุ่มไม่ถูกต้อง');
  return withLock_(() => { const key = 'work/g' + g; const cur = kvGet_(key) || {}; cur.ex = p.ex || null; cur.at = Date.now(); kvSet_(key, cur); return { ok: true }; });
}

/* ---------- homework ---------- */
function findStudent_(sid) {
  sid = String(sid || '').trim();
  const r = roster_().filter(x => x.id === sid)[0];
  if (!r) throw new Error('ไม่พบรหัสนิสิตนี้ในรายวิชา ตรวจรหัสอีกครั้ง');
  return r;
}
function submitHomework_(p) {
  const r = findStudent_(p.studentId);
  const imgs = (p.images || []).slice(0, 3);
  if (!imgs.length) throw new Error('ยังไม่ได้เลือกรูป');
  const folder = DriveApp.getFolderById(FOLDERS.homework);
  const ids = imgs.map((im, k) => {
    if (!isImage_(im.mime)) throw new Error('รองรับเฉพาะรูป JPG หรือ PNG');
    return folder.createFile(blobFrom_(im.data, im.mime, 'HW_' + r.id + '_' + stripTitle_(r.name) + '_' + stamp_() + '_หน้า' + (k + 1) + '.jpg')).getId();
  });
  withLock_(() => { kvSet_('hw/' + r.id, { files: ids, at: Date.now(), by: p.by === 'teacher' && isTeacher_(p.key) ? 'teacher' : 'student' }); writeReports_(); });
  return { ok: true, name: shortName_(r.name), pages: ids.length };
}
function checkHomework_(p) {
  const r = findStudent_(p.studentId);
  const h = kvGet_('hw/' + r.id);
  return { ok: true, name: shortName_(r.name), submitted: !!h, at: h ? h.at : null, pages: h ? h.files.length : 0 };
}
function homeworkSummary_(kv) {
  const out = {};
  Object.keys(kv).forEach(k => { if (k.indexOf('hw/') === 0) out[k.slice(3)] = kv[k]; });
  return out;
}
/* ---------- reset (backup first; Drive files are kept) ---------- */
const RESET_KEYS = { groups: ['groups'], work: ['work/'], present: ['present/'], exit: ['exit/'], homework: ['hw/'], slides: ['slides'], settings: ['settings'] };
function resetAll_(parts) {
  const want = [];
  parts.forEach(k => (RESET_KEYS[k] || []).forEach(x => want.push(x)));
  if (!want.length) throw new Error('ยังไม่ได้เลือกสิ่งที่จะรีเซ็ต');
  return withLock_(() => {
    const ss = ss_(); const sh = kvSheet_();
    const bk = sh.copyTo(ss); bk.setName('_backup_' + stamp_() + '_' + Utilities.getUuid().slice(0, 4)); bk.hideSheet();
    const v = sh.getDataRange().getValues();
    const keep = [v[0]]; let removed = 0;
    for (let i = 1; i < v.length; i++) {
      const k = String(v[i][0] || '');
      if (!k) continue;
      const hit = want.some(x => x.slice(-1) === '/' ? k.indexOf(x) === 0 : k === x);
      if (hit) removed++; else keep.push(v[i]);
    }
    sh.clearContents();
    sh.getRange(1, 1, keep.length, 3).setValues(keep);
    writeReports_();
    return { ok: true, removed: removed, backup: bk.getName() };
  });
}
function homeworkList_() {
  const kv = kvAll_(); const hw = homeworkSummary_(kv);
  return roster_().map(r => ({ id: r.id, name: r.name, hw: hw[r.id] || null }));
}

/* ---------- readable report tabs in the same Sheet ---------- */
function tab_(name, header, rows) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clearContents();
  const all = [header].concat(rows);
  sh.getRange(1, 1, all.length, header.length).setValues(all.map(r => header.map((_, i) => r[i] === undefined || r[i] === null ? '' : r[i])));
  sh.setFrozenRows(1);
}
function fmt_(ts) { return ts ? Utilities.formatDate(new Date(ts), 'Asia/Bangkok', 'd MMM yyyy HH:mm') : ''; }
function writeReports_() {
  const kv = kvAll_(); const ro = roster_();
  const byId = {}; ro.forEach(r => byId[r.id] = r);
  const nm = id => id ? (byId[id] ? byId[id].name : id) : '';
  const gs = kv.groups && kv.groups.groups || [];
  const gOf = {}; gs.forEach((m, i) => m.forEach(id => gOf[id] = i + 1));
  tab_('กลุ่ม', ['กลุ่ม', 'สถานการณ์', 'รหัสนิสิต', 'ชื่อ-สกุล', 'หลักสูตร'],
    [].concat.apply([], gs.map((m, i) => m.map(id => [i + 1, SCEN[i], id, nm(id), byId[id] ? byId[id].prog : '']))));
  const pres = [];
  for (let g = 1; g <= 9; g++) {
    const p = kv['present/g' + g] || {}; const ag = g % 9 + 1;
    pres.push([g, SCEN[g - 1], nm(p.presenter), ag, ['A', 'B', 'C'][(ag - 1) % 3], nm(p.asker), p.done ? 'นำเสนอแล้ว' : '']);
  }
  tab_('การนำเสนอ', ['กลุ่ม', 'สถานการณ์', 'ผู้นำเสนอ', 'กลุ่มผู้ถาม', 'แบบคำถาม', 'ผู้ถาม', 'สถานะ'], pres);
  const ex = [];
  for (let g = 1; g <= 9; g++) {
    const e = kv['exit/g' + g]; if (!e) continue;
    (e.q || []).forEach((a, i) => (a.tries || []).forEach((t, k) => ex.push([g, i + 1, EXITQ[i], k + 1, t.who, nm(t.who), RATING[t.rating] || '', k === a.tries.length - 1 ? a.note || '' : '', fmt_(t.at), e.done ? 'ออกได้' : ''])));
  }
  tab_('บัตรออก', ['กลุ่ม', 'ข้อ', 'คำถาม', 'ครั้งที่', 'รหัสผู้ตอบ', 'ผู้ตอบ', 'ผล', 'บันทึก', 'เวลา', 'กลุ่ม'], ex);
  const hw = homeworkSummary_(kv);
  tab_('การบ้าน', ['รหัสนิสิต', 'ชื่อ-สกุล', 'กลุ่ม', 'สถานะ', 'ส่งเมื่อ', 'จำนวนหน้า', 'ลิงก์ไฟล์'],
    ro.map(r => { const h = hw[r.id]; return [r.id, r.name, gOf[r.id] || '', h ? 'ส่งแล้ว' : 'ยังไม่ส่ง', h ? fmt_(h.at) : '', h ? h.files.length : '', h ? h.files.map(id => 'https://drive.google.com/file/d/' + id + '/view').join(' ') : '']; }));
}
