/**
 * Mala Order - Backend v11 (Google Apps Script + Google Sheets)
 * โครงเดียวกับ TOTU Concept: หน้าเว็บ Index.html เรียกสคริปต์นี้ผ่าน Web App URL
 *
 * ติดตั้งครั้งแรก
 *  1) สร้าง Google Sheet ใหม่ > ส่วนขยาย > Apps Script > วางโค้ดนี้ทับทั้งไฟล์ > Ctrl+S
 *  2) เลือกฟังก์ชัน setup ด้านบน > Run (อนุญาตสิทธิ์) : สร้างชีท Employees / Tables / Menu / Sessions / Orders พร้อมเมนูตัวอย่าง
 *  3) Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone > Deploy
 *  4) คัดลอก URL ที่ลงท้าย /exec ไปใส่ GAS_URL ใน index.html
 * อัปเดตครั้งถัดไป: วางโค้ดทับ > Deploy > Manage deployments > ✏️ > New version
 *
 * ชีท
 *  Employees : name, pin, role            (พนักงานที่เข้าหน้าโต๊ะ / ครัว / เสิร์ฟได้)
 *  Menu      : id, category, code, name, price, spicy, available, sort, image   (แก้เมนูที่นี่)
 *              available = FALSE คือของหมด: ลูกค้ายังเห็นเมนูแต่ขึ้นว่า "หมด" และสั่งไม่ได้ (พนักงานกดสลับได้จากแท็บ เมนู)
 *                          จะเลิกขายถาวรให้ลบแถวนั้นออก
 *              image = รูปอาหาร ใส่ได้ 2 ทาง: กด "ใส่รูป" ในแท็บ เมนู ของหน้าเว็บ (เฉพาะ role = admin)
 *              หรือวางลิงก์รูปในชีทเอง (ลิงก์ Google Drive ที่แชร์แบบ "ทุกคนที่มีลิงก์" หรือ https:// อื่น)
 *              รูปที่อัปโหลดเก็บในโฟลเดอร์ Drive ชื่อ "Mala Menu Images" 
 *  Tables    : name   ชื่อโต๊ะทั้งหมด 1 แถวต่อ 1 โต๊ะ (เช่น 1, 2, A1, B18) แต่ละโต๊ะมี QR ประจำโต๊ะใบเดียวใช้ตลอด
 *  Settings  : key, value   ตั้งค่าร้าน
 *              promptpay_id   = เบอร์มือถือ (10 หลัก) หรือเลขบัตรประชาชน / เลขผู้เสียภาษี (13 หลัก) ที่ผูกพร้อมเพย์ไว้ ใช้สร้าง QR รับเงิน
 *                               หรือเลขอ้างอิง 15 หลักใต้ QR รับเงินของแอปธนาคาร (พิมพ์ ' นำหน้าถ้าเลขขึ้นต้นด้วย 0)
 *              promptpay_note = ข้อความบอกบัญชี เช่น กสิกรไทย xxx-x-x7663-x (ไม่บังคับ)
 *              promptpay_name = ชื่อบัญชีที่จะแสดงให้ลูกค้าตรวจ   (เว้น promptpay_id ว่าง = ไม่แสดงปุ่ม QR รับเงิน)
 *  Calls     : ลูกค้ากดเรียก 1 ครั้ง = 1 แถว  type: staff (เรียกพนักงาน) / bill (ขอเช็คบิล)  doneAt, doneBy = พนักงานรับทราบแล้ว
 *  Sessions  : บิล 1 ใบ = 1 แถว เปิดเองเมื่อโต๊ะว่างแล้วมีคนสั่งครั้งแรก ปิดเมื่อพนักงานกดคิดเงิน
 *              payMethod = cash / transfer (none = ปิดโต๊ะที่ไม่มีออเดอร์)  closedBy = พนักงานที่กดปิดบิล  total = ยอดที่เก็บ
 *              slip = ลิงก์รูปสลิปโอนเงิน เก็บในโฟลเดอร์ Drive "Mala Payment Slips" แบบส่วนตัว (เปิดได้เฉพาะเจ้าของบัญชี Google)
 *  Orders    : การกดสั่ง 1 รอบ = 1 แถว  status: new (เข้าครัว) > cooking (กำลังย่าง) > ready (รอเสิร์ฟ) > served (เสิร์ฟครบ)
 *
 * ความปลอดภัย
 *  - ลูกค้าทำได้แค่ดูเมนู สั่งของ และดูบิลปัจจุบันของโต๊ะที่สแกน ราคาคิดจากชีท Menu ฝั่งเซิร์ฟเวอร์ (แก้ราคาจากมือถือไม่ได้)
 *  - QR ประจำโต๊ะใช้ได้ตลอด ใครมีรูป QR ก็สั่งเข้าโต๊ะนั้นได้ พนักงานจึงควรดูออเดอร์แปลก ๆ ก่อนทำ
 *  - เปิดโต๊ะ เปลี่ยนสถานะ คิดเงิน ต้องล็อกอินด้วยชื่อ + PIN พนักงาน
 */
const SERVER_VERSION = 14;  // v14: ตรวจสลิปอัตโนมัติด้วย Thunder (API key เก็บใน Script Properties ไม่ส่งไปหน้าเว็บ)
const THUNDER_URL = 'https://api.thunder.in.th/v2/verify/bank';
const SLIP_MAX_AGE_MIN = 180;   // สลิปเก่ากว่านี้ให้เตือน
// v13: v13: บันทึกเวลาเริ่มย่าง / ย่างเสร็จ / เสิร์ฟ และรายงานตามช่วงเวลา (ลูกค้าเข้าช่วงไหน ครัวย่างได้เท่าไร)
// v12: v12: ครัวมีขั้น "กำลังย่าง" (status cooking) และเลขออเดอร์ (no)
// v11: v11: เพิ่ม / แก้ไข / ลบเมนูจากหน้าเว็บ (เฉพาะ role = admin)
// v10: ลูกค้ากดเรียกพนักงาน / ขอเช็คบิล (ชีท Calls) แจ้งเตือนบนหน้าพนักงาน
// v9: ของหมด (available = FALSE แสดงว่าหมด สั่งไม่ได้ พนักงานกดจากแท็บ เมนู)
// v8: เก็บรูปสลิปโอนเงินตอนปิดบิล (โฟลเดอร์ Drive ส่วนตัว)
const SLIP_FOLDER = 'Mala Payment Slips';
// v7: ปิดบิลต้องเลือกวิธีจ่าย โอนต้องยืนยันว่าเห็นยอดเข้าบัญชี บันทึกคนปิดบิล
// v6: QR รับเงินพร้อมเพย์ (ชีท Settings)
// v5: สรุปยอดขาย (แท็บ ยอดขาย เฉพาะ role = admin)
const TIMEZONE = 'Asia/Bangkok';   // ใช้ตัดวันของยอดขาย
// v4: ส่งลำดับคิวในครัวให้หน้าลูกค้า
// v3: QR ประจำโต๊ะ (ชีท Tables) ลูกค้าสแกนแล้วสั่งได้เลย ไม่ต้องให้พนักงานเปิดโต๊ะ
// v2: รูปอาหาร (คอลัมน์ image ในชีท Menu + อัปโหลดจากแท็บ "เมนู")
const IMG_FOLDER = 'Mala Menu Images';
const IMG_MIMES = { 'image/jpeg': 'jpg', 'image/png': 'png', 'image/webp': 'webp' };
const IMG_MAX = 1500 * 1024;   // หน้าเว็บย่อรูปเหลือราว 100 KB ก่อนส่ง
const HEADERS = {
  Employees: ['name', 'pin', 'role'],
  Tables: ['name'],
  Settings: ['key', 'value'],
  Menu: ['id', 'category', 'code', 'name', 'price', 'spicy', 'available', 'sort', 'image'],
  Sessions: ['id', 'tableNo', 'guests', 'paid', 'openedAt', 'closedAt', 'openedBy', 'payMethod', 'closedBy', 'total', 'slip', 'slipRef', 'slipCheck'],
  Orders: ['id', 'sessionId', 'tableNo', 'items', 'status', 'paid', 'createdAt', 'cookAt', 'readyAt', 'servedAt'],
  Calls: ['id', 'tableNo', 'type', 'createdAt', 'doneAt', 'doneBy']
};
const STATUSES = ['new', 'cooking', 'ready', 'served'];
const TOKEN_DAYS = 30;
const MAX_FAILS = 5;
const LOCK_SECONDS = 900;
const MAX_LINES = 60;          // รายการต่อการสั่ง 1 รอบ
const MAX_ORDERS_PER_MIN = 12; // ต่อ 1 โต๊ะ

const SAMPLE_MENU = [
  ['m001', 'เนื้อสัตว์', 'O1', 'สามชั้น', 29, true], ['m002', 'เนื้อสัตว์', 'O3', 'กุ้งขาว', 29, false],
  ['m003', 'เนื้อสัตว์', 'Y1', 'สันนอก สับปะรด', 45, true], ['m004', 'เนื้อสัตว์', 'Y4', 'เอ็นข้อไก่', 29, true],
  ['m005', 'เนื้อสัตว์', 'P5', 'ไส้อ่อนหมูพะโล้', 29, true], ['m006', 'เนื้อสัตว์', 'Bu2', 'เบคอนห่อเห็ดเข็มทอง', 39, false],
  ['m007', 'ลูกชิ้น · เต้าหู้', 'PU1', 'ไส้กรอกไก่', 29, true], ['m008', 'ลูกชิ้น · เต้าหู้', 'Pu3', 'เต้าหู้ปลา', 29, false],
  ['m009', 'ผัก · เห็ด', 'R2', 'กระเทียม', 19, false], ['m010', 'ผัก · เห็ด', 'R3', 'กระเจี๊ยบ', 19, false],
  ['m011', 'ผัก · เห็ด', 'Gr5', 'เห็ดออรินจิย่าง', 19, true], ['m012', 'ผัก · เห็ด', 'Bu5', 'มะเขือเทศ เชอรี่', 39, false],
  ['m013', 'ผัก · เห็ด', 'W1', 'ข้าวโพดข้าวเหนียว', 29, false], ['m014', 'ผัก · เห็ด', 'W2', 'ข้าวโพดหวาน', 29, false],
  ['m015', 'จานร้อน · ซุป', '', 'พริกสกรู กระทะร้อน', 129, false], ['m016', 'จานร้อน · ซุป', '', 'ซุปมิโซะ', 20, false],
  ['m017', 'เครื่องดื่ม', '', 'น้ำเปล่า', 20, false], ['m018', 'เครื่องดื่ม', '', 'น้ำแข็งเปล่า', 2, false],
  ['m019', 'เครื่องดื่ม', '', 'Whisky', 110, false], ['m020', 'เครื่องดื่ม', '', 'จินโร โซจู', 249, false]
];

// ---------------------------------------------------------------- setup (รันเองครั้งเดียว)
function setup() {
  Object.keys(HEADERS).forEach(n => ensureHeaders_(n));
  imgFolder_();   // ให้ Google ขอสิทธิ์ Drive ตั้งแต่ตอนรัน setup
  const emp = getSheet_('Employees');
  if (emp.getLastRow() < 2) emp.appendRow(['Admin', '1234', 'admin']);   // เปลี่ยน PIN ในชีทหลังติดตั้ง
  settings_();
  const tb = getSheet_('Tables');
  if (tb.getLastRow() < 2) for (let i = 1; i <= 12; i++) tb.appendRow(["'" + i]);   // แก้ชื่อ / เพิ่ม / ลบโต๊ะในชีทได้เลย
  const menu = getSheet_('Menu');
  if (menu.getLastRow() < 2) SAMPLE_MENU.forEach((m, i) => menu.appendRow([m[0], m[1], m[2], m[3], m[4], m[5], true, (i + 1) * 10, '']));
}

// ---------------------------------------------------------------- sheet helpers
function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(HEADERS[name]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function rows_(name) {
  const data = getSheet_(name).getDataRange().getValues();
  const head = data[0] || [], out = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i].every(v => v === '')) continue;
    const o = { _row: i + 1 };
    head.forEach((h, j) => o[h] = data[i][j]);
    out.push(o);
  }
  return out;
}
function ensureHeaders_(name) {   // เพิ่มหัวคอลัมน์ที่ยังไม่มีต่อท้าย (ใช้ตอนอัปเดตเวอร์ชัน)
  const sh = getSheet_(name);
  const first = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0] || [];
  let next = first.filter(h => h !== '').length;
  HEADERS[name].forEach(h => { if (first.indexOf(h) < 0) { next++; sh.getRange(1, next, 1, 1).setValues([[h]]); } });
}
function setCells_(name, row, patch) {
  const sh = getSheet_(name), head = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
  Object.keys(patch).forEach(k => { const c = head.indexOf(k); if (c >= 0) sh.getRange(row, c + 1, 1, 1).setValues([[patch[k]]]); });
}
function bool_(v) { return v === true || String(v).trim().toLowerCase() === 'true'; }
function iso_(v) { return v instanceof Date ? v.toISOString() : String(v || ''); }
function str_(v, max) { return String(v === undefined || v === null ? '' : v).replace(/[\u0000-\u001F]/g, '').trim().slice(0, max); }
function jsonOut_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function parse_(s) { try { const v = JSON.parse(s); return Array.isArray(v) ? v : []; } catch (e) { return []; } }

function menuOut_() {
  return rows_('Menu').filter(m => String(m.id) !== '' && String(m.name) !== '').sort((a, b) => Number(a.sort) - Number(b.sort)).map(m => ({
    id: String(m.id), category: String(m.category), code: String(m.code || ''), name: String(m.name), price: Number(m.price) || 0, spicy: bool_(m.spicy), image: imgUrl_(m.image), available: bool_(m.available)
  }));
}
// ลิงก์ Google Drive ทุกแบบ -> ลิงก์รูปย่อที่ใส่ใน <img> ได้ (แบบเดียวกับรูปในแชท TOTU)
function imgUrl_(v) {
  const s = String(v || '').trim();
  if (!/^https:\/\//.test(s)) return '';
  if (/drive\.google\.com|docs\.google\.com/.test(s)) {
    const m = /\/d\/([\w-]{20,})/.exec(s) || /[?&]id=([\w-]{20,})/.exec(s);
    return m ? 'https://drive.google.com/thumbnail?id=' + m[1] + '&sz=w640' : '';
  }
  return s;
}
function driveId_(v) { const m = /[?&]id=([\w-]{20,})/.exec(String(v || '')) || /\/d\/([\w-]{20,})/.exec(String(v || '')); return m ? m[1] : ''; }
function imgFolder_() {
  const it = DriveApp.getFoldersByName(IMG_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(IMG_FOLDER);
}
function subFolder_(parent, name) { const it = parent.getFoldersByName(name); return it.hasNext() ? it.next() : parent.createFolder(name); }
function isAdmin_(emp) { return String(emp.role || '').trim().toLowerCase() === 'admin'; }
function settings_() {
  let rows = rows_('Settings');
  if (!rows.length) { const sh = getSheet_('Settings'); sh.getRange('B:B').setNumberFormat('@');   // เก็บเป็นข้อความ เลข 0 นำหน้าจะไม่หาย
    sh.appendRow(['promptpay_id', '']); sh.appendRow(['promptpay_name', '']); sh.appendRow(['promptpay_note', '']); rows = []; }
  const m = {};
  rows.forEach(r => { m[String(r.key).trim()] = String(r.value === undefined || r.value === null ? '' : r.value).trim(); });
  return { promptpayId: m.promptpay_id || '', promptpayName: m.promptpay_name || '', promptpayNote: m.promptpay_note || '' };
}
function tables_() { return rows_('Tables').map(r => String(r.name).trim()).filter(Boolean); }
function openSession_(table) { return rows_('Sessions').find(r => !bool_(r.paid) && String(r.tableNo) === table) || null; }
function sessionOut_(s) {
  return { id: String(s.id), tableNo: String(s.tableNo), guests: Number(s.guests) || 1, paid: bool_(s.paid), openedAt: iso_(s.openedAt), closedAt: iso_(s.closedAt) };
}
function orderOut_(o) {
  return { no: 99 + o._row, id: String(o.id), sessionId: String(o.sessionId), tableNo: String(o.tableNo), items: parse_(o.items), status: String(o.status), paid: bool_(o.paid), createdAt: iso_(o.createdAt) };
}

// ---------------------------------------------------------------- staff tokens (แบบเดียวกับ TOTU)
function secret_() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('MALA_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('MALA_SECRET', s); }
  return s;
}
function sign_(msg) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(msg, secret_())); }
function safeEq_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  let d = 0;
  for (let i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}
function makeToken_(emp) {   // เปลี่ยน PIN = token เก่าใช้ไม่ได้
  const payload = String(emp.name).trim() + '|' + (Date.now() + TOKEN_DAYS * 86400000);
  return Utilities.base64EncodeWebSafe(payload, Utilities.Charset.UTF_8) + '.' + sign_(payload + '|' + String(emp.pin).trim());
}
function verifyToken_(token) {
  if (!token || typeof token !== 'string') return null;
  const i = token.indexOf('.');
  if (i < 1) return null;
  let payload;
  try { payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(token.slice(0, i))).getDataAsString(); } catch (err) { return null; }
  const j = payload.lastIndexOf('|');
  if (j < 1) return null;
  const name = payload.slice(0, j);
  if (!(Number(payload.slice(j + 1)) > Date.now())) return null;
  const emp = rows_('Employees').find(r => String(r.name).trim() === name);
  if (!emp) return null;
  return safeEq_(token.slice(i + 1), sign_(payload + '|' + String(emp.pin).trim())) ? emp : null;
}

// ---------------------------------------------------------------- entry points
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'names') return jsonOut_({ ok: true, names: rows_('Employees').map(r => String(r.name).trim()).filter(Boolean), version: SERVER_VERSION });
  if (action === 'menu') return jsonOut_({ ok: true, menu: menuOut_(), tables: tables_(), settings: settings_() });
  return jsonOut_({ ok: false, error: 'use_post', version: SERVER_VERSION });
}
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return jsonOut_({ ok: false, error: 'bad_json' }); }
  const action = body.action;
  // ลูกค้า (ไม่ต้องล็อกอิน ใช้รหัสใน QR)
  if (action === 'login') return login_(body);
  if (action === 'custSync') return custSync_(body);
  if (action === 'custOrder') return withLock_(() => custOrder_(body));
  if (action === 'custCall') return withLock_(() => custCall_(body));
  // พนักงาน
  const emp = verifyToken_(body.token);
  if (!emp) return jsonOut_({ ok: false, error: 'auth' });
  if (action === 'sync') return sync_();
  if (action === 'setOrder') return withLock_(() => setOrder_(body));
  if (action === 'ackCall') return withLock_(() => ackCall_(emp, body));
  if (action === 'pay') return withLock_(() => pay_(emp, body));
  if (action === 'saveMenu') return withLock_(() => saveMenu_(emp, body));
  if (action === 'deleteMenu') return withLock_(() => deleteMenu_(emp, body));
  if (action === 'setAvailable') return withLock_(() => setAvailable_(body));
  if (action === 'report') return report_(emp, body);
  if (action === 'checkSlip') return checkSlip_(emp, body);
  if (action === 'slipKey') return slipKeySet_(emp, body);
  if (action === 'setMenuImage') return withLock_(() => setMenuImage_(emp, body));
  return jsonOut_({ ok: false, error: 'unknown_action' });
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) { return jsonOut_({ ok: false, error: 'busy' }); }
  try { return fn(); } finally { lock.releaseLock(); }
}

// ---------------------------------------------------------------- actions
function login_(body) {
  const name = str_(body.name, 60), pin = str_(body.pin, 20);
  if (!name || !pin) return jsonOut_({ ok: false, error: 'invalid_login' });
  const cache = CacheService.getScriptCache(), key = 'fail:' + name;
  const fails = Number(cache.get(key) || 0);
  if (fails >= MAX_FAILS) return jsonOut_({ ok: false, error: 'locked' });
  const emp = rows_('Employees').find(r => String(r.name).trim() === name && String(r.pin).trim() === pin);
  if (!emp) { cache.put(key, String(fails + 1), LOCK_SECONDS); return jsonOut_({ ok: false, error: 'invalid_login' }); }
  cache.remove(key);
  return jsonOut_({ ok: true, token: makeToken_(emp), name: name });
}

function custSync_(body) {
  const table = str_(body.table, 20);
  if (tables_().indexOf(table) < 0) return jsonOut_({ ok: true, valid: false, session: null, orders: [], soldOut: [] });
  const soldOut = rows_('Menu').filter(m => !bool_(m.available)).map(m => String(m.id));   // ให้หน้าลูกค้าอัปเดตของหมดทุกรอบ
  const s = openSession_(table);
  const calls = openCalls_().filter(c => String(c.tableNo) === table).map(c => String(c.type));
  if (!s) return jsonOut_({ ok: true, valid: true, session: null, orders: [], soldOut: soldOut, calls: calls });
  const sid = String(s.id), all = rows_('Orders');
  // คิวในครัว = ลำดับของออเดอร์นี้ในบรรดาออเดอร์ที่ครัวยังไม่ได้ทำ (ทุกโต๊ะ เรียงตามเวลาที่สั่ง)
  const waiting = all.filter(o => !bool_(o.paid) && String(o.status) === 'new').map(o => String(o.id));
  return jsonOut_({ ok: true, valid: true, session: sessionOut_(s), soldOut: soldOut, calls: calls, orders: all.filter(o => String(o.sessionId) === sid).map(o => {
    const out = orderOut_(o); out.queue = waiting.indexOf(out.id) + 1; return out; }) });
}

function custOrder_(body) {
  const table = str_(body.table, 20);
  if (tables_().indexOf(table) < 0) return jsonOut_({ ok: false, error: 'bad_table' });
  const cache = CacheService.getScriptCache(), key = 'ord:' + table;
  const n = Number(cache.get(key) || 0);
  if (n >= MAX_ORDERS_PER_MIN) return jsonOut_({ ok: false, error: 'too_fast' });
  const menu = {}, gone = [];
  rows_('Menu').forEach(m => { menu[String(m.id)] = m; });
  const items = [];
  (Array.isArray(body.items) ? body.items : []).slice(0, MAX_LINES).forEach(it => {
    const m = menu[String(it && it.id)], qty = Math.floor(Number(it && it.qty));
    if (!m || !(qty >= 1 && qty <= 99)) return;
    if (!bool_(m.available)) { gone.push(String(m.id)); return; }
    items.push({ id: String(m.id), name: String(m.name), code: String(m.code || ''), price: Number(m.price) || 0, qty: qty,
                 opt: bool_(m.spicy) ? str_(it.opt, 30) : '', served: false });
  });
  // มีของหมดปนมา = ไม่รับทั้งรอบ ให้ลูกค้าเห็นและยืนยันรายการที่เหลือเองก่อน
  if (gone.length) return jsonOut_({ ok: false, error: 'sold_out', ids: gone });
  if (!items.length) return jsonOut_({ ok: false, error: 'empty' });
  cache.put(key, String(n + 1), 60);
  const id = 'o-' + Utilities.getUuid().slice(0, 13);
  // โต๊ะว่าง = เปิดบิลใหม่ให้อัตโนมัติ ( ' นำหน้าให้ชีทเก็บชื่อโต๊ะเป็นข้อความ เช่น 05, B18 )
  const s = openSession_(table);
  let sid = s ? String(s.id) : '';
  if (!s) {
    sid = 's-' + Utilities.getUuid().replace(/-/g, '').slice(0, 20);
    getSheet_('Sessions').appendRow([sid, "'" + table, '', false, new Date().toISOString(), '', 'QR']);
  }
  getSheet_('Orders').appendRow([id, sid, "'" + table, JSON.stringify(items), 'new', false, new Date().toISOString()]);
  return jsonOut_({ ok: true, id: id });
}

function sync_() {
  return jsonOut_({ ok: true,
    sessions: rows_('Sessions').filter(s => !bool_(s.paid)).map(sessionOut_),
    orders: rows_('Orders').filter(o => !bool_(o.paid)).map(orderOut_),
    calls: openCalls_().map(c => ({ id: String(c.id), tableNo: String(c.tableNo), type: String(c.type), createdAt: iso_(c.createdAt) })) });
}

// ---------------------------------------------------------------- เรียกพนักงาน / ขอเช็คบิล
function openCalls_() { return rows_('Calls').filter(c => String(c.doneAt || '') === ''); }
function custCall_(body) {
  const table = str_(body.table, 20), type = body.type === 'bill' ? 'bill' : 'staff';
  if (tables_().indexOf(table) < 0) return jsonOut_({ ok: false, error: 'bad_table' });
  // โต๊ะเดียวกันเรียกซ้ำประเภทเดิมได้ครั้งเดียวจนกว่าพนักงานจะรับทราบ (กันกดรัว)
  if (openCalls_().some(c => String(c.tableNo) === table && String(c.type) === type)) return jsonOut_({ ok: true, already: true });
  getSheet_('Calls').appendRow(['c-' + Utilities.getUuid().slice(0, 13), "'" + table, type, new Date().toISOString(), '', '']);
  return jsonOut_({ ok: true });
}
function ackCall_(emp, body) {
  const c = rows_('Calls').find(r => String(r.id) === str_(body.id, 60));
  if (!c) return jsonOut_({ ok: false, error: 'not_found' });
  if (String(c.doneAt || '') === '') setCells_('Calls', c._row, { doneAt: new Date().toISOString(), doneBy: String(emp.name).trim() });
  return jsonOut_({ ok: true });
}

function setOrder_(body) {
  const o = rows_('Orders').find(r => String(r.id) === str_(body.id, 60));
  if (!o) return jsonOut_({ ok: false, error: 'not_found' });
  const patch = {};
  if (STATUSES.indexOf(body.status) >= 0) {
    patch.status = body.status;
    // v13: จดเวลาครั้งแรกที่ถึงแต่ละขั้น ใช้คำนวณเวลารอและเวลาย่างในรายงาน
    const stamp = { cooking: 'cookAt', ready: 'readyAt', served: 'servedAt' }[body.status];
    if (stamp && String(o[stamp] || '') === '') { ensureHeaders_('Orders'); patch[stamp] = new Date().toISOString(); }
  }
  if (Array.isArray(body.served)) {   // รับเฉพาะเครื่องหมายเสิร์ฟ ไม่รับรายการหรือราคาจากหน้าเว็บ
    const items = parse_(o.items);
    items.forEach((it, i) => { it.served = body.served[i] === true; });
    patch.items = JSON.stringify(items);
  }
  setCells_('Orders', o._row, patch);
  return jsonOut_({ ok: true });
}

function pay_(emp, body) {
  const sid = str_(body.sessionId, 60);
  const s = rows_('Sessions').find(r => String(r.id) === sid);
  if (!s) return jsonOut_({ ok: false, error: 'not_found' });
  const mine = rows_('Orders').filter(o => String(o.sessionId) === sid && !bool_(o.paid));
  let method = String(body.method || '');
  if (!mine.length) method = 'none';
  else if (method !== 'cash' && method !== 'transfer') return jsonOut_({ ok: false, error: 'need_method' });
  // กันสลิปปลอม: บิลที่จ่ายด้วยการโอน พนักงานต้องยืนยันว่าเห็นยอดเงินเข้าบัญชีจริง
  if (method === 'transfer' && body.seen !== true) return jsonOut_({ ok: false, error: 'need_confirm' });
  ensureHeaders_('Sessions');
  // รูปสลิป (ถ้ามี): บันทึกให้สำเร็จก่อนค่อยปิดบิล ไม่แชร์ไฟล์ เพราะมีชื่อและเลขบัญชีของลูกค้า
  let slip = '';
  if (method === 'transfer' && body.slipData) {
    try {
      const ext = IMG_MIMES[body.slipMime];
      const bytes = Utilities.base64Decode(String(body.slipData));
      if (!ext || !bytes.length || bytes.length > IMG_MAX) return jsonOut_({ ok: false, error: 'slip_failed' });
      // v13: แยกโฟลเดอร์ ปี / เดือน / วัน เช่น Mala Payment Slips / 2026 / 2026-10 / 2026-10-11
      const it = DriveApp.getFoldersByName(SLIP_FOLDER), root = it.hasNext() ? it.next() : DriveApp.createFolder(SLIP_FOLDER);
      const today = day_(new Date()), folder = subFolder_(subFolder_(subFolder_(root, today.slice(0, 4)), today.slice(0, 7)), today);
      const stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMdd-HHmm');
      const file = folder.createFile(Utilities.newBlob(bytes, body.slipMime, 'slip-' + stamp + '-table' + String(s.tableNo) + '-' + (Math.max(0, Number(body.total) || 0)) + '.' + ext));
      slip = 'https://drive.google.com/file/d/' + file.getId() + '/view';
    } catch (err) { return jsonOut_({ ok: false, error: 'slip_failed' }); }
  }
  // v14: ผลตรวจสลิปที่เซิร์ฟเวอร์ตรวจไว้เอง (ไม่รับจากหน้าเว็บ) บันทึกลงบิล
  let slipRef = '', slipCheck = '';
  if (method === 'transfer') { const c = parse1_(CacheService.getScriptCache().get('slipok:' + sid)); if (c) { slipRef = String(c.ref || ''); slipCheck = String(c.status || '') + ' ' + String(c.amount); } }
  mine.forEach(o => setCells_('Orders', o._row, { paid: true }));
  // ปิดบิลแล้ว = การเรียกที่ค้างของโต๊ะนี้ถือว่าจบ
  openCalls_().forEach(c => { if (String(c.tableNo) === String(s.tableNo)) setCells_('Calls', c._row, { doneAt: new Date().toISOString(), doneBy: String(emp.name).trim() }); });
  setCells_('Sessions', s._row, { paid: true, closedAt: new Date().toISOString(), payMethod: method,
    closedBy: String(emp.name).trim(), total: Math.max(0, Number(body.total) || 0), slip: slip, slipRef: slipRef, slipCheck: slipCheck });
  return jsonOut_({ ok: true });
}

// ใส่ / เปลี่ยน / ลบรูปอาหาร (data ว่าง = ลบรูป)
function setMenuImage_(emp, body) {
  if (!isAdmin_(emp)) return jsonOut_({ ok: false, error: 'forbidden' });
  ensureHeaders_('Menu');
  const m = rows_('Menu').find(r => String(r.id) === str_(body.id, 60));
  if (!m) return jsonOut_({ ok: false, error: 'not_found' });
  let url = '';
  if (body.data) {
    const ext = IMG_MIMES[body.mime];
    if (!ext) return jsonOut_({ ok: false, error: 'bad_image' });
    let bytes;
    try { bytes = Utilities.base64Decode(String(body.data)); } catch (err) { return jsonOut_({ ok: false, error: 'bad_image' }); }
    if (!bytes.length || bytes.length > IMG_MAX) return jsonOut_({ ok: false, error: 'too_big' });
    const file = imgFolder_().createFile(Utilities.newBlob(bytes, body.mime, String(m.id) + '-' + Date.now() + '.' + ext));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);   // ลูกค้าต้องเปิดรูปได้โดยไม่ล็อกอิน
    url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w640';
  }
  // ทิ้งรูปเก่าเฉพาะไฟล์ที่ระบบอัปโหลดเอง (อยู่ในโฟลเดอร์ของระบบ) ไม่แตะไฟล์ที่วางลิงก์เอง
  const old = driveId_(m.image);
  if (old) {
    try {
      const f = DriveApp.getFileById(old), ps = f.getParents();
      if (ps.hasNext() && ps.next().getName() === IMG_FOLDER) f.setTrashed(true);
    } catch (err) {}
  }
  setCells_('Menu', m._row, { image: url });
  return jsonOut_({ ok: true, image: url });
}

// สรุปยอดขายช่วง from..to (yyyy-MM-dd ตามเวลาไทย) นับเฉพาะออเดอร์ของบิลที่ปิดแล้ว
// ยอดเป็นราคาอาหารตามชีท Orders ยังไม่รวม VAT / ค่าบริการที่ตั้งในหน้าเว็บ
function day_(v) {
  const d = v instanceof Date ? v : new Date(String(v));
  return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, TIMEZONE, 'yyyy-MM-dd');
}
// ---------------------------------------------------------------- v14: ตรวจสลิปด้วย Thunder
function parse1_(s) { try { return s ? JSON.parse(s) : null; } catch (err) { return null; } }
function slipKey_() { return PropertiesService.getScriptProperties().getProperty('THUNDER_API_KEY') || ''; }
// admin ใส่ / ลบ API key จากหน้าเว็บ (ส่ง key = '' เพื่อลบ) ตอบกลับแค่ว่ามี key หรือไม่ ไม่ส่ง key กลับ
function slipKeySet_(emp, body) {
  if (!isAdmin_(emp)) return jsonOut_({ ok: false, error: 'forbidden' });
  if (typeof body.key === 'string') PropertiesService.getScriptProperties().setProperty('THUNDER_API_KEY', str_(body.key, 300));
  return jsonOut_({ ok: true, has: !!slipKey_() });
}
function checkSlip_(emp, body) {
  const key = slipKey_();
  if (!key) return jsonOut_({ ok: false, error: 'no_slip_key' });
  const sid = str_(body.sessionId, 60), total = Math.max(0, Number(body.total) || 0);
  const sessions = rows_('Sessions'), s = sessions.find(r => String(r.id) === sid);
  if (!s) return jsonOut_({ ok: false, error: 'not_found' });
  const data = String(body.slipData || '');
  if (!IMG_MIMES[body.slipMime] || !data || data.length > IMG_MAX * 1.4) return jsonOut_({ ok: false, error: 'slip_failed' });
  let code = 0, r = null;
  try {
    const res = UrlFetchApp.fetch(THUNDER_URL, { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + key },
      payload: JSON.stringify({ base64: data, matchAmount: total, checkDuplicate: true, remark: 'table ' + String(s.tableNo) }) });
    code = res.getResponseCode(); r = parse1_(res.getContentText());
  } catch (err) { return jsonOut_({ ok: false, error: 'slip_api' }); }
  if (!r || r.success !== true || !r.data) {
    const ec = String((r && r.error && r.error.code) || '');
    if (ec === 'SLIP_NOT_FOUND') return jsonOut_({ ok: true, result: { status: 'fail', reasons: ['not_found'] } });
    const map = { SLIP_PENDING: 'slip_pending', QUOTA_EXCEEDED: 'slip_quota', INVALID_API_KEY: 'slip_key', MISSING_API_KEY: 'slip_key', IMAGE_SIZE_TOO_LARGE: 'slip_failed', INVALID_IMAGE_FORMAT: 'slip_failed' };
    return jsonOut_({ ok: false, error: map[ec] || 'slip_api', code: ec || String(code) });
  }
  const d = r.data, raw = d.rawSlip || {}, rc = raw.receiver || {}, sd = raw.sender || {};
  const nameOf = p => { const n = (p.account && p.account.name) || {}; return String(n.th || n.en || ''); };
  const acctOf = p => { const a = p.account || {}; return String((a.bank && a.bank.account) || (a.proxy && a.proxy.account) || ''); };
  const amount = Number(d.amountInSlip !== undefined ? d.amountInSlip : (raw.amount && raw.amount.amount)) || 0;
  const ref = String(raw.transRef || ''), reasons = [], cache = CacheService.getScriptCache();
  // สลิปซ้ำ: Thunder เคยเห็นแล้ว (ยกเว้นเป็นการตรวจซ้ำของบิลเดียวกันนี้เอง) หรือเคยใช้ปิดบิลอื่นในชีทเรา
  const mine = parse1_(cache.get('slipok:' + sid));
  let dup = d.isDuplicate === true && !(mine && mine.ref && mine.ref === ref);
  if (ref && sessions.some(x => String(x.slipRef || '') === ref && String(x.id) !== sid)) dup = true;
  if (dup) reasons.push('duplicate');
  const amountOk = d.isAmountMatched !== undefined ? d.isAmountMatched === true : Math.abs(amount - total) < 0.005;
  if (!amountOk) reasons.push('amount');
  const when = date_(raw.date), age = when ? Math.round((Date.now() - when.getTime()) / 60000) : null;
  if (age !== null && (age > SLIP_MAX_AGE_MIN || age < -10)) reasons.push('old');
  // บัญชีผู้รับ: เทียบเลขที่มองเห็นในสลิปกับเลขบัญชี / พร้อมเพย์ของร้านในชีท Settings ถ้าเทียบไม่ได้ให้พนักงานดูชื่อเอง
  const st = settings_(), shopDigits = (st.promptpayNote + ' ' + st.promptpayId).replace(/\D/g, ' ').split(' ').join('');
  const runs = acctOf(rc).split(/\D+/).filter(x => x.length >= 4).sort((a, b) => b.length - a.length);
  const account = runs.length && shopDigits.indexOf(runs[0]) >= 0 ? 'ok' : 'unknown';
  if (account !== 'ok') reasons.push('account');
  const status = dup || !amountOk ? 'fail' : reasons.length ? 'warn' : 'pass';
  if (!dup) cache.put('slipok:' + sid, JSON.stringify({ ref: ref, status: status, amount: amount }), 1800);
  return jsonOut_({ ok: true, result: { status: status, reasons: reasons, amount: amount, expected: total, date: when ? when.toISOString() : '', ageMin: age,
    sender: nameOf(sd), senderBank: String((sd.bank && (sd.bank.short || sd.bank.name)) || ''),
    receiver: nameOf(rc), receiverBank: String((rc.bank && (rc.bank.short || rc.bank.name)) || ''), receiverAccount: acctOf(rc), account: account, ref: ref } });
}

// v13: สถิติตามช่วงเวลา (เวลาไทย UTC+7) รายชั่วโมงและรายวันในสัปดาห์
function date_(v) { const d = v instanceof Date ? v : new Date(String(v || '')); return isNaN(d.getTime()) ? null : d; }
function timeStats_(from, to) {
  const mk = () => ({ tables: 0, orders: 0, qty: 0, done: 0, waitSum: 0, waitN: 0, grillSum: 0, grillN: 0 });
  const hours = [], dow = [], dowDays = [], dates = {}, slot = {};
  for (let i = 0; i < 24; i++) hours.push(mk());
  for (let i = 0; i < 7; i++) { dow.push(mk()); dowDays.push({}); }
  const th = d => new Date(d.getTime() + 7 * 3600000);   // เลื่อนเป็นเวลาไทยแล้วอ่านแบบ UTC
  rows_('Sessions').forEach(s => {
    const t = date_(s.openedAt), d = day_(s.openedAt); if (!t || d < from || d > to) return;
    const x = th(t); hours[x.getUTCHours()].tables++; dow[x.getUTCDay()].tables++; dates[d] = 1; dowDays[x.getUTCDay()][d] = 1;
  });
  rows_('Orders').forEach(o => {
    const t = date_(o.createdAt), d = day_(o.createdAt); if (!t || d < from || d > to) return;
    const x = th(t), H = hours[x.getUTCHours()], W = dow[x.getUTCDay()];
    let qty = 0; parse_(o.items).forEach(c => { qty += Number(c.qty) || 0; });
    dates[d] = 1; dowDays[x.getUTCDay()][d] = 1;
    [H, W].forEach(a => { a.orders++; a.qty += qty; });
    const r = date_(o.readyAt), c = date_(o.cookAt);
    if (r) {
      const w = (r - t) / 60000; [H, W].forEach(a => { a.waitSum += w; a.waitN++; });
      if (c) { const gm = (r - c) / 60000; [H, W].forEach(a => { a.grillSum += gm; a.grillN++; }); }
      const rx = th(r), k = day_(r) + ' ' + rx.getUTCHours();
      hours[rx.getUTCHours()].done += qty; dow[rx.getUTCDay()].done += qty; slot[k] = (slot[k] || 0) + qty;
    }
  });
  let peak = null;
  Object.keys(slot).forEach(k => { if (!peak || slot[k] > peak.qty) peak = { date: k.split(' ')[0], hour: Number(k.split(' ')[1]), qty: slot[k] }; });
  return { hours: hours, dow: dow, dowDays: dowDays.map(x => Object.keys(x).length), ndays: Object.keys(dates).length, peak: peak };
}
function report_(emp, body) {
  if (!isAdmin_(emp)) return jsonOut_({ ok: false, error: 'forbidden' });
  const from = str_(body.from, 10), to = str_(body.to, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to)) return jsonOut_({ ok: false, error: 'bad_range' });
  const days = {}, items = {}, methods = { cash: 0, transfer: 0, other: 0 }, how = {};
  rows_('Sessions').forEach(x => { how[String(x.id)] = String(x.payMethod || ''); });
  let open = 0;
  rows_('Orders').forEach(o => {
    const d = day_(o.createdAt);
    if (!d || d < from || d > to) return;
    const list = parse_(o.items);
    let amt = 0, qty = 0;
    list.forEach(c => { amt += (Number(c.price) || 0) * (Number(c.qty) || 0); qty += Number(c.qty) || 0; });
    if (!bool_(o.paid)) { open += amt; return; }
    const day = days[d] || (days[d] = { date: d, total: 0, qty: 0, bills: {} });
    day.total += amt; day.qty += qty; day.bills[String(o.sessionId)] = 1;
    const m = how[String(o.sessionId)];
    methods[m === 'cash' ? 'cash' : m === 'transfer' ? 'transfer' : 'other'] += amt;
    list.forEach(c => {
      const k = String(c.id || c.name);
      const it = items[k] || (items[k] = { id: String(c.id || ''), name: String(c.name), code: String(c.code || ''), qty: 0, total: 0 });
      it.qty += Number(c.qty) || 0; it.total += (Number(c.price) || 0) * (Number(c.qty) || 0);
    });
  });
  return jsonOut_({ ok: true, open: open, methods: methods, time: timeStats_(from, to),
    days: Object.keys(days).sort().map(k => ({ date: k, total: days[k].total, qty: days[k].qty, bills: Object.keys(days[k].bills).length })),
    items: Object.keys(items).map(k => items[k]) });
}

// ตั้งของหมด / กลับมาขาย (พนักงานทุกคนกดได้ เพราะครัวต้องใช้)
function setAvailable_(body) {
  const m = rows_('Menu').find(r => String(r.id) === str_(body.id, 60));
  if (!m) return jsonOut_({ ok: false, error: 'not_found' });
  setCells_('Menu', m._row, { available: body.available === true });
  return jsonOut_({ ok: true });
}

// เพิ่มเมนูใหม่ (ไม่มี id) หรือแก้ไขเมนูเดิม: ชื่อ ราคา หมวด รหัสไม้ เลือกเผ็ดได้
function saveMenu_(emp, body) {
  if (!isAdmin_(emp)) return jsonOut_({ ok: false, error: 'forbidden' });
  const name = str_(body.name, 80), category = str_(body.category, 40), code = str_(body.code, 10), price = Number(body.price);
  if (!name || !category || !(price >= 0 && price <= 100000)) return jsonOut_({ ok: false, error: 'bad_menu' });
  ensureHeaders_('Menu');
  const rows = rows_('Menu'), id = str_(body.id, 60);
  if (id) {
    const m = rows.find(r => String(r.id) === id);
    if (!m) return jsonOut_({ ok: false, error: 'not_found' });
    setCells_('Menu', m._row, { name: name, category: category, code: "'" + code, price: price, spicy: body.spicy === true });
    return jsonOut_({ ok: true, id: id });
  }
  let max = 0, sort = 0;
  rows.forEach(r => { const k = /^m(\d+)$/.exec(String(r.id)); if (k) max = Math.max(max, Number(k[1])); sort = Math.max(sort, Number(r.sort) || 0); });
  const n = String(max + 1), newId = 'm' + (n.length >= 3 ? n : ('00' + n).slice(-3));
  const obj = { id: newId, category: category, code: "'" + code, name: name, price: price, spicy: body.spicy === true, available: true, sort: sort + 10, image: '' };
  const sh = getSheet_('Menu'), head = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
  sh.appendRow(head.map(h => (obj[h] === undefined ? '' : obj[h])));
  return jsonOut_({ ok: true, id: newId });
}
// ลบเมนูออกจากชีท (ออเดอร์เก่าเก็บชื่อและราคาไว้เอง ยอดขายย้อนหลังจึงไม่หาย)
function deleteMenu_(emp, body) {
  if (!isAdmin_(emp)) return jsonOut_({ ok: false, error: 'forbidden' });
  const m = rows_('Menu').find(r => String(r.id) === str_(body.id, 60));
  if (!m) return jsonOut_({ ok: false, error: 'not_found' });
  const old = driveId_(m.image);
  if (old) { try { const f = DriveApp.getFileById(old), ps = f.getParents(); if (ps.hasNext() && ps.next().getName() === IMG_FOLDER) f.setTrashed(true); } catch (err) {} }
  getSheet_('Menu').deleteRow(m._row);
  return jsonOut_({ ok: true });
}
