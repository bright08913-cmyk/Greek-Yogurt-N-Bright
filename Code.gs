/**
 * Greek Yogurt Order App — Google Apps Script backend
 *
 * Sheets created by setup():
 *   Menu       เมนู + ราคา + จำนวนผลไม้ที่เลือกได้   (แก้ราคาที่นี่ → แอปเปลี่ยนตาม)
 *   AddOns     ท็อปปิ้งเพิ่ม (scoop, hershey)
 *   Fruits     รายการผลไม้ + เปิด/ปิด ของหมด
 *   Settings   รอบรับ, วิธีรับ/ค่าส่ง, PromptPay, เปิด/ปิดรับออเดอร์
 *   Orders     แท็บเดียวดูได้ครบ: 1 แถว = 1 รายการในตะกร้า (เมนู ผลไม้ add-on จำนวน วันรับ จ่ายเงิน สลิป สถานะ)
 *   Orders Archive  ออเดอร์สัปดาห์ก่อนๆ (เมนู 🥣 Greek Shop → เก็บออเดอร์) ลูกค้ายังเห็นในประวัติเหมือนเดิม
 *   Costs      จดต้นทุน + สรุปกำไรรายสัปดาห์
 *   Summary    สรุปต่อรอบ: จำนวนแต่ละเมนู, ผลไม้ที่ต้องเตรียม, ยอดเงิน จ่ายแล้ว/ค้าง
 */

const SH = {
  ROUND: 'Rounds', MENU: 'Menu', ADDON: 'AddOns', FRUIT: 'Fruits', SET: 'Settings',
  ORD: 'Orders', SUM: 'Summary', ARCH: 'Orders Archive', COST: 'Costs'
};
const PAY_STATUS = ['ยังไม่จ่าย', 'ส่งสลิปแล้ว', 'จ่ายแล้ว'];
const PAY_NOW = 'จ่ายเลย (แนบสลิป)';
const PAY_LATER = 'จ่ายตอนรับ (ขุนทอง)';
const SLIP_FOLDER = 'Bright Greek Slips';
const ORDER_STATUS = ['ใหม่', 'ทำแล้ว', 'ส่งแล้ว', 'ยกเลิก'];

// คอลัมน์เรียงตามลำดับที่ลูกค้ากดในแอป: เมนู → ผลไม้ → add-on → จำนวน → วันรับ → ชื่อ/หมายเหตุ → จ่ายเงิน
const ORD_HEAD = ['OrderID', 'เวลาสั่ง', 'ชื่อ', 'รับวันที่', 'เมนู', 'ผลไม้', 'Scoop เพิ่ม', 'ท็อปปิ้ง',
  'จำนวน', 'ราคา/ถ้วย', 'รวม', 'หมายเหตุ', 'วิธีจ่าย', 'ยอดโอน (ทั้งออเดอร์)', 'สถานะจ่าย', 'สลิป', 'ลิงก์สลิป',
  'สถานะออเดอร์', 'token'];
const COL = { ID: 1, TIME: 2, NAME: 3, DAY: 4, MENU: 5, FRUIT: 6, SCOOP: 7, TOP: 8, QTY: 9, UNIT: 10, TOTAL: 11,
  NOTE: 12, METHOD: 13, SUM: 14, PAY: 15, SLIP: 16, LINK: 17, STATUS: 18, TOKEN: 19 };
const at_ = (r, c) => r[c - 1];          // อ่านค่าจากแถว (array) ตามชื่อคอลัมน์
const base_ = id => String(id).replace(/-\d+$/, '');
const ARCH_HEAD = ORD_HEAD.concat(['Week', 'เก็บเมื่อ']);
const COST_HEAD = ['วันที่ซื้อ', 'Week', 'รายการ', 'หมวด', 'จำนวน', 'หน่วย', 'ราคารวม (บาท)', 'ซื้อที่', 'หมายเหตุ'];
const COST_CATS = ['โยเกิร์ต/นม', 'ผลไม้', 'ท็อปปิ้ง', 'แพ็กเกจจิ้ง', 'ค่าส่ง/เดินทาง', 'อื่นๆ'];
// rows of the live Orders tab, then the archive (so customers still see last week's orders)
function orderSources_() {
  const ss = SpreadsheetApp.getActive();
  return [ss.getSheetByName(SH.ORD), ss.getSheetByName(SH.ARCH)].filter(Boolean).map(sh => ({
    sh, vals: sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, ORD_HEAD.length).getValues() }));
}

/* ---------------- Web app ---------------- */

/* หน้าเว็บ (GitHub Pages) เรียกผ่าน API นี้:
 *   GET  .../exec?action=config            → เมนู/ราคา/รอบ
 *   POST .../exec  {action:'order', order}  → สร้างออเดอร์
 *   POST .../exec  {action:'slip', orderId, token, image}  → แนบสลิป
 * (POST ส่งเป็น text/plain เพื่อไม่ติด CORS) */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'order') return json_({ ok: true, data: submitOrder(body.order) });
    if (body.action === 'slip') {
      try { return json_({ ok: true, data: attachSlip(body) }); }
      catch (err) {
        // เจ้าของร้านยังไม่ได้กด Run "authorize" — ลูกค้าไม่ต้องเห็นข้อความเทคนิค
        if (/permission|DriveApp|authoriz/i.test(String(err.message))) {
          console.error(err);
          throw new Error('ร้านยังเปิดรับสลิปไม่ได้ตอนนี้ ออเดอร์เข้าแล้ว ส่งสลิปให้ร้านทาง LINE ได้เลยค่า');
        }
        throw err;
      }
    }
    if (body.action === 'config') return json_({ ok: true, data: getConfig() });
    if (body.action === 'history') return json_({ ok: true, data: myOrders(body) });
    if (body.action === 'cancel') return json_({ ok: true, data: cancelOrder(body) });
    throw new Error('unknown action');
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'config') {
    try { return json_({ ok: true, data: getConfig() }); }
    catch (err) { return json_({ ok: false, error: String(err.message || err) }); }
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('สั่ง Greek Yogurt')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🥣 Greek Shop')
    .addItem('📅 เรียงตามวันรับ', 'sortByPickupDay')
    .addItem('🕒 เรียงตามเวลาสั่ง', 'sortByOrderTime')
    .addSeparator()
    .addItem('📦 เก็บออเดอร์ (ปิดรอบสัปดาห์นี้)', 'archiveOrders')
    .addSeparator()
    .addItem('ตั้งค่าครั้งแรก (สร้างชีตที่ยังไม่มี)', 'setup')
    .addToUi();
}

/* ---------------- เรียงแถว ----------------
 * "9 Nov" เป็นข้อความ ถ้าใช้ sort A→Z ของชีต "10 Nov" จะมาก่อน "9 Nov" — ปุ่มนี้เรียงตามวันจริง
 * แถวของออเดอร์เดียวกันอยู่ติดกันเสมอ · ใช้กับแท็บที่เปิดอยู่ (Orders หรือ Orders Archive) */
function isCellImage_(v) { return !!v && typeof v === 'object' && typeof v.getContentUrl === 'function'; }
function sortByPickupDay() { sortOrders_('day'); }
function sortByOrderTime() { sortOrders_('time'); }
function sortOrders_(mode) {
  const ss = SpreadsheetApp.getActive(), active = ss.getActiveSheet();
  const sh = active && active.getName() === SH.ARCH ? active : ss.getSheetByName(SH.ORD);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const n = sh.getLastRow() - 1; if (n < 2) return;
    const rng = sh.getRange(2, 1, n, sh.getLastColumn());
    const vals = rng.getValues(), forms = rng.getFormulas();
    const num = id => { const m = String(id).match(/^G(\d+)(?:-(\d+))?/); return m ? [+m[1], +(m[2] || 0)] : [Infinity, 0]; };
    const rows = vals.map((r, i) => {
      const id = String(at_(r, COL.ID)), d = dayDate_(at_(r, COL.DAY)), t = at_(r, COL.TIME);
      const img = isCellImage_(r[COL.SLIP - 1]);
      return { i, cells: r.map((v, j) => forms[i][j] || (isCellImage_(v) ? '' : v)), img: img ? r[COL.SLIP - 1] : null,
        slip: !!forms[i][COL.SLIP - 1] || img, blank: !id,
        day: d ? d.getTime() : Infinity, time: t instanceof Date ? t.getTime() : Infinity, id: num(id) };
    });
    rows.sort((a, b) => (a.blank - b.blank) ||
      (mode === 'day' ? (a.day - b.day) : 0) || (a.time - b.time) || (a.id[0] - b.id[0]) || (a.id[1] - b.id[1]) || (a.i - b.i));
    rng.setValues(rows.map(x => x.cells));
    rows.forEach((x, k) => { if (x.img) sh.getRange(k + 2, COL.SLIP).setValue(x.img); });   // slip previews (pictures in cells) move with their row
    rows.forEach((x, k) => sh.setRowHeight(k + 2, x.slip ? 80 : 21));
  } finally { lock.releaseLock(); }
  ss.toast(mode === 'day' ? 'เรียงตามวันรับแล้ว 📅' : 'เรียงตามเวลาสั่งแล้ว 🕒', sh.getName(), 4);
}

/* ---------------- ปิดรอบ: ย้ายออเดอร์ไป Orders Archive ---------------- */
// "9 Nov" → Date (ปีปัจจุบัน; ธ.ค. → ม.ค. ข้ามปีให้)
function dayDate_(v) {
  if (v instanceof Date) return v;
  const EN = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
  const TH = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
  const m = String(v).trim().match(/(\d{1,2})\s*(\S+)/); if (!m) return null;
  let mi = EN.indexOf(m[2].slice(0, 3).toLowerCase()); if (mi < 0) mi = TH.indexOf(m[2]); if (mi < 0) return null;
  const now = new Date(); let y = now.getFullYear();
  if (mi < now.getMonth() - 6) y++; else if (mi > now.getMonth() + 6) y--;
  return new Date(y, mi, +m[1]);
}
/* ---------------- โปรโมชั่นตามวันรับ ----------------
 * Settings: promo_on (ติ๊ก), promo_percent (เช่น 15), promo_label, promo_days
 * promo_days: "Fri" / "ศุกร์" = ทุกวันศุกร์ที่เปิดรับ · "13 Nov, 14 Nov" = เฉพาะวันนั้น · เว้นว่าง = ทุกวันรับ
 * ลดเฉพาะราคาเมนู ปัดขึ้นเป็นบาทเต็ม · add-on ราคาเดิม */
function promo_(s, rounds) {
  const pct = isOn_(s.promo_on) ? Math.min(90, Math.max(0, Number(s.promo_percent) || 0)) : 0;
  if (!pct) return null;
  const WD = { sun: 0, mon: 1, tue: 2, wed: 3, thu: 4, fri: 5, sat: 6, 'อาทิตย์': 0, 'จันทร์': 1, 'อังคาร': 2, 'พุธ': 3, 'พฤหัส': 4, 'ศุกร์': 5, 'เสาร์': 6 };
  const toks = String(s.promo_days || '').split(/[,\n]+/).map(t => t.trim()).filter(Boolean);
  let days = [];
  if (toks.length) days = rounds.filter(r => toks.some(t => {
    const k = t.toLowerCase().replace(/^วัน/, '').replace(/บดี$/, ''), wd = WD[k.slice(0, 3)] !== undefined ? WD[k.slice(0, 3)] : WD[k];
    if (wd !== undefined) { const d = dayDate_(r); return d && d.getDay() === wd; }
    const a = dayDate_(t), b = dayDate_(r); return a && b && a.getTime() === b.getTime();
  }));
  if (toks.length && !days.length) return null;          // the sale days aren't open this round
  return { pct, label: String(s.promo_label || '').trim() || ('ลดทุกเมนู ' + pct + '%'), days };
}
const isPromoDay_ = (p, r) => !!p && (!p.days.length || p.days.indexOf(r) >= 0);
const promoPrice_ = (p, price, r) => isPromoDay_(p, r) ? Math.ceil(price * (100 - p.pct) / 100 - 1e-9) : price;

// วันรับที่ผ่านไปแล้วสั่งไม่ได้ (เวลาไทย) · ของวันนี้สั่งได้ถึงเวลา cutoff (Settings: same_day_cutoff, ไม่มีแถว = 09:00, เว้นว่าง = ไม่ตัด)
function cutoffMin_(v) {
  if (v === undefined || v === null) v = '09:00';
  if (v instanceof Date) v = Utilities.formatDate(v, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm');   // typed as a time in the sheet
  const m = String(v).trim().match(/^(\d{1,2})(?:[:.](\d{2}))?$/);
  return m ? Math.min(24 * 60, +m[1] * 60 + +(m[2] || 0)) : null;      // blank / unreadable = no same-day cutoff
}
function isPastDay_(r, cutoff) {
  const d = dayDate_(r); if (!d) return false;
  const key = d.getFullYear() * 10000 + (d.getMonth() + 1) * 100 + d.getDate();
  const now = new Date(Date.now()), today = Number(Utilities.formatDate(now, 'Asia/Bangkok', 'yyyyMMdd'));
  if (key < today) return true;
  if (key === today && cutoff !== null && cutoff !== undefined) {
    const hm = Utilities.formatDate(now, 'Asia/Bangkok', 'HH:mm').split(':');
    return +hm[0] * 60 + +hm[1] >= cutoff;
  }
  return false;
}
const hhmm_ = m => ('0' + Math.floor(m / 60)).slice(-2) + ':' + ('0' + m % 60).slice(-2);
// ISO week, same as the sheet's ISOWEEKNUM → "Week 46 2026"
function weekLabel_(d) {
  const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const day = t.getUTCDay() || 7; t.setUTCDate(t.getUTCDate() + 4 - day);
  const y0 = new Date(Date.UTC(t.getUTCFullYear(), 0, 1));
  return 'Week ' + Math.ceil(((t - y0) / 86400000 + 1) / 7) + ' ' + t.getUTCFullYear();
}
function archiveOrders() {
  const ui = SpreadsheetApp.getUi(), ss = SpreadsheetApp.getActive();
  const ord = ss.getSheetByName(SH.ORD);
  const n = ord.getLastRow() - 1;
  if (n < 1) { ui.alert('แท็บ Orders ว่างอยู่แล้ว ไม่มีอะไรต้องเก็บ'); return; }
  const vals = ord.getRange(2, 1, n, ORD_HEAD.length).getValues().filter(r => String(at_(r, COL.ID)));
  const live = vals.filter(r => String(at_(r, COL.STATUS)) !== 'ยกเลิก');
  const ids = {}; live.forEach(r => ids[base_(at_(r, COL.ID))] = 1);
  const unpaid = {}; live.filter(r => String(at_(r, COL.PAY)) !== 'จ่ายแล้ว').forEach(r => unpaid[base_(at_(r, COL.ID))] = String(at_(r, COL.NAME)));
  const days = vals.map(r => dayDate_(at_(r, COL.DAY))).filter(Boolean).sort((a, b) => a - b);
  const week = weekLabel_(days[0] || new Date());
  const sales = live.reduce((t, r) => t + (Number(at_(r, COL.TOTAL)) || 0), 0);
  const warn = Object.keys(unpaid).length ? '\n\n⚠️ ยังไม่ขึ้น "จ่ายแล้ว" ' + Object.keys(unpaid).length + ' ออเดอร์: ' +
    Object.keys(unpaid).slice(0, 8).map(k => k + ' ' + unpaid[k]).join(', ') + (Object.keys(unpaid).length > 8 ? ' …' : '') +
    '\n(เก็บแล้วยังแก้สถานะได้ในแท็บ Orders Archive)' : '';
  const ok = ui.alert('📦 เก็บออเดอร์ ' + week,
    Object.keys(ids).length + ' ออเดอร์ · ' + live.reduce((t, r) => t + (Number(at_(r, COL.QTY)) || 0), 0) + ' ถ้วย · ยอดขาย ฿' + sales.toLocaleString('en-US') +
    '\n\nย้ายทุกแถวไปแท็บ "Orders Archive" แล้วล้างแท็บ Orders ให้ว่างรอสัปดาห์หน้า\nลูกค้ายังเห็นออเดอร์เดิมใน "ออเดอร์ของฉัน" เหมือนเดิม' + warn, ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;

  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const arch = archiveSheet_();
    const at = arch.getLastRow() + 1, rows = ord.getLastRow() - 1;
    arch.getRange(at, COL.DAY, rows, 1).setNumberFormat('@');
    ord.getRange(2, 1, rows, ORD_HEAD.length).copyTo(arch.getRange(at, 1));          // keeps slip pictures / links
    const stamp = new Date();
    arch.getRange(at, ORD_HEAD.length + 1, rows, 2).setValues(Array.from({ length: rows }, () => [week, stamp]));
    const slips = ord.getRange(2, COL.SLIP, rows, 1).getFormulas();
    slips.forEach((f, i) => { if (f[0]) arch.setRowHeight(at + i, 80); });
    ord.deleteRows(2, rows);
    // costs bought "this week" before the archive was named
    const cost = ss.getSheetByName(SH.COST);
    if (cost && cost.getLastRow() > 1) {
      const wk = cost.getRange(2, 2, cost.getLastRow() - 1, 1);
      wk.setValues(wk.getValues().map(x => [x[0] === 'รอบนี้' ? week : x[0]]));
    }
  } finally { lock.releaseLock(); }
  ss.toast('เก็บแล้ว ✅ ' + week + ' — แท็บ Orders พร้อมสำหรับรอบหน้า', 'Greek Shop', 8);
}
function archiveSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SH.ARCH);
  if (sh) return sh;
  sh = sheet_(ss, SH.ARCH, ARCH_HEAD, []);
  sh.getRange(2, COL.TIME, sh.getMaxRows() - 1, 1).setNumberFormat('dd/MM HH:mm');
  sh.getRange(2, ARCH_HEAD.length, sh.getMaxRows() - 1, 1).setNumberFormat('dd/MM/yyyy');
  sh.hideColumns(COL.TOKEN);
  return sh;
}

/* ---------------- ต้นทุน ---------------- */
// พิมพ์รายการแล้ว วันที่ / Week เติมให้เอง (แก้ได้)
function costAutofill_(r) {
  const sh = r.getSheet(), row = r.getRow();
  if (r.getColumn() !== 3 || !String(r.getValue()).trim()) return;
  if (!sh.getRange(row, 1).getValue()) sh.getRange(row, 1).setValue(new Date());
  if (!sh.getRange(row, 2).getValue()) sh.getRange(row, 2).setValue(currentWeek_());
}
// Week ของรอบที่ขายอยู่ = วันรับที่เร็วที่สุดในแท็บ Orders (ไม่มีออเดอร์ = รอบที่เปิดในแท็บ Rounds, ไม่มีอีก = วันนี้)
function currentWeek_() {
  const ss = SpreadsheetApp.getActive(), ord = ss.getSheetByName(SH.ORD);
  let days = ord && ord.getLastRow() > 1 ? ord.getRange(2, COL.DAY, ord.getLastRow() - 1, 1).getValues().map(x => dayDate_(x[0])).filter(Boolean) : [];
  if (!days.length) { try { days = getConfig().rounds.map(dayDate_).filter(Boolean); } catch (e) {} }
  days.sort((a, b) => a - b);
  return weekLabel_(days[0] || new Date());
}

/** ส่งเมนู/ราคา/ตัวเลือกให้หน้าแอป */
function getConfig() {
  const ss = SpreadsheetApp.getActive();
  const menu = rows_(ss.getSheetByName(SH.MENU)).filter(r => isOn_(r[6])).map(r => ({
    id: String(r[0]), name: String(r[1]), price: Number(r[2]) || 0, fullPrice: Number(r[2]) || 0,
    fruitPick: Number(r[3]) || 0, desc: String(r[4] || ''), emoji: String(r[5] || '🥣'),
    addons: r[7] === '' || r[7] === undefined ? true : isOn_(r[7])   // คอลัมน์ H "ใส่ Add-on ได้"
  }));
  const addons = rows_(ss.getSheetByName(SH.ADDON)).filter(r => isOn_(r[5])).map(r => ({
    id: String(r[0]), name: String(r[1]), price: Number(r[2]) || 0,
    type: String(r[3] || 'toggle'), max: Number(r[4]) || 1
  }));
  const fruits = rows_(ss.getSheetByName(SH.FRUIT)).filter(r => r[0]).map(r => ({
    name: String(r[0]), emoji: String(r[1] || ''), available: isOn_(r[2]),
    shape: String(r[3] || ''), color: /^#[0-9a-f]{6}$/i.test(String(r[4] || '').trim()) ? String(r[4]).trim() : ''
  }));
  const s = settings_();
  // รอบรับ: อ่านจากแท็บ Rounds (เฉพาะที่ติ๊ก "เปิดรับ"); ถ้าไม่มีแท็บนี้ ใช้ Settings แบบเดิม
  const roundSh = ss.getSheetByName(SH.ROUND);
  const roundRows = roundSh ? rows_(roundSh).filter(r => isOn_(r[1])) : null;
  const tz = ss.getSpreadsheetTimeZone();
  const roundName = v => v instanceof Date ? Utilities.formatDate(v, tz, 'd MMM') : String(v).trim();
  const rounds = roundRows ? roundRows.map(r => roundName(r[0])) : list_(s.pickup_rounds);
  const roundNotes = {};
  if (roundRows) roundRows.forEach(r => { if (r[2]) roundNotes[roundName(r[0])] = String(r[2]); });
  return {
    shopName: s.shop_name || 'Greek Yogurt',
    promo: promo_(s, rounds),
    open: isOn_(s.order_open) && (!roundRows || roundRows.length > 0) && (!rounds.length || rounds.some(r => !isPastDay_(r, cutoffMin_(s.same_day_cutoff)))),
    sameDayCutoff: cutoffMin_(s.same_day_cutoff) === null ? '' : hhmm_(cutoffMin_(s.same_day_cutoff)),  // ไม่มีรอบเปิด / ผ่านไปหมดแล้ว = ปิดรับ
    closedMsg: s.closed_message || 'ตอนนี้ปิดรับออเดอร์ชั่วคราว',
    rounds, roundNotes,
    delivery: list_(s.delivery_options).map(x => {
      const [name, fee] = x.split(':');
      return { name: name.trim(), fee: Number(fee) || 0 };
    }),
    payMethods: list_(s.pay_methods),
    payNowLabel: PAY_NOW, payLaterLabel: PAY_LATER,
    payLaterNote: String(s.pay_later_note || 'ผ่านขุนทองในกลุ่ม LINE'),
    bankAccount: String(s.bank_account || ''),
    accountName: String(s.account_name || ''),
    bankName: String(s.bank_name || ''),
    salePeriod: String(s.sale_period || '').trim() || period_(rounds),
    maxCups: limit_(s.max_cups_per_order, 10),
    lineGroupUrl: /^https:\/\/(line\.me|lin\.ee)\//.test(String(s.line_group_url || '').trim()) ? String(s.line_group_url).trim() : '',
    menu, addons, fruits
  };
}

/**
 * รับออเดอร์จากแอป — คำนวณราคาใหม่ฝั่ง server เสมอ (ไม่เชื่อราคาจากหน้าเว็บ)
 * order = { name, note, payLater, round?, items:[{menuId, fruits:[], addons:{SCP:2,HER:1}, qty, round}] }
 * แต่ละถ้วยมีวันรับของตัวเอง → ออเดอร์เดียวแยกเป็นแถวตามวันรับ (G0041-1, G0041-2) จ่าย/แนบสลิปครั้งเดียว
 */
function submitOrder(order) {
  const cfg = getConfig();
  if (!cfg.open) throw new Error(cfg.closedMsg);
  if (!order || !String(order.name || '').trim()) throw new Error('กรุณาใส่ชื่อ');
  if (!order.items || !order.items.length) throw new Error('ยังไม่มีรายการในตะกร้า');
  checkName_(order.name);
  const cups = order.items.reduce((t, it) => t + Math.max(1, parseInt(it.qty, 10) || 1), 0);
  if (cfg.maxCups && cups > cfg.maxCups) throw new Error('สั่งได้สูงสุด ' + cfg.maxCups + ' ถ้วยต่อออเดอร์ ถ้าอยากสั่งมากกว่านี้ทักร้านได้เลยนะ');

  const menuById = {}; cfg.menu.forEach(m => menuById[m.id] = m);
  const addonById = {}; cfg.addons.forEach(a => addonById[a.id] = a);
  const fruitOk = {}; cfg.fruits.forEach(f => { if (f.available) fruitOk[f.name] = true; });
  const allFruits = cfg.fruits.filter(f => f.available).map(f => f.name);

  const lines = order.items.map(it => {
    const m = menuById[it.menuId];
    if (!m) throw new Error('ไม่พบเมนู ' + it.menuId);
    const round = String(it.round || order.round || '');
    if (cfg.rounds.length && cfg.rounds.indexOf(round) < 0) throw new Error(m.name + ': กรุณาเลือกวันรับ (วันที่เลือกอาจปิดรับไปแล้ว)');
    if (isPastDay_(round, cfg.sameDayCutoff ? cutoffMin_(cfg.sameDayCutoff) : null))
      throw new Error(m.name + ': วันรับ ' + round + ' ปิดรับแล้ว' + (cfg.sameDayCutoff ? ' (ของวันเดียวกันสั่งได้ถึง ' + cfg.sameDayCutoff + ')' : '') + ' เลือกวันใหม่นะ');
    const qty = Math.max(1, Math.min(50, parseInt(it.qty, 10) || 1));
    const fruits = (it.fruits || []).filter(f => fruitOk[f]);
    if (fruits.length > m.fruitPick) throw new Error(m.name + ' เลือกผลไม้ได้สูงสุด ' + m.fruitPick + ' อย่าง');
    if (m.fruitPick > 0 && !fruits.length) throw new Error(m.name + ': กรุณาเลือกผลไม้อย่างน้อย 1 อย่าง');
    // "ไม่เอา" only means something for cups that normally get every fruit (e.g. Original 4 of 4)
    const excluded = m.fruitPick >= allFruits.length && fruits.length < allFruits.length
      ? allFruits.filter(f => fruits.indexOf(f) < 0) : [];

    let addonPrice = 0; const addonCount = {};
    Object.keys(m.addons ? (it.addons || {}) : {}).forEach(id => {
      const a = addonById[id]; if (!a) return;
      const n = Math.max(0, Math.min(a.max, parseInt(it.addons[id], 10) || 0));
      if (n) { addonCount[id] = n; addonPrice += n * a.price; }
    });
    const unit = promoPrice_(cfg.promo, m.price, round) + addonPrice;   // sale only on the promotion's pickup days
    return { m, round, qty, fruits, excluded, addonCount, addonPrice, unit, total: unit * qty };
  });

  // group by pickup day, in the order the days are listed in the Rounds tab
  const rounds = [];
  lines.forEach(l => { if (rounds.indexOf(l.round) < 0) rounds.push(l.round); });
  rounds.sort((x, y) => cfg.rounds.indexOf(x) - cfg.rounds.indexOf(y));

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let result, alert;
  try {
    const dev = deviceCheck_(order.device);          // per-phone limits (inside the lock so two taps can't race)
    const ss = SpreadsheetApp.getActive();
    const ordSh = ss.getSheetByName(SH.ORD);
    const base = 'G' + ('000' + nextOrderNo_(ordSh)).slice(-4);
    const now = new Date();
    const name = safeText_(String(order.name).trim().slice(0, 60));
    const saleDays = cfg.promo ? rounds.filter(r => isPromoDay_(cfg.promo, r)) : [];
    const note = safeText_((saleDays.length ? '[โปร -' + cfg.promo.pct + '% รับ ' + saleDays.join(', ') + '] ' : '') + String(order.note || '').replace(/[\u0000-\u0008\u000B-\u001F]/g, '').slice(0, 300));
    const payMethod = order.payLater ? PAY_LATER : PAY_NOW;
    const token = Utilities.getUuid();   // ใช้ยืนยันตอนแนบสลิป/ยกเลิก (กันคนอื่นมาแก้ออเดอร์เรา)
    // ยอดโอนทั้งออเดอร์ (ไม่นับวันที่ยกเลิก) โชว์ที่แถวแรกของออเดอร์ ไว้เทียบกับยอดในสลิป
    const sumF = '=SUMIFS(K:K,A:A,"' + base + '",R:R,"<>ยกเลิก")+SUMIFS(K:K,A:A,"' + base + '-*",R:R,"<>ยกเลิก")';

    const rows = [];
    const parts = rounds.map((round, k) => {
      const ls = lines.filter(l => l.round === round);
      const id = rounds.length > 1 ? base + '-' + (k + 1) : base;
      const total = ls.reduce((s, l) => s + l.total, 0);
      const cups = ls.reduce((s, l) => s + l.qty, 0);
      ls.forEach(l => {
        const first = !rows.length;
        const tops = Object.keys(l.addonCount).filter(a => a !== 'SCP')
          .map(a => addonById[a].name + (l.addonCount[a] > 1 ? ' x' + l.addonCount[a] : '')).join(', ');
        rows.push([id, now, name, round, l.m.name, l.fruits.join(', '), l.addonCount.SCP || '', tops,
          l.qty, l.unit, l.total, first ? note : '', payMethod, first ? sumF : '', PAY_STATUS[0], '', '',
          ORDER_STATUS[0], token]);
      });
      return { orderId: id, round, cups, total, summary: ls.map(l => l.m.name + (l.qty > 1 ? ' x' + l.qty : '')).join(', ') };
    });
    const r0 = ordSh.getLastRow() + 1;
    ordSh.getRange(r0, COL.DAY, rows.length, 1).setNumberFormat('@');      // "9 Nov" เก็บเป็นข้อความ ไม่ให้กลายเป็นวันที่
    ordSh.getRange(r0, 1, rows.length, ORD_HEAD.length).setValues(rows);

    result = { orderId: base, token, payLater: !!order.payLater, parts,
      total: parts.reduce((s, p) => s + p.total, 0), cups: parts.reduce((s, p) => s + p.cups, 0),
      bankAccount: cfg.bankAccount, accountName: cfg.accountName, bankName: cfg.bankName };
    alert = { base, name: name.replace(/^'/, ''), note: note.replace(/^'/, ''), payLater: !!order.payLater, rows, total: result.total, cups: result.cups };
    deviceRecord_(dev, base);
  } finally {
    lock.releaseLock();
  }
  alertOrder_(alert);
  return result;
}
/**
 * "ออเดอร์ของฉัน": แอปส่งเลขออเดอร์ + token ที่มือถือเครื่องนั้นจำไว้ → ส่งสถานะล่าสุดกลับ
 * (ต้องมี token ตรง คนอื่นจึงดูออเดอร์ของเราไม่ได้)
 */
function myOrders(req) {
  const asked = (req && req.orders || []).slice(0, 60);
  // __v บอกแอปว่าเป็นคำตอบเวอร์ชันนี้: ออเดอร์ไหนไม่มีในคำตอบ = ไม่อยู่ในชีตแล้ว (ทั้ง Orders และ Orders Archive) แอปจะลบออกจากประวัติในมือถือ
  if (!asked.length) return { __v: 2 };
  const vals = [].concat(...orderSources_().map(x => x.vals));
  if (!vals.length) return { __v: 2 };
  const out = { __v: 2 };
  asked.forEach(a => {
    const base = String(a.orderId || ''), tok = String(a.token || '');
    if (!base || !tok) return;
    const parts = {}, order = [];
    vals.forEach(r => {
      const id = String(at_(r, COL.ID));
      if (base_(id) !== base || String(at_(r, COL.TOKEN)) !== tok) return;
      if (!parts[id]) { parts[id] = { id, round: dayText_(at_(r, COL.DAY)), total: 0, pay: String(at_(r, COL.PAY)), status: '', slip: false, live: [] }; order.push(id); }
      const p = parts[id], st = String(at_(r, COL.STATUS));
      if (st !== 'ยกเลิก') { p.total += Number(at_(r, COL.TOTAL)) || 0; p.live.push(st); }
      if (String(at_(r, COL.SLIP))) p.slip = true;
    });
    if (order.length) out[base] = order.map(id => {
      const p = parts[id];
      p.status = p.live.length ? (p.live.indexOf('ใหม่') >= 0 ? 'ใหม่' : p.live[0]) : 'ยกเลิก';
      delete p.live; return p;
    });
  });
  return out;
}
function dayText_(v) { return v instanceof Date ? Utilities.formatDate(v, 'Asia/Bangkok', 'd MMM') : String(v); }

/**
 * ลูกค้ายกเลิกเอง: ได้เฉพาะแถวที่ยังไม่แนบสลิป / ยังไม่จ่าย / ยังไม่รับของ
 * req = { orderId, token, partId? }  (partId = เช่น G0041-2 เพื่อยกเลิกเฉพาะวันนั้น; ไม่ส่ง = ทั้งออเดอร์)
 */
function cancelOrder(req) {
  const base = String(req && req.orderId || ''), tok = String(req && req.token || '');
  if (!base || !tok) throw new Error('ไม่พบออเดอร์');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let cancelled = [];
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SH.ORD);
    const vals = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), ORD_HEAD.length).getValues();
    const rows = [];
    vals.forEach((r, i) => {
      const id = String(at_(r, COL.ID));
      if (base_(id) === base && String(at_(r, COL.TOKEN)) === tok && (!req.partId || id === String(req.partId))) rows.push({ row: i + 2, r });
    });
    if (!rows.length) {
      const arch = orderSources_().filter(x => x.sh.getName() === SH.ARCH)[0];
      if (arch && arch.vals.some(r => base_(at_(r, COL.ID)) === base && String(at_(r, COL.TOKEN)) === tok)) throw new Error('ออเดอร์นี้ปิดรอบไปแล้ว ยกเลิกเองไม่ได้ ทักร้านได้เลยนะ');
      throw new Error('ไม่พบออเดอร์');
    }
    const blocked = rows.find(x => String(at_(x.r, COL.PAY)) !== 'ยังไม่จ่าย' || String(at_(x.r, COL.SLIP)) || String(at_(x.r, COL.STATUS)) === 'ส่งแล้ว');
    if (blocked) throw new Error('ออเดอร์ ' + at_(blocked.r, COL.ID) + ' แนบสลิปหรือจ่ายแล้ว ยกเลิกเองไม่ได้ ทักร้านได้เลยนะ');
    const stamp = 'ลูกค้ายกเลิกเอง ' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM HH:mm');
    const stamped = {};
    cancelled = rows.filter(x => String(at_(x.r, COL.STATUS)) !== 'ยกเลิก').map(x => x.r);
    rows.forEach(x => {
      if (String(at_(x.r, COL.STATUS)) === 'ยกเลิก') return;
      sh.getRange(x.row, COL.STATUS).setValue('ยกเลิก');
      const id = String(at_(x.r, COL.ID));
      if (!stamped[id]) { stamped[id] = true;      // เขียนหมายเหตุครั้งเดียวต่อวันรับ
        sh.getRange(x.row, COL.NOTE).setValue([String(at_(x.r, COL.NOTE) || ''), stamp].filter(Boolean).join(' · ')); }
    });
  } finally {
    lock.releaseLock();
  }
  if (cancelled.length) alertCancel_(base, cancelled);
  if (!req.partId) deviceForget_(req.device, base);
  return myOrders({ orders: [{ orderId: base, token: tok }] });
}

/* ---------------- กันสแปม / กันกดซ้ำ ----------------
 * Settings: max_cups_per_order (10), max_orders_per_phone_per_day (3), order_cooldown_sec (30) · 0 = ไม่จำกัด
 * "มือถือ" = รหัสสุ่มที่แอปสร้างเก็บไว้ในเครื่อง (ล้างข้อมูลเบราว์เซอร์/เปลี่ยนเครื่อง = เครื่องใหม่) */
// a name/note starting with = + - @ would be read by Sheets as a formula (could send sheet data to an outside site)
// → store it as plain text instead (a leading ' is hidden by Sheets)
function safeText_(v) { v = String(v == null ? '' : v); return /^[=+\-@\t\r]/.test(v) ? "'" + v : v; }
function limit_(v, dflt) { const n = parseInt(v, 10); return isNaN(n) ? dflt : Math.max(0, n); }
function checkName_(raw) {
  const name = String(raw).trim();
  if ([...name].length < 2) throw new Error('ชื่อสั้นไป ใส่อย่างน้อย 2 ตัวอักษรนะ');
  if ([...name].length > 30) throw new Error('ชื่อยาวไป ใส่ไม่เกิน 30 ตัวอักษรนะ');
  if (/https?:|www\.|\.(com|net|org|io|co|th|ly|me)\b|line\.me|@\w{3,}/i.test(name)) throw new Error('ชื่อห้ามใส่ลิงก์นะ');
  if (/[<>{}\[\]\\|=;`$^~*]/.test(name) || /[\u0000-\u001F]/.test(name)) throw new Error('ชื่อมีสัญลักษณ์แปลกๆ ใช้แค่ตัวอักษร/ตัวเลขนะ');
}
function deviceKey_(id) {
  id = String(id || '');
  if (!/^[A-Za-z0-9-]{8,64}$/.test(id)) throw new Error('แอปเวอร์ชันเก่า ปิดแล้วเปิดใหม่ (หรือรีเฟรช) แล้วกดสั่งอีกครั้งนะ');
  return 'DEV_' + id;
}
const today_ = () => Utilities.formatDate(new Date(Date.now()), 'Asia/Bangkok', 'yyyyMMdd');
function deviceState_(key) {
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  const [day, last, list] = String(raw || '').split('|');
  return day === today_() ? { last: Number(last) || 0, orders: list ? list.split(',') : [] } : { last: day ? Number(last) || 0 : 0, orders: [] };
}
function deviceCheck_(id) {
  const key = deviceKey_(id), s = settings_(), st = deviceState_(key);
  const wait = limit_(s.order_cooldown_sec, 30), max = limit_(s.max_orders_per_phone_per_day, 3);
  const left = Math.ceil((st.last + wait * 1000 - Date.now()) / 1000);
  if (wait && left > 0) throw new Error('เพิ่งสั่งไปเมื่อกี้ รออีก ' + left + ' วินาทีแล้วค่อยกดใหม่นะ');
  if (max && st.orders.length >= max) throw new Error('วันนี้มือถือเครื่องนี้สั่งครบ ' + max + ' ออเดอร์แล้ว ถ้าอยากสั่งเพิ่มทักร้านได้เลยนะ');
  return key;
}
// runs after the order is saved: it must never turn a saved order into an error for the customer
function deviceRecord_(key, base) {
  try {
    const props = PropertiesService.getScriptProperties(), st = deviceState_(key);
    st.orders.push(base);
    props.setProperty(key, [today_(), Date.now(), st.orders.join(',')].join('|'));
    if (Math.random() < 0.05) {          // tidy up phones from earlier days now and then
      const all = props.getProperties(), d = today_();
      Object.keys(all).forEach(k => { if (k.indexOf('DEV_') === 0 && String(all[k]).split('|')[0] !== d) props.deleteProperty(k); });
    }
  } catch (e) { console.error('deviceRecord_: ' + e); }
}
function deviceForget_(id, base) {
  try {
    const key = deviceKey_(id), st = deviceState_(key);
    const i = st.orders.indexOf(base); if (i < 0) return;
    st.orders.splice(i, 1);
    PropertiesService.getScriptProperties().setProperty(key, [today_(), st.last, st.orders.join(',')].join('|'));
  } catch (e) {}
}

/** เลขออเดอร์วิ่งต่อจากเลขสูงสุดในแท็บ Orders */
function nextOrderNo_(ordSh) {
  const props = PropertiesService.getScriptProperties();
  let n = Number(props.getProperty('ORDER_SEQ')) || 0;
  if (!n && ordSh.getLastRow() > 1) {
    ordSh.getRange(2, 1, ordSh.getLastRow() - 1, 1).getValues().forEach(r => {
      const m = String(r[0]).match(/^G(\d+)/); if (m) n = Math.max(n, Number(m[1]));
    });
  }
  n += 1;
  props.setProperty('ORDER_SEQ', String(n));
  return n;
}

/**
 * แนบสลิป: เก็บรูปในโฟลเดอร์ Drive "Bright Greek Slips" แล้วโชว์รูปย่อในแท็บ Orders
 * req = { orderId, token, image: 'data:image/jpeg;base64,...' }
 */
function attachSlip(req) {
  const m = String(req && req.image || '').match(/^data:(image\/(jpeg|png|webp));base64,(.+)$/);
  if (!m) throw new Error('ไฟล์ต้องเป็นรูปภาพ');
  const bytes = Utilities.base64Decode(m[3]);
  if (bytes.length > 5 * 1024 * 1024) throw new Error('รูปใหญ่เกินไป');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let slipInfo = null;
  try {
    const base = String(req.orderId);
    let sh = null, rows = [];
    for (const src of orderSources_()) {
      src.vals.forEach((r, i) => { if (base_(at_(r, COL.ID)) === base) rows.push({ row: i + 2, vals: r }); });
      if (rows.length) { sh = src.sh; break; }
    }
    if (!rows.length) throw new Error('ไม่พบออเดอร์ ' + base);
    if (!req.token || rows.some(x => String(at_(x.vals, COL.TOKEN)) !== String(req.token))) throw new Error('ลิงก์แนบสลิปไม่ถูกต้อง');
    // the app reads the verification QR on the slip; one slip = one order (any phone)
    const ref = /^[A-Za-z0-9]{8,40}$/.test(String(req.slipRef || '')) ? String(req.slipRef) : '';
    const refKey = ref && 'SLIP_' + ref, props = PropertiesService.getScriptProperties();
    const usedBy = refKey && props.getProperty(refKey);
    if (usedBy && usedBy !== base) throw new Error('สลิปนี้เคยใช้แล้วกับออเดอร์ ' + usedBy + ' ถ้าโอนใหม่แล้ว เลือกสลิปใบใหม่นะ');

    slipInfo = { base, name: at_(rows[0].vals, COL.NAME), rows: rows.map(x => x.vals) };
    const ext = m[2] === 'jpeg' ? 'jpg' : m[2];
    const file = slipFolder_().createFile(
      Utilities.newBlob(bytes, m[1], base + '_' + at_(rows[0].vals, COL.NAME) + '_' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'MMdd-HHmm') + '.' + ext));
    const id = file.getId();          // the file stays private (only the shop's account can open it)
    const thumb = /^data:image\/jpeg;base64,[A-Za-z0-9+\/=]+$/.test(String(req.thumb || '')) && String(req.thumb).length < 300000 ? String(req.thumb) : '';
    slipInfo.thumb = thumb;
    let preview = null;
    if (thumb) try { preview = SpreadsheetApp.newCellImage().setSourceUrl(thumb).setAltTextTitle('สลิป ' + base).build(); } catch (e) { preview = null; }
    slipInfo.url = file.getUrl(); slipInfo.fileId = id;
    if (refKey) props.setProperty(refKey, base);
    rows.forEach(x => {
      if (String(at_(x.vals, COL.STATUS)) === 'ยกเลิก') return;
      const cell = sh.getRange(x.row, COL.SLIP);
      try { if (preview) cell.setValue(preview); else cell.setValue('📎 มีสลิป'); } catch (e) { cell.setValue('📎 มีสลิป'); }
      sh.getRange(x.row, COL.LINK).setFormula(
        '=HYPERLINK("' + file.getUrl() + '","' + (ref ? 'เปิดสลิป · ' + String(req.slipBank || '').replace(/["\\]/g, '').slice(0, 20) + ' …' + ref.slice(-6) : 'เปิดสลิป') + '")');
      if (at_(x.vals, COL.PAY) !== 'จ่ายแล้ว') sh.getRange(x.row, COL.PAY).setValue('ส่งสลิปแล้ว');
      sh.setRowHeight(x.row, 80);
    });
  } finally {
    lock.releaseLock();
  }
  alertSlip_(slipInfo);
  return { orderId: req.orderId };
}
/** กด Run ฟังก์ชันนี้ 1 ครั้งใน Apps Script เพื่ออนุญาตให้เก็บสลิปใน Google Drive */
function authorize() {
  const f = slipFolder_();
  Logger.log('โฟลเดอร์สลิปพร้อมแล้ว: ' + f.getUrl());
  Logger.log('อีเมลแจ้งเตือนเหลือส่งได้วันนี้: ' + MailApp.getRemainingDailyQuota() + ' ฉบับ · ส่งไปที่ ' + alertTo_());
}
/** กด Run เพื่อลองส่งอีเมลแจ้งเตือน 1 ฉบับ */
function testAlert() {
  sendAlert_('order', '🧪 ทดสอบแจ้งเตือน Bright Greek', '<p>ถ้าเห็นอีเมลนี้ แปลว่าแจ้งเตือนออเดอร์ใช้งานได้แล้ว ✅</p>');
}

/* ---------------- แจ้งเตือนเจ้าของร้านทางอีเมล ----------------
 * Settings: notify_email (เว้นว่าง = อีเมลเจ้าของสคริปต์), notify_events (order, slip, cancel — เว้นว่าง = ทั้งหมด, off = ปิด) */
function alertTo_() {
  const s = settings_(), to = String(s.notify_email || '').trim();
  return to || Session.getEffectiveUser().getEmail();
}
function sendAlert_(event, subject, html, inlineImages) {
  try {
    const s = settings_(), ev = String(s.notify_events || '').toLowerCase().trim();
    if (ev === 'off' || (ev && ev.split(/[\s,]+/).indexOf(event) < 0)) return;
    const to = alertTo_(); if (!to) return;
    const ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(SH.ORD);
    const link = ss.getUrl() + '#gid=' + sh.getSheetId();
    MailApp.sendEmail({ to, subject, name: 'Bright Greek', inlineImages: inlineImages || {}, htmlBody:
      '<div style="font-family:sans-serif;font-size:15px;color:#3B2A40;max-width:520px">' + html +
      '<p style="margin-top:18px"><a href="' + link + '" style="background:#FF8FAB;color:#3B2A40;padding:10px 16px;text-decoration:none;font-weight:bold;border:2px solid #3B2A40">เปิดแท็บ Orders</a></p></div>' });
  } catch (err) {
    console.error('alert failed: ' + err);      // แจ้งเตือนพลาดไม่ทำให้ออเดอร์พัง
  }
}
const h_ = v => String(v == null ? '' : v).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const baht_ = n => '฿' + Number(n || 0).toLocaleString('en-US');
function itemLines_(rows) {
  const days = {};
  rows.forEach(r => { const d = dayText_(at_(r, COL.DAY)); (days[d] = days[d] || []).push(r); });
  return Object.keys(days).map(d => '<p style="margin:12px 0 4px"><b>📅 รับวันที่ ' + h_(d) + '</b></p><ul style="margin:0;padding-left:20px">' +
    days[d].map(r => '<li><b>' + h_(at_(r, COL.MENU)) + ' × ' + at_(r, COL.QTY) + '</b> — ' + baht_(at_(r, COL.TOTAL)) +
      [at_(r, COL.FRUIT) ? '<br><span style="color:#7A6A7E">ผลไม้: ' + h_(at_(r, COL.FRUIT)) + '</span>' : '',
       at_(r, COL.SCOOP) ? '<br><span style="color:#7A6A7E">+' + at_(r, COL.SCOOP) + ' scoop</span>' : '',
       at_(r, COL.TOP) ? '<br><span style="color:#7A6A7E">' + h_(at_(r, COL.TOP)) + '</span>' : ''].join('') + '</li>').join('') + '</ul>').join('');
}
function alertOrder_(a) {
  if (!a) return;
  sendAlert_('order', '🛎️ ออเดอร์ใหม่ ' + a.base + ' · ' + a.name + ' · ' + a.cups + ' ถ้วย · ' + baht_(a.total),
    '<h2 style="margin:0 0 6px">🛎️ ออเดอร์ใหม่ ' + h_(a.base) + '</h2>' +
    '<p style="margin:0">👤 <b>' + h_(a.name) + '</b> · ' + a.cups + ' ถ้วย · <b>' + baht_(a.total) + '</b></p>' +
    '<p style="margin:4px 0">💳 ' + (a.payLater ? 'จ่ายตอนรับ (ขุนทอง)' : 'จ่ายเลย — รอสลิป') + '</p>' +
    (a.note ? '<p style="margin:4px 0">📝 ' + h_(a.note) + '</p>' : '') + itemLines_(a.rows));
}
function alertSlip_(x) {
  if (!x) return;
  const live = x.rows.filter(r => String(at_(r, COL.STATUS)) !== 'ยกเลิก');
  const total = live.reduce((t, r) => t + (Number(at_(r, COL.TOTAL)) || 0), 0);
  sendAlert_('slip', '📎 สลิปเข้า ' + x.base + ' · ' + x.name + ' · ' + baht_(total),
    '<h2 style="margin:0 0 6px">📎 ลูกค้าแนบสลิปแล้ว ' + h_(x.base) + '</h2>' +
    '<p style="margin:0">👤 <b>' + h_(x.name) + '</b> · ยอดที่ต้องได้ <b>' + baht_(total) + '</b></p>' +
    (x.thumb ? '<p><a href="' + x.url + '"><img src="cid:slip" alt="สลิป" style="max-width:240px;border:2px solid #3B2A40"></a></p>' : '<p><a href="' + x.url + '">เปิดสลิป</a></p>') +
    '<p style="color:#7A6A7E">เช็คยอดในแอปธนาคารก่อน แล้วเปลี่ยน "สถานะจ่าย" เป็น "จ่ายแล้ว" ในแท็บ Orders</p>',
    x.thumb ? { slip: Utilities.newBlob(Utilities.base64Decode(x.thumb.split(',')[1]), 'image/jpeg', 'slip.jpg') } : null);
}
function alertCancel_(base, rows) {
  const name = at_(rows[0], COL.NAME);
  sendAlert_('cancel', '❌ ยกเลิก ' + base + ' · ' + name,
    '<h2 style="margin:0 0 6px">❌ ลูกค้ายกเลิกเอง ' + h_(base) + '</h2><p style="margin:0">👤 <b>' + h_(name) + '</b></p>' + itemLines_(rows) +
    '<p style="color:#7A6A7E">ไม่ต้องเตรียมรายการนี้แล้ว</p>');
}
/** กด Run ครั้งเดียว: ปิดลิงก์สาธารณะของสลิปเก่าทุกใบในโฟลเดอร์ "Bright Greek Slips" (ให้เปิดได้แค่บัญชีร้าน) */
function lockSlips() {
  const folder = slipFolder_(), it = folder.getFiles(); let n = 0;
  try { folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); } catch (e) {}
  while (it.hasNext()) { const f = it.next();
    if (f.getSharingAccess() !== DriveApp.Access.PRIVATE) { f.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); n++; } }
  Logger.log('ปิดลิงก์สาธารณะแล้ว ' + n + ' ไฟล์ · โฟลเดอร์สลิป: ' + folder.getUrl());
}
function slipFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SLIP_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const it = DriveApp.getFoldersByName(SLIP_FOLDER);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(SLIP_FOLDER);
  props.setProperty('SLIP_FOLDER_ID', folder.getId());
  return folder;
}

/**
 * เจ้าของร้านแก้ในแท็บ Orders ช่องเดียว แล้วแถวอื่นของออเดอร์เดียวกันเปลี่ยนตาม
 *   สถานะจ่าย   → ทุกแถวของออเดอร์ (จ่ายทีเดียวทั้งออเดอร์)
 *   สถานะออเดอร์ → ทุกแถวของวันรับนั้น (เช่น G0041-2)
 */
function onEdit(e) {
  const r = e && e.range;
  if (!r || r.getRow() < 2 || r.getNumRows() !== 1 || r.getNumColumns() !== 1) return;
  const name = r.getSheet().getName();
  if (name === SH.COST) return costAutofill_(r);
  if (name !== SH.ORD && name !== SH.ARCH) return;
  const col = r.getColumn();
  if (col !== COL.PAY && col !== COL.STATUS) return;
  const sh = r.getSheet(), v = r.getValue();
  const id = String(sh.getRange(r.getRow(), COL.ID).getValue());
  if (!id) return;
  const ids = sh.getRange(2, COL.ID, sh.getLastRow() - 1, 1).getValues();
  ids.forEach((x, i) => {
    const row = i + 2, other = String(x[0]);
    if (row === r.getRow()) return;
    if (col === COL.PAY ? base_(other) === base_(id) : other === id) sh.getRange(row, col).setValue(v);
  });
}

/* ---------------- Setup ---------------- */

function setup() {
  const ss = SpreadsheetApp.getActive();

  sheet_(ss, SH.ROUND, ['วันที่รับ (เช่น 6 Nov)', 'เปิดรับ', 'หมายเหตุ (ไม่บังคับ โชว์ใต้ปุ่ม)'], [
    ['6 Nov', true, '']
  ]);
  sheet_(ss, SH.MENU, ['ID', 'ชื่อเมนู', 'ราคา', 'เลือกผลไม้ได้ (อย่าง)', 'คำอธิบาย', 'อีโมจิ', 'เปิดขาย', 'ใส่ Add-on ได้'], [
    ['ORI', 'Original', 80, 4, 'กรีกโยเกิร์ต + ผลไม้ 4 อย่าง', '🍓', true, true],
    ['BIS', 'Biscoff', 90, 2, 'ครัมบ์ Biscoff + ผลไม้ 2 อย่าง', '🍪', true, true],
    ['ORE', 'Oreo', 90, 4, 'Oreo + ผลไม้ 4 อย่าง', '🖤', true, true],
    ['BAN', 'Greek Banoffee', 100, 0, 'กล้วย + คาราเมล', '🍌', true, false],
    ['NOF', 'Greek no fruits', 40, 0, 'กรีกเปล่า ไม่ใส่ผลไม้', '🥛', true, true]
  ]);
  sheet_(ss, SH.ADDON, ['ID', 'ชื่อ', 'ราคา', 'แบบ (count/toggle)', 'สูงสุด', 'เปิดขาย'], [
    ['SCP', 'เพิ่ม Scoop', 10, 'count', 5, true],
    ['HER', 'Hershey syrup', 10, 'toggle', 1, true]
  ]);
  sheet_(ss, SH.FRUIT, ['ผลไม้', 'อีโมจิ', 'มีของ', 'รูปทรง (ไม่บังคับ)', 'สี เช่น #FF7A8A (ไม่บังคับ)'], [
    ['กล้วยหอม', '🍌', true],
    ['องุ่น', '🍇', true],
    ['บลูเบอร์รี่', '🫐', true],
    ['ไชน์มัสคัท', '🍏', true]
  ]);
  sheet_(ss, SH.SET, ['key', 'value', 'คำอธิบาย'], [
    ['shop_name', 'Bright Greek', 'ชื่อร้านที่โชว์บนแอป'],
    ['order_open', true, 'TRUE = เปิดรับออเดอร์, FALSE = ปิด'],
    ['closed_message', 'รอบนี้ปิดรับแล้วน้า รอรอบหน้า 🙏', 'ข้อความตอนปิดรับ'],
    ['pickup_rounds', '', 'ย้ายไปแท็บ Rounds แล้ว'],
    ['delivery_options', '', 'เว้นว่าง = ไม่ถามวิธีรับ'],
    ['pay_methods', '', 'เว้นว่าง = จ่าย QR / โอนกสิกรอย่างเดียว'],
    ['bank_account', '235-3-39428-5', 'เลขบัญชี (ลูกค้ากดคัดลอกได้)'],
    ['account_name', 'นาย สหรัฐ กาญจนางกูรพันธุ์', 'ชื่อบัญชี'],
    ['bank_name', 'ธนาคารกสิกรไทย', 'ชื่อธนาคาร'],
    ['sale_period', '', 'เว้นว่าง = คำนวณจากแท็บ Rounds อัตโนมัติ (พิมพ์เองได้ถ้าอยากใช้ข้อความอื่น)'],
    ['pay_later_note', 'ผ่านขุนทองในกลุ่ม LINE', 'ข้อความใต้ตัวเลือก "จ่ายตอนรับ"'],
    ['line_group_url', '', 'ลิงก์เชิญกลุ่ม LINE (https://line.me/...) เว้นว่าง = ให้ลูกค้าเลือกแชทเอง'],
    ['notify_email', '', 'อีเมลรับแจ้งเตือน (เว้นว่าง = อีเมลเจ้าของสคริปต์)'],
    ['same_day_cutoff', '09:00', 'สั่งของ "วันนี้" ได้ถึงกี่โมง (เวลาไทย) · เว้นว่าง = ไม่ตัด'],
    ['promo_on', false, 'ติ๊ก = เปิดโปรโมชั่น (ทุกเมนูลดราคา, add-on ราคาเดิม)'],
    ['promo_percent', 15, 'ลดกี่ % (ปัดขึ้นเป็นบาทเต็ม เช่น 90 ลด 15% = 77)'],
    ['promo_label', 'Friday Clearance', 'ชื่อโปรที่โชว์ในแอป'],
    ['promo_days', 'Fri', 'ลดเฉพาะวันรับไหน: Fri / ศุกร์ = ทุกศุกร์ · 13 Nov = วันนั้น · คั่นด้วย , ได้ · เว้นว่าง = ทุกวัน'],
    ['max_cups_per_order', 10, 'สั่งได้สูงสุดกี่ถ้วยต่อออเดอร์ (0 = ไม่จำกัด)'],
    ['max_orders_per_phone_per_day', 3, 'มือถือเครื่องเดียวสั่งได้กี่ออเดอร์ต่อวัน (ยกเลิกเองไม่นับ · 0 = ไม่จำกัด)'],
    ['order_cooldown_sec', 30, 'ต้องรอกี่วินาทีก่อนสั่งออเดอร์ถัดไปจากเครื่องเดิม (0 = ไม่ต้องรอ)'],
    ['notify_events', 'order, slip, cancel', 'แจ้งเตือนเมื่อ: order = ออเดอร์ใหม่, slip = แนบสลิป, cancel = ลูกค้ายกเลิก · off = ปิดทั้งหมด']
  ]);
  ['Rounds', 'Menu', 'AddOns', 'Fruits'].forEach(n => {
    const sh = ss.getSheetByName(n);
    const col = n === 'Rounds' ? 2 : n === 'Menu' ? 7 : n === 'AddOns' ? 6 : 3;
    sh.getRange(2, col, sh.getMaxRows() - 1, 1).insertCheckboxes();
  });

  const ord = sheet_(ss, SH.ORD, ORD_HEAD, []);
  formatOrders_(ord);

  // หน้า Summary ใช้สูตรอัตโนมัติอยู่แล้ว สร้างใหม่เฉพาะตอนที่ยังไม่มีแท็บนี้
  if (!ss.getSheetByName(SH.SUM)) buildSummary();
  SpreadsheetApp.getActive().toast('ตั้งค่าเสร็จแล้ว ✅ ต่อไป Deploy เป็น Web app', 'Greek Shop', 6);
}

/** หน้าตาแท็บ Orders: dropdown สถานะ, สีตามสถานะ, ซ่อน token */
function formatOrders_(ord) {
  const n = ord.getMaxRows() - 1;
  ord.getRange(2, COL.PAY, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(PAY_STATUS).build());
  ord.getRange(2, COL.STATUS, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ORDER_STATUS).build());
  ord.getRange(2, COL.TIME, n, 1).setNumberFormat('dd/MM HH:mm');
  ord.getRange(2, COL.DAY, n, 1).setNumberFormat('@');
  ord.getRange(2, COL.UNIT, n, 2).setNumberFormat('#,##0');
  ord.getRange(2, COL.SUM, n, 1).setNumberFormat('#,##0');
  const all = ord.getRange(2, 1, n, COL.STATUS);
  ord.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$R2="ยกเลิก"')
      .setBackground('#e0e0e0').setFontColor('#888888').setStrikethrough(true).setRanges([all]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$O2="จ่ายแล้ว"')
      .setBackground('#b7f5b0').setRanges([all]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$O2="ส่งสลิปแล้ว"')
      .setBackground('#d6ecff').setRanges([all]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($A2<>"",$O2="ยังไม่จ่าย")')
      .setBackground('#fff3cd').setRanges([all]).build()
  ]);
  ord.hideColumns(COL.TOKEN);
}

/** สูตรหน้า Summary (อ่านจากแท็บ Orders แท็บเดียว) — เลือกวันรับที่ช่อง B1 */
function summaryFormulas_() {
  const R = '((($B$1="ทุกรอบ")+(TO_TEXT(Orders!$D$2:$D)=TO_TEXT($B$1)))>0)';   // วันรับที่เลือก
  const NC = '(Orders!$R$2:$R<>"ยกเลิก")*(Orders!$A$2:$A<>"")';                // ไม่นับที่ยกเลิก
  const sp = x => '=SUMPRODUCT(' + R + '*' + NC + '*' + x + ')';
  const day = '"&IF($B$1="ทุกรอบ",""," and D = \'"&$B$1&"\'")&"';
  return {
    menuCups: a => '=IF(' + a + '="","",SUMPRODUCT((Orders!$E$2:$E=' + a + ')*' + R + '*' + NC + '*Orders!$I$2:$I))',
    menuMoney: a => '=IF(' + a + '="","",SUMPRODUCT((Orders!$E$2:$E=' + a + ')*' + R + '*' + NC + '*Orders!$K$2:$K))',
    fruit: a => '=IF(' + a + '="","",SUMPRODUCT(ISNUMBER(SEARCH(' + a + ',Orders!$F$2:$F))*' + R + '*' + NC + '*Orders!$I$2:$I))',
    scoop: sp('Orders!$G$2:$G*Orders!$I$2:$I'),
    hershey: sp('ISNUMBER(SEARCH("Hershey",Orders!$H$2:$H))*Orders!$I$2:$I'),
    orders: '=IFERROR(ROWS(UNIQUE(FILTER(ARRAYFORMULA(REGEXEXTRACT(Orders!$A$2:$A&"","^G\\d+")),' + R + ',' + NC + '))),0)',
    sales: sp('Orders!$K$2:$K'),
    cups: sp('Orders!$I$2:$I'),
    paid: sp('(Orders!$O$2:$O="จ่ายแล้ว")*Orders!$K$2:$K'),
    unpaid: sp('(Orders!$O$2:$O="ยังไม่จ่าย")*Orders!$K$2:$K'),
    slip: sp('(Orders!$O$2:$O="ส่งสลิปแล้ว")*Orders!$K$2:$K'),
    unpaidList: '=IFERROR(QUERY(Orders!A2:R,"select C, sum(K), A where A is not null and O = \'ยังไม่จ่าย\' and R <> \'ยกเลิก\'' + day + ' group by C, A label sum(K) \'\'",0),"— ไม่มี —")',
    khunThong: '=IFERROR(LET(q,QUERY(Orders!A2:R,"select C, sum(K) where A is not null and R <> \'ยกเลิก\' and O <> \'จ่ายแล้ว\' and M = \'' + PAY_LATER + '\'' + day + ' group by C label sum(K) \'\'",0),TEXTJOIN(CHAR(10),TRUE,ARRAYFORMULA(INDEX(q,,1)&" ฿"&INDEX(q,,2)))),"— ไม่มีคนจ่ายตอนรับ —")'
  };
}

/** สร้างหน้า Summary ใหม่ทั้งหน้า */
function buildSummary() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SH.SUM);
  if (!sh) sh = ss.insertSheet(SH.SUM);
  sh.clear(); sh.clearConditionalFormatRules();
  const f = summaryFormulas_();
  const put = (a1, v) => { const c = sh.getRange(a1); String(v).charAt(0) === '=' ? c.setFormula(v) : c.setValue(v); return c; };

  put('A1', 'เลือกรอบ ▶').setFontWeight('bold');
  put('F1', 'ทุกรอบ'); put('F2', '=FILTER(Rounds!A2:A,Rounds!A2:A<>"")');
  put('B1', 'ทุกรอบ').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(sh.getRange('F1:F60')).build())
    .setBackground('#fff59d').setFontWeight('bold');
  const head = (row, t, b, c) => { sh.getRange(row, 1, 1, 3).setValues([[t, b || '', c || '']]).setFontWeight('bold').setBackground('#c9daf8'); };

  head(3, '🥣 จำนวนแต่ละเมนู', 'ถ้วย', 'ยอดเงิน');
  for (let i = 0; i < 10; i++) {
    const r = 4 + i;
    put('A' + r, '=IF(Menu!B' + (2 + i) + '="","",Menu!B' + (2 + i) + ')');
    put('B' + r, f.menuCups('A' + r)); put('C' + r, f.menuMoney('A' + r));
  }
  sh.getRange('A14:C14').setValues([['รวม', '=SUM(B4:B13)', '=SUM(C4:C13)']]).setFontWeight('bold');

  head(16, '🍇 ผลไม้ที่ต้องเตรียม', 'ถ้วย');
  for (let i = 0; i < 10; i++) {
    const r = 17 + i;
    put('A' + r, '=IF(Fruits!A' + (2 + i) + '="","",Fruits!A' + (2 + i) + ')'); put('B' + r, f.fruit('A' + r));
  }
  head(28, '➕ Add-on');
  put('A29', 'Scoop เพิ่ม (scoop)'); put('B29', f.scoop);
  put('A30', 'Hershey (ถ้วย)'); put('B30', f.hershey);

  head(32, '💰 ยอดเงิน (บาท)');
  [['จำนวนออเดอร์', f.orders], ['ยอดขายรวม', f.sales], ['จำนวนถ้วยรวม', f.cups], ['✅ จ่ายแล้ว', f.paid],
   ['⏳ ยังไม่จ่าย', f.unpaid], ['📎 ส่งสลิปแล้ว (รอเช็ค)', f.slip]].forEach((x, i) => { put('A' + (33 + i), x[0]); put('B' + (33 + i), x[1]); });
  head(39, '⏳ ค้างจ่าย: ชื่อ', 'ยอด', 'OrderID');
  put('A40', f.unpaidList);
  put('A42', '📋 ยอดรายคนที่ "จ่ายตอนรับ" (copy ไปสร้างบิลในขุนทอง)').setFontWeight('bold');
  put('A43', f.khunThong);

  sh.setColumnWidth(1, 220); sh.setColumnWidth(2, 120); sh.setColumnWidth(3, 120);
  sh.getRange('B33:B38').setNumberFormat('#,##0'); sh.getRange('C4:C14').setNumberFormat('#,##0');
}

/* ---------------- helpers ---------------- */

function sheet_(ss, name, header, defaults) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, header.length).setValues([header])
      .setFontWeight('bold').setBackground('#f9a825').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    if (defaults.length) sh.getRange(2, 1, defaults.length, defaults[0].length).setValues(defaults);
  }
  return sh;
}
function rows_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues().filter(r => r[0] !== '');
}
function settings_() {
  const o = {};
  rows_(SpreadsheetApp.getActive().getSheetByName(SH.SET)).forEach(r => o[String(r[0]).trim()] = r[1]);
  return o;
}
/** วันที่ที่เปิดรับ → ป้าย "รอบส่ง": ["6 Nov","7 Nov","11 Nov","12 Nov"] → "6-7, 11-12 Nov 2026" (วันที่ไม่ติดกันจะแยกช่วงให้) */
function period_(rounds) {
  const EN = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
  const TH = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
  const now = new Date(), y0 = now.getFullYear();
  const days = [];
  for (const r of rounds) {
    const m = String(r).trim().match(/^(\d{1,2})\s*(\S+)/);
    if (!m) return '';
    let mi = EN.indexOf(m[2].slice(0, 3).toLowerCase());
    if (mi < 0) mi = TH.indexOf(m[2]);
    if (mi < 0) return '';
    const y = mi < now.getMonth() - 6 ? y0 + 1 : y0;     // ธ.ค. → ม.ค. ข้ามปี
    days.push({ d: +m[1], mi, label: m[2], t: new Date(y, mi, +m[1]).getTime(), y });
  }
  if (!days.length) return '';
  days.sort((a, b) => a.t - b.t);
  const segs = [];
  days.forEach(x => {
    const last = segs[segs.length - 1];
    if (last && x.t - last.end.t <= 86400000 * 1.5) last.end = x; else segs.push({ start: x, end: x });
  });
  const out = segs.map((g, i) => {
    const a = g.start, b = g.end, next = segs[i + 1];
    let txt = a.mi === b.mi ? (a.d === b.d ? '' + a.d : a.d + '-' + b.d) : a.d + ' ' + a.label + ' - ' + b.d;
    if (!next || next.start.mi !== b.mi) txt += ' ' + b.label;
    return txt;
  });
  return out.join(', ') + ' ' + days[days.length - 1].y;
}
function list_(v) { return String(v || '').split(',').map(s => s.trim()).filter(Boolean); }
function isOn_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
