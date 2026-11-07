/**
 * Greek Yogurt Order App — Google Apps Script backend
 *
 * Sheets created by setup():
 *   Menu       เมนู + ราคา + จำนวนผลไม้ที่เลือกได้   (แก้ราคาที่นี่ → แอปเปลี่ยนตาม)
 *   AddOns     ท็อปปิ้งเพิ่ม (scoop, hershey)
 *   Fruits     รายการผลไม้ + เปิด/ปิด ของหมด
 *   Settings   รอบรับ, วิธีรับ/ค่าส่ง, PromptPay, เปิด/ปิดรับออเดอร์
 *   Orders     1 แถว = 1 ออเดอร์ (ยอดรวม, สถานะจ่าย, สถานะออเดอร์)
 *   OrderItems 1 แถว = 1 รายการ (ละเอียด: ผลไม้, scoop, hershey, ราคาแยก)
 *   Summary    สรุปต่อรอบ: จำนวนแต่ละเมนู, ผลไม้ที่ต้องเตรียม, ยอดเงิน จ่ายแล้ว/ค้าง
 */

const SH = {
  ROUND: 'Rounds', MENU: 'Menu', ADDON: 'AddOns', FRUIT: 'Fruits', SET: 'Settings',
  ORD: 'Orders', ITEM: 'OrderItems', SUM: 'Summary'
};
const PAY_STATUS = ['ยังไม่จ่าย', 'ส่งสลิปแล้ว', 'จ่ายแล้ว'];
const PAY_NOW = 'จ่ายเลย (แนบสลิป)';
const PAY_LATER = 'จ่ายตอนรับ (ขุนทอง)';
const SLIP_FOLDER = 'Bright Greek Slips';
const ORDER_STATUS = ['ใหม่', 'ทำแล้ว', 'ส่งแล้ว', 'ยกเลิก'];

const ORD_HEAD = ['OrderID', 'เวลาสั่ง', 'ชื่อ', 'ติดต่อ', 'รอบรับ', 'วิธีรับ', 'ค่าส่ง',
  'ยอดอาหาร', 'ยอดรวม', 'ช่องทางจ่าย', 'สถานะจ่าย', 'สถานะออเดอร์', 'หมายเหตุ', 'รายการ (สรุป)', 'จำนวนถ้วย',
  'สลิป', 'ลิงก์สลิป', 'token'];
const ITEM_HEAD = ['OrderID', 'เวลาสั่ง', 'รอบรับ', 'ชื่อ', 'เมนู', 'ผลไม้ที่เลือก', 'ไม่เอา',
  'Scoop เพิ่ม', 'Hershey', 'ราคาเมนู', 'ราคา Add-on', 'ราคา/ถ้วย', 'จำนวน', 'รวม', 'Remark',
  'สถานะจ่าย', 'สถานะออเดอร์'];

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
    .addItem('ตั้งค่าครั้งแรก (สร้างชีตที่ยังไม่มี)', 'setup')
    .addToUi();
}

/** ส่งเมนู/ราคา/ตัวเลือกให้หน้าแอป */
function getConfig() {
  const ss = SpreadsheetApp.getActive();
  const menu = rows_(ss.getSheetByName(SH.MENU)).filter(r => isOn_(r[6])).map(r => ({
    id: String(r[0]), name: String(r[1]), price: Number(r[2]) || 0,
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
    open: isOn_(s.order_open) && (!roundRows || roundRows.length > 0),  // ไม่มีรอบเปิด = ปิดรับ
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

  const menuById = {}; cfg.menu.forEach(m => menuById[m.id] = m);
  const addonById = {}; cfg.addons.forEach(a => addonById[a.id] = a);
  const fruitOk = {}; cfg.fruits.forEach(f => { if (f.available) fruitOk[f.name] = true; });
  const allFruits = cfg.fruits.filter(f => f.available).map(f => f.name);
  const deliv = cfg.delivery.find(d => d.name === order.delivery) || cfg.delivery[0] || { name: '-', fee: 0 };

  const lines = order.items.map(it => {
    const m = menuById[it.menuId];
    if (!m) throw new Error('ไม่พบเมนู ' + it.menuId);
    const round = String(it.round || order.round || '');
    if (cfg.rounds.length && cfg.rounds.indexOf(round) < 0) throw new Error(m.name + ': กรุณาเลือกวันรับ (วันที่เลือกอาจปิดรับไปแล้ว)');
    const qty = Math.max(1, Math.min(50, parseInt(it.qty, 10) || 1));
    const fruits = (it.fruits || []).filter(f => fruitOk[f]);
    if (fruits.length > m.fruitPick) throw new Error(m.name + ' เลือกผลไม้ได้สูงสุด ' + m.fruitPick + ' อย่าง');
    if (m.fruitPick > 0 && !fruits.length) throw new Error(m.name + ': กรุณาเลือกผลไม้อย่างน้อย 1 อย่าง');
    const excluded = m.fruitPick > 0 && fruits.length < allFruits.length
      ? allFruits.filter(f => fruits.indexOf(f) < 0) : [];

    let addonPrice = 0; const addonCount = {};
    Object.keys(m.addons ? (it.addons || {}) : {}).forEach(id => {
      const a = addonById[id]; if (!a) return;
      const n = Math.max(0, Math.min(a.max, parseInt(it.addons[id], 10) || 0));
      if (n) { addonCount[id] = n; addonPrice += n * a.price; }
    });
    const unit = m.price + addonPrice;
    return { m, round, qty, fruits, excluded, addonCount, addonPrice, unit, total: unit * qty };
  });

  // group by pickup day, in the order the days are listed in the Rounds tab
  const rounds = [];
  lines.forEach(l => { if (rounds.indexOf(l.round) < 0) rounds.push(l.round); });
  rounds.sort((x, y) => cfg.rounds.indexOf(x) - cfg.rounds.indexOf(y));

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActive();
    const ordSh = ss.getSheetByName(SH.ORD), itemSh = ss.getSheetByName(SH.ITEM);
    const base = 'G' + ('000' + nextOrderNo_(ordSh)).slice(-4);
    const now = new Date();
    const name = String(order.name).trim().slice(0, 60);
    const payMethod = order.payLater ? PAY_LATER : PAY_NOW;
    const token = Utilities.getUuid();   // ใช้ยืนยันตอนแนบสลิป (กันคนอื่นแนบมั่ว)
    const payF = '=IFERROR(VLOOKUP(INDIRECT("A"&ROW()),Orders!A:L,11,FALSE),"")';
    const statF = '=IFERROR(VLOOKUP(INDIRECT("A"&ROW()),Orders!A:L,12,FALSE),"")';

    const parts = rounds.map((round, k) => {
      const ls = lines.filter(l => l.round === round);
      const id = rounds.length > 1 ? base + '-' + (k + 1) : base;
      const food = ls.reduce((s, l) => s + l.total, 0);
      const cups = ls.reduce((s, l) => s + l.qty, 0);
      const fee = k === 0 ? deliv.fee : 0;
      const summary = ls.map(l => l.m.name + (l.qty > 1 ? ' x' + l.qty : '')).join(', ');
      ordSh.appendRow([id, now, name, String(order.contact || '').slice(0, 80), round,
        deliv.name, fee, food, food + fee, payMethod, PAY_STATUS[0], ORDER_STATUS[0],
        String(order.note || '').slice(0, 300), summary, cups, '', '', token]);
      const itemRows = ls.map(l => {
        const scoop = l.addonCount.SCP || 0, her = l.addonCount.HER || 0;
        return [id, now, round, name, l.m.name, l.fruits.join(', '), l.excluded.join(', '),
          scoop, her ? 'ใช่' : '', l.m.price, l.addonPrice, l.unit, l.qty, l.total,
          remark_(round, l, addonById), payF, statF];
      });
      itemSh.getRange(itemSh.getLastRow() + 1, 1, itemRows.length, ITEM_HEAD.length).setValues(itemRows);
      return { orderId: id, round, cups, total: food + fee, summary };
    });

    return { orderId: base, token, payLater: !!order.payLater, parts,
      total: parts.reduce((s, p) => s + p.total, 0), cups: parts.reduce((s, p) => s + p.cups, 0),
      bankAccount: cfg.bankAccount, accountName: cfg.accountName, bankName: cfg.bankName };
  } finally {
    lock.releaseLock();
  }
}
/**
 * "ออเดอร์ของฉัน": แอปส่งเลขออเดอร์ + token ที่มือถือเครื่องนั้นจำไว้ → ส่งสถานะล่าสุดกลับ
 * (ต้องมี token ตรง คนอื่นจึงดูออเดอร์ของเราไม่ได้)
 */
function myOrders(req) {
  const asked = (req && req.orders || []).slice(0, 60);
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.ORD);
  if (!asked.length || sh.getLastRow() < 2) return {};
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, ORD_HEAD.length).getValues();
  const out = {};
  asked.forEach(a => {
    const base = String(a.orderId || ''), tok = String(a.token || '');
    if (!base || !tok) return;
    const rows = vals.filter(r => (String(r[0]) === base || String(r[0]).indexOf(base + '-') === 0) && String(r[17]) === tok);
    if (rows.length) out[base] = rows.map(r => ({ id: String(r[0]), round: r[4] instanceof Date ? Utilities.formatDate(r[4], 'Asia/Bangkok', 'd MMM') : String(r[4]),
      total: Number(r[8]) || 0, pay: String(r[10]), status: String(r[11]), slip: !!String(r[15]) }));
  });
  return out;
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
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SH.ORD);
    const vals = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), ORD_HEAD.length).getValues();
    const base = String(req.orderId);
    const rows = [];
    vals.forEach((r, i) => { const id = String(r[0]); if (id === base || id.indexOf(base + '-') === 0) rows.push({ row: i + 2, vals: r }); });
    if (!rows.length) throw new Error('ไม่พบออเดอร์ ' + base);
    if (!req.token || rows.some(x => String(x.vals[17]) !== String(req.token))) throw new Error('ลิงก์แนบสลิปไม่ถูกต้อง');

    const ext = m[2] === 'jpeg' ? 'jpg' : m[2];
    const file = slipFolder_().createFile(
      Utilities.newBlob(bytes, m[1], base + '_' + rows[0].vals[2] + '_' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'MMdd-HHmm') + '.' + ext));
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    const id = file.getId();
    rows.forEach(x => {
      sh.getRange(x.row, 16, 1, 2).setFormulas([[
        '=IMAGE("https://drive.google.com/thumbnail?id=' + id + '&sz=w400")',
        '=HYPERLINK("' + file.getUrl() + '","เปิดสลิป")'
      ]]);
      if (x.vals[10] !== 'จ่ายแล้ว') sh.getRange(x.row, 11).setValue('ส่งสลิปแล้ว');
      sh.setRowHeight(x.row, 80);
    });
    return { orderId: req.orderId };
  } finally {
    lock.releaseLock();
  }
}
/** กด Run ฟังก์ชันนี้ 1 ครั้งใน Apps Script เพื่ออนุญาตให้เก็บสลิปใน Google Drive */
function authorize() {
  const f = slipFolder_();
  Logger.log('โฟลเดอร์สลิปพร้อมแล้ว: ' + f.getUrl());
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

/** Remark แบบเดิมที่เคยจด เช่น "รับ 5 + 1scoop No Shine Muscat" */
function remark_(round, l, addonById) {
  const p = [round || ''];
  if (l.addonCount.SCP) p.push('+' + l.addonCount.SCP + 'scoop');
  if (l.addonCount.HER) p.push('+hershey');
  Object.keys(l.addonCount).forEach(id => {
    if (id !== 'SCP' && id !== 'HER') p.push('+' + addonById[id].name);
  });
  if (l.excluded.length) p.push('ไม่เอา ' + l.excluded.join('/'));
  return p.join(' ').trim();
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
    ['line_group_url', '', 'ลิงก์เชิญกลุ่ม LINE (https://line.me/...) เว้นว่าง = ให้ลูกค้าเลือกแชทเอง']
  ]);
  ['Rounds', 'Menu', 'AddOns', 'Fruits'].forEach(n => {
    const sh = ss.getSheetByName(n);
    const col = n === 'Rounds' ? 2 : n === 'Menu' ? 7 : n === 'AddOns' ? 6 : 3;
    sh.getRange(2, col, sh.getMaxRows() - 1, 1).insertCheckboxes();
  });

  const ord = sheet_(ss, SH.ORD, ORD_HEAD, []);
  ord.getRange('K2:K').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(PAY_STATUS).build());
  ord.getRange('L2:L').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ORDER_STATUS).build());
  ord.getRange('B2:B').setNumberFormat('dd/MM HH:mm');
  ord.getRange('G2:I').setNumberFormat('#,##0');
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$L2="ยกเลิก"')
      .setBackground('#e0e0e0').setFontColor('#888888').setStrikethrough(true).setRanges([ord.getRange('A2:O')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$K2="จ่ายแล้ว"')
      .setBackground('#b7f5b0').setRanges([ord.getRange('A2:O')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($A2<>"",$K2="ยังไม่จ่าย")')
      .setBackground('#fff3cd').setRanges([ord.getRange('A2:O')]).build()
  ];
  ord.setConditionalFormatRules(rules);

  const item = sheet_(ss, SH.ITEM, ITEM_HEAD, []);
  item.getRange('B2:B').setNumberFormat('dd/MM HH:mm');
  item.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$Q2="ยกเลิก"')
      .setBackground('#e0e0e0').setFontColor('#888888').setRanges([item.getRange('A2:Q')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$P2="จ่ายแล้ว"')
      .setBackground('#b7f5b0').setRanges([item.getRange('A2:Q')]).build()
  ]);

  // หน้า Summary ใช้สูตรอัตโนมัติอยู่แล้ว สร้างใหม่เฉพาะตอนที่ยังไม่มีแท็บนี้
  if (!ss.getSheetByName(SH.SUM)) buildSummary();
  SpreadsheetApp.getActive().toast('ตั้งค่าเสร็จแล้ว ✅ ต่อไป Deploy เป็น Web app', 'Greek Shop', 6);
}

/** หน้าสรุปต่อรอบ — เลือกรอบที่ช่อง B1 */
function buildSummary() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SH.SUM);
  if (!sh) sh = ss.insertSheet(SH.SUM);
  sh.clear(); sh.clearConditionalFormatRules();

  const cfg = getConfig();
  const rounds = ['ทุกรอบ'].concat(cfg.rounds);
  sh.getRange('A1').setValue('เลือกรอบ ▶').setFontWeight('bold');
  sh.getRange('B1').setValue(cfg.rounds[cfg.rounds.length - 1] || 'ทุกรอบ')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(rounds).build())
    .setBackground('#fff59d').setFontWeight('bold');

  // เงื่อนไขกรองรอบ + ไม่นับออเดอร์ที่ยกเลิก
  const rI = 'IF($B$1="ทุกรอบ","*",$B$1)';
  const itemCond = 'OrderItems!$C:$C,' + rI + ',OrderItems!$Q:$Q,"<>ยกเลิก"';
  const ordCond = 'Orders!$E:$E,' + rI + ',Orders!$L:$L,"<>ยกเลิก"';

  let r = 3;
  const head = (t) => { sh.getRange(r, 1, 1, 3).setValues([[t, '', '']]).setFontWeight('bold').setBackground('#c9daf8'); r++; };

  head('🥣 จำนวนแต่ละเมนู (ถ้วย)');
  const allMenu = rows_(ss.getSheetByName(SH.MENU)).map(x => String(x[1]));
  allMenu.forEach(n => {
    sh.getRange(r, 1, 1, 3).setValues([[n,
      '=SUMIFS(OrderItems!$M:$M,OrderItems!$E:$E,A' + r + ',' + itemCond + ')',
      '=SUMIFS(OrderItems!$N:$N,OrderItems!$E:$E,A' + r + ',' + itemCond + ')']]);
    r++;
  });
  sh.getRange(r, 1, 1, 3).setValues([['รวม', '=SUM(B' + (r - allMenu.length) + ':B' + (r - 1) + ')',
    '=SUM(C' + (r - allMenu.length) + ':C' + (r - 1) + ')']]).setFontWeight('bold');
  r += 2;

  head('🍇 ผลไม้ที่ต้องเตรียม (ถ้วยที่ใส่)');
  const rnd = 'IF($B$1="ทุกรอบ",TRUE,OrderItems!$C$2:$C=$B$1)';
  rows_(ss.getSheetByName(SH.FRUIT)).forEach(f => {
    sh.getRange(r, 1, 1, 2).setValues([[String(f[0]),
      '=SUMPRODUCT(ISNUMBER(SEARCH(A' + r + ',OrderItems!$F$2:$F))*(' + rnd + ')*(OrderItems!$Q$2:$Q<>"ยกเลิก")*OrderItems!$M$2:$M)']]);
    r++;
  });
  r++;

  head('➕ Add-on');
  sh.getRange(r++, 1, 1, 2).setValues([['Scoop เพิ่ม (scoop)',
    '=SUMPRODUCT((' + rnd + ')*(OrderItems!$Q$2:$Q<>"ยกเลิก")*OrderItems!$H$2:$H*OrderItems!$M$2:$M)']]);
  sh.getRange(r++, 1, 1, 2).setValues([['Hershey (ถ้วย)',
    '=SUMPRODUCT((' + rnd + ')*(OrderItems!$Q$2:$Q<>"ยกเลิก")*(OrderItems!$I$2:$I="ใช่")*OrderItems!$M$2:$M)']]);
  r++;

  head('💰 ยอดเงิน (บาท)');
  const money = [
    ['จำนวนออเดอร์', '=COUNTIFS(' + ordCond + ',Orders!$A:$A,"G*")'],
    ['ยอดขายรวม (รวมค่าส่ง)', '=SUMIFS(Orders!$I:$I,' + ordCond + ')'],
    ['ค่าส่งรวม', '=SUMIFS(Orders!$G:$G,' + ordCond + ')'],
    ['✅ จ่ายแล้ว', '=SUMIFS(Orders!$I:$I,' + ordCond + ',Orders!$K:$K,"จ่ายแล้ว")'],
    ['⏳ ยังไม่จ่าย', '=SUMIFS(Orders!$I:$I,' + ordCond + ',Orders!$K:$K,"ยังไม่จ่าย")']
  ];
  sh.getRange(r, 1, money.length, 2).setValues(money);
  sh.getRange(r, 2, money.length, 1).setNumberFormat('#,##0');
  r += money.length + 1;

  head('⏳ รายชื่อที่ยังไม่จ่าย');
  sh.getRange(r, 1).setFormula('=IFERROR(QUERY(Orders!A2:L,"select C, I, A where K = \'ยังไม่จ่าย\' and L <> \'ยกเลิก\'' +
    '"&IF($B$1="ทุกรอบ",""," and E = \'"&$B$1&"\'"),0),"— ไม่มี —")');

  sh.setColumnWidth(1, 220); sh.setColumnWidth(2, 120); sh.setColumnWidth(3, 120);
  sh.getRange('C4:C').setNumberFormat('#,##0');
  sh.getRange('C3').setValue('ยอดเงิน');
  sh.getRange('B3').setValue('ถ้วย');
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
