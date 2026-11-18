/**
 * BENCHAMA101 Jersey Pre-order — Google Apps Script backend
 *
 * ทำอะไร:
 *  - รับออเดอร์จากหน้าเว็บ (POST) → ตรวจข้อมูล → คำนวณยอดใหม่ฝั่งเซิร์ฟเวอร์
 *  - รันเลขออเดอร์ yyyymm-N (เริ่มที่ 0 ทุกเดือน) แบบไม่ซ้ำ แม้สั่งพร้อมกัน
 *  - เก็บรูปสลิปลง Google Drive และบันทึกออเดอร์ลงชีต
 *  - ให้หน้าเว็บเช็กสถานะออเดอร์ได้ (GET ?order=...&phone=...)
 *
 * ตั้งค่า: แก้ SLIP_FOLDER_ID ด้านล่าง แล้วรันฟังก์ชัน setup() หนึ่งครั้ง
 */

const SLIP_FOLDER_ID = 'ใส่_ID_โฟลเดอร์_Google_Drive_ที่นี่';
const TZ = 'Asia/Bangkok';

const PRICES = { FANCLUB: 399, PLAYER: 329, GOALKEEPER: 329 };
const MODEL_NAMES = { FANCLUB: 'FANCLUB JERSEY', PLAYER: 'PLAYER JERSEY', GOALKEEPER: 'GOALKEEPER JERSEY' };
const SIZES = ['SS','S','M','L','XL','2XL','3XL','4XL','5XL','6XL','7XL','8XL'];
const SHIP_FEE = 50;

const ORDERS_SHEET = 'Orders';
const ITEMS_SHEET = 'Items';
const ORDER_HEADERS = ['เลขออเดอร์','วันเวลา','สถานะ','ชื่อ-นามสกุล','เบอร์โทร','การรับสินค้า','ที่อยู่','อำเภอ/เขต','จังหวัด','รหัสไปรษณีย์','หมายเหตุ','รายการเสื้อ','จำนวนตัว','ค่าเสื้อ','ค่าจัดส่ง','ยอดรวม','ลิงก์สลิป'];
const ITEM_HEADERS = ['เลขออเดอร์','ลำดับ','แบบ','ไซซ์','ชื่อบนเสื้อ','หมายเลข','ราคา','ชื่อผู้สั่ง','การรับสินค้า','สถานะ'];
const STATUS_COL = 3; // คอลัมน์ "สถานะ" ในชีต Orders

/** รันครั้งเดียวตอนตั้งค่า: สร้างชีตและหัวตาราง */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [[ORDERS_SHEET, ORDER_HEADERS], [ITEMS_SHEET, ITEM_HEADERS]].forEach(([name, headers]) => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#E5251F').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  });
  // ช่องสถานะให้เลือกจากรายการ
  const orders = ss.getSheetByName(ORDERS_SHEET);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['รอตรวจสอบสลิป','ชำระแล้ว','สลิปไม่ถูกต้อง','จัดส่งแล้ว','รับสินค้าแล้ว','ยกเลิก'], true).build();
  orders.getRange(2, STATUS_COL, orders.getMaxRows() - 1, 1).setDataValidation(rule);
  DriveApp.getFolderById(SLIP_FOLDER_ID); // เช็กว่า ID โฟลเดอร์ถูกต้อง และขอสิทธิ์ Drive
}

/** รับออเดอร์ใหม่ */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const data = JSON.parse(e.postData.contents);
    const clean = validate_(data); // โยน error ถ้าข้อมูลไม่ครบ

    lock.waitLock(30000); // กันเลขออเดอร์ซ้ำเวลาสั่งพร้อมกัน
    const orderNo = nextOrderNo_();
    const slipUrl = saveSlip_(clean.slip, orderNo);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = new Date();
    const itemsText = clean.items.map((it, i) =>
      `${i + 1}) ${it.model} ${it.size}${it.name ? ' ' + it.name : ''}${it.number !== '' ? ' #' + it.number : ''}`).join('\n');

    ss.getSheetByName(ORDERS_SHEET).appendRow([
      orderNo, now, 'รอตรวจสอบสลิป', safeCell_(clean.name), "'" + clean.phone, clean.delivery,
      safeCell_(clean.address.address), safeCell_(clean.address.district), clean.address.province, clean.address.zip ? "'" + clean.address.zip : '',
      safeCell_(clean.note), itemsText, clean.items.length, clean.subtotal, clean.shipping, clean.total, slipUrl
    ]);
    const itemRows = clean.items.map((it, i) => [orderNo, i + 1, it.model, it.size, safeCell_(it.name), it.number === '' ? '' : "'" + it.number, it.price, clean.name, clean.delivery, 'รอตรวจสอบสลิป']);
    const itemsSheet = ss.getSheetByName(ITEMS_SHEET);
    itemsSheet.getRange(itemsSheet.getLastRow() + 1, 1, itemRows.length, ITEM_HEADERS.length).setValues(itemRows);
    SpreadsheetApp.flush();

    return json_({ ok: true, orderNo, total: clean.total, status: 'รอตรวจสอบสลิป' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/** เช็กสถานะ: GET ?phone=0812345678 → ทุกออเดอร์ของเบอร์นี้ (ใหม่สุดก่อน) */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const phone = String(p.phone || '').replace(/\D/g, '');
  if (!phone) return json_({ ok: true, service: 'jersey-preorder' });
  if (!/^0\d{8,9}$/.test(phone)) return json_({ ok: false, error: 'เบอร์โทรไม่ถูกต้อง' });
  const rows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET).getDataRange().getDisplayValues();
  const orders = rows.slice(1)
    .filter(r => r[0] && r[4].replace(/\D/g, '') === phone)
    .reverse()
    .slice(0, 20)
    .map(r => ({ orderNo: r[0], status: r[2], delivery: r[5], items: r[11], total: r[15] }));
  if (!orders.length) return json_({ ok: false, error: 'ไม่พบออเดอร์ของเบอร์นี้ กรุณาตรวจสอบเบอร์อีกครั้ง' });
  return json_({ ok: true, orders });
}

/* ---------- helpers ---------- */

function validate_(d) {
  const name = String(d.customer && d.customer.name || '').trim();
  const phone = String(d.customer && d.customer.phone || '').replace(/\D/g, '');
  if (name.length < 2) throw new Error('กรุณากรอกชื่อ-นามสกุล');
  if (!/^0\d{8,9}$/.test(phone)) throw new Error('เบอร์โทรไม่ถูกต้อง');
  if (!Array.isArray(d.items) || d.items.length === 0 || d.items.length > 50) throw new Error('รายการเสื้อไม่ถูกต้อง');

  // คำนวณราคาใหม่ฝั่งเซิร์ฟเวอร์ ไม่เชื่อยอดที่หน้าเว็บส่งมา
  const items = d.items.map(it => {
    const key = String(it.modelKey || '').toUpperCase();
    if (!PRICES[key]) throw new Error('แบบเสื้อไม่ถูกต้อง');
    if (SIZES.indexOf(it.size) === -1) throw new Error('ไซซ์ไม่ถูกต้อง');
    const number = String(it.number == null ? '' : it.number).trim();
    if (number !== '' && !/^\d{1,3}$/.test(number)) throw new Error('หมายเลขต้องอยู่ระหว่าง 0–999');
    return { model: MODEL_NAMES[key], size: it.size, name: Array.from(String(it.name || '').trim()).slice(0, 20).join(''), number, price: PRICES[key] };
  });

  const ship = d.delivery === 'ship';
  const a = d.address || {};
  const address = ship
    ? { address: String(a.address || '').trim(), district: String(a.district || '').trim(), province: String(a.province || '').trim(), zip: String(a.zip || '').trim() }
    : { address: '', district: '', province: '', zip: '' };
  if (ship && (!address.address || !address.district || !address.province || !/^\d{5}$/.test(address.zip))) throw new Error('ที่อยู่จัดส่งไม่ครบ');

  if (!d.slip || !/^data:image\/(png|jpe?g|webp|heic|heif);base64,/.test(d.slip)) throw new Error('กรุณาแนบสลิป');

  const subtotal = items.reduce((s, it) => s + it.price, 0);
  const shipping = ship ? SHIP_FEE : 0;
  return { name, phone, items, delivery: ship ? 'จัดส่ง' : 'รับที่งาน', address, note: String(d.note || '').slice(0, 500), slip: d.slip, subtotal, shipping, total: subtotal + shipping };
}

function nextOrderNo_() {
  const ym = Utilities.formatDate(new Date(), TZ, 'yyyyMM');
  const props = PropertiesService.getScriptProperties();
  const key = 'seq_' + ym;
  const cur = props.getProperty(key);
  const n = cur === null ? 0 : Number(cur) + 1;
  props.setProperty(key, String(n));
  return ym + '-' + n;
}

function saveSlip_(dataUrl, orderNo) {
  const m = dataUrl.match(/^data:(image\/[\w.+-]+);base64,(.+)$/);
  const ext = m[1].split('/')[1].replace('jpeg', 'jpg');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], `slip_${orderNo}.${ext}`);
  return DriveApp.getFolderById(SLIP_FOLDER_ID).createFile(blob).getUrl();
}

/** กันข้อความที่ขึ้นต้นด้วย = + - @ ไม่ให้ชีตตีความเป็นสูตร */
function safeCell_(v) {
  const s = String(v == null ? '' : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
