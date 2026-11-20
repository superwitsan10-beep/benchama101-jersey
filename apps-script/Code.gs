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
    const itemsSheet = ss.getSheetByName(ITEMS_SHEET);
    const startRow = itemsSheet.getLastRow() + 1;
    // คอลัมน์ "สถานะ" ใน Items ดึงจาก Orders อัตโนมัติ (แก้สถานะที่ Orders ที่เดียว)
    const itemRows = clean.items.map((it, i) => [orderNo, i + 1, it.model, it.size, safeCell_(it.name), it.number === '' ? '' : "'" + it.number, it.price, safeCell_(clean.name), clean.delivery, statusFormula_(startRow + i)]);
    itemsSheet.getRange(startRow, 1, itemRows.length, ITEM_HEADERS.length).setValues(itemRows);
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

/** สูตรดึงสถานะจากชีต Orders ตามเลขออเดอร์ในคอลัมน์ A ของแถวนั้น */
function statusFormula_(row) {
  return `=IFERROR(VLOOKUP($A${row},${ORDERS_SHEET}!$A:$C,${STATUS_COL},FALSE),"")`;
}

/** รันครั้งเดียว: เปลี่ยนคอลัมน์สถานะของแถวเดิมใน Items ให้ดึงจาก Orders อัตโนมัติ */
function fixItemsStatus() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ITEMS_SHEET);
  const last = sh.getLastRow();
  if (last < 2) return;
  const col = ITEM_HEADERS.indexOf('สถานะ') + 1;
  const formulas = [];
  for (let r = 2; r <= last; r++) formulas.push([statusFormula_(r)]);
  sh.getRange(2, col, formulas.length, 1).setFormulas(formulas);
}

/* ===================== SUMMARY ===================== */
const SUMMARY_SHEET = 'Summary';
const SCREEN_SHEET = 'รายชื่อสกรีน';
const EXCLUDE_STATUS = ['ยกเลิก', 'สลิปไม่ถูกต้อง'];   // ไม่นับในยอด
const PAID_STATUS = ['ชำระแล้ว', 'จัดส่งแล้ว', 'รับสินค้าแล้ว'];
const ALL_STATUS = ['รอตรวจสอบสลิป', 'ชำระแล้ว', 'สลิปไม่ถูกต้อง', 'จัดส่งแล้ว', 'รับสินค้าแล้ว', 'ยกเลิก'];

/**
 * รันครั้งเดียว (รันซ้ำได้ จะสร้างใหม่ทับของเดิม): สร้างแท็บ Summary และ รายชื่อสกรีน
 * ทุกช่องเป็นสูตร ตัวเลขอัปเดตเองเมื่อมีออเดอร์ใหม่หรือแก้สถานะในแท็บ Orders
 */
function setupSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  buildSummary_(ss);
  buildScreenList_(ss);
}

function buildSummary_(ss) {
  const sh = ss.getSheetByName(SUMMARY_SHEET) || ss.insertSheet(SUMMARY_SHEET, 0);
  sh.clear(); sh.clearFormats();
  const RED = '#E5251F', DARK = '#18181B', SOFT = '#FDECEC';
  const ex = EXCLUDE_STATUS.map(s => `,Items!$J$2:$J,"<>${s}"`).join('');
  const exO = EXCLUDE_STATUS.map(s => `,Orders!$C$2:$C,"<>${s}"`).join('');
  const models = Object.keys(MODEL_NAMES).map(k => MODEL_NAMES[k]);
  const lastCol = SIZES.length + 2; // แบบ + ไซซ์ + รวม

  sh.getRange(1, 1).setValue('สรุปยอดพรีออเดอร์ BENCHAMA101').setFontSize(16).setFontWeight('bold');
  sh.getRange(2, 1).setValue('นับทุกออเดอร์ ยกเว้น "' + EXCLUDE_STATUS.join('" และ "') + '" · อัปเดตอัตโนมัติจากแท็บ Orders / Items').setFontColor('#71717A');

  // 1) แบบ × ไซซ์
  let r = 4;
  sh.getRange(r, 1).setValue('ยอดเสื้อแยกตามแบบ × ไซซ์ (สำหรับสั่งผลิต)').setFontWeight('bold').setFontSize(12);
  r++;
  const head = ['แบบ'].concat(SIZES, ['รวม']);
  sh.getRange(r, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground(RED).setFontColor('#FFFFFF').setHorizontalAlignment('center');
  const headRow = r;
  models.forEach((m, i) => {
    const row = headRow + 1 + i;
    const vals = [m];
    SIZES.forEach((s, j) => {
      const col = colLetter_(j + 2);
      vals.push(`=COUNTIFS(Items!$C$2:$C,$A${row},Items!$D$2:$D,${col}$${headRow}${ex})`);
    });
    vals.push(`=SUM(B${row}:${colLetter_(SIZES.length + 1)}${row})`);
    sh.getRange(row, 1, 1, vals.length).setValues([vals]);
  });
  const totalRow = headRow + models.length + 1;
  const tot = ['รวมทุกแบบ'];
  for (let c = 2; c <= lastCol; c++) { const L = colLetter_(c); tot.push(`=SUM(${L}${headRow + 1}:${L}${totalRow - 1})`); }
  sh.getRange(totalRow, 1, 1, tot.length).setValues([tot]).setFontWeight('bold').setBackground(SOFT);
  sh.getRange(headRow + 1, 2, models.length + 1, lastCol - 1).setHorizontalAlignment('center');
  sh.getRange(headRow + 1, lastCol, models.length + 1, 1).setFontWeight('bold');
  sh.getRange(headRow, 1, models.length + 2, lastCol).setBorder(true, true, true, true, true, true, '#D4D4D8', SpreadsheetApp.BorderStyle.SOLID);

  // 2) ภาพรวม
  r = totalRow + 3;
  sh.getRange(r, 1).setValue('ภาพรวม').setFontWeight('bold').setFontSize(12);
  const paid = PAID_STATUS.map(s => `SUMIFS(Orders!$P$2:$P,Orders!$C$2:$C,"${s}")`).join('+');
  const overview = [
    ['จำนวนออเดอร์', `=COUNTIFS(Orders!$A$2:$A,"?*"${exO})`, 'ออเดอร์'],
    ['จำนวนเสื้อ', `=SUMIFS(Orders!$M$2:$M,Orders!$A$2:$A,"?*"${exO})`, 'ตัว'],
    ['ค่าเสื้อ', `=SUMIFS(Orders!$N$2:$N,Orders!$A$2:$A,"?*"${exO})`, 'บาท'],
    ['ค่าจัดส่ง', `=SUMIFS(Orders!$O$2:$O,Orders!$A$2:$A,"?*"${exO})`, 'บาท'],
    ['ยอดรวมทั้งหมด', `=SUMIFS(Orders!$P$2:$P,Orders!$A$2:$A,"?*"${exO})`, 'บาท'],
    ['ยอดที่ชำระแล้ว', `=${paid}`, 'บาท'],
    ['รับที่งาน', `=COUNTIFS(Orders!$F$2:$F,"รับที่งาน"${exO})`, 'ออเดอร์'],
    ['จัดส่ง', `=COUNTIFS(Orders!$F$2:$F,"จัดส่ง"${exO})`, 'ออเดอร์']
  ];
  sh.getRange(r + 1, 1, overview.length, 3).setValues(overview);
  sh.getRange(r + 1, 2, overview.length, 1).setNumberFormat('#,##0').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange(r + 5, 1, 1, 3).setBackground(SOFT);
  sh.getRange(r + 1, 1, overview.length, 3).setBorder(true, true, true, true, true, true, '#D4D4D8', SpreadsheetApp.BorderStyle.SOLID);

  // 3) สถานะ
  const sr = r;
  sh.getRange(sr, 5).setValue('จำนวนออเดอร์ตามสถานะ (นับทุกออเดอร์)').setFontWeight('bold').setFontSize(12);
  const st = ALL_STATUS.map(s => [s, `=COUNTIF(Orders!$C$2:$C,"${s}")`]);
  st.push(['รวม', `=SUM(F${sr + 1}:F${sr + ALL_STATUS.length})`]);
  sh.getRange(sr + 1, 5, st.length, 2).setValues(st);
  sh.getRange(sr + 1, 6, st.length, 1).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(sr + st.length, 5, 1, 2).setBackground(SOFT).setFontWeight('bold');
  sh.getRange(sr + 1, 5, st.length, 2).setBorder(true, true, true, true, true, true, '#D4D4D8', SpreadsheetApp.BorderStyle.SOLID);

  sh.setColumnWidth(1, 170);
  for (let c = 2; c <= lastCol; c++) sh.setColumnWidth(c, 56);
  sh.setColumnWidth(3, 70); sh.setColumnWidth(5, 130);
  sh.setFrozenRows(0);
  sh.getRange(headRow, 1, 1, lastCol).setFontColor('#FFFFFF');
  sh.setTabColor(RED);
}

function buildScreenList_(ss) {
  const sh = ss.getSheetByName(SCREEN_SHEET) || ss.insertSheet(SCREEN_SHEET, 1);
  sh.clear(); sh.clearFormats();
  const head = ['แบบ', 'ไซซ์', 'ชื่อบนเสื้อ', 'หมายเลข', 'เลขออเดอร์', 'ชื่อผู้สั่ง', 'สถานะ', '(ลำดับไซซ์)'];
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#E5251F').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  const sizeList = '{' + SIZES.map(s => `"${s}"`).join(',') + '}';
  const keep = EXCLUDE_STATUS.map(s => `,Items!J2:J<>"${s}"`).join('');
  const f = `=IFERROR(SORT(FILTER({Items!C2:C,Items!D2:D,Items!E2:E,Items!F2:F,Items!A2:A,Items!H2:H,Items!J2:J,MATCH(Items!D2:D,${sizeList},0)},Items!A2:A<>""${keep}),1,TRUE,8,TRUE,4,TRUE),"ยังไม่มีออเดอร์")`;
  sh.getRange(2, 1).setFormula(f);
  sh.hideColumns(8);
  sh.setColumnWidth(1, 150); sh.setColumnWidth(3, 170); sh.setColumnWidth(5, 110); sh.setColumnWidth(6, 160);
}

function colLetter_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/** กันข้อความที่ขึ้นต้นด้วย = + - @ ไม่ให้ชีตตีความเป็นสูตร */
function safeCell_(v) {
  const s = String(v == null ? '' : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
