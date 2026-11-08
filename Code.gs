/**
 * Code.gs — บันทึกผลการเรียน "สระลดรูปผจญภัย" ลง Google Sheet
 *
 * วิธีติดตั้ง (ทำครั้งเดียว):
 *   1) เปิด Google Sheet ของคุณ → เมนู Extensions (ส่วนขยาย) → Apps Script
 *   2) ลบโค้ดเดิมในไฟล์ Code.gs แล้ววางโค้ดนี้ทับทั้งหมด → กดบันทึก (Ctrl+S)
 *   3) กด Deploy → New deployment → ไอคอนเฟือง → เลือก Web app
 *        Description : SLR save
 *        Execute as  : Me (ฉัน)
 *        Who has access : Anyone (ทุกคน)
 *      → Deploy → Allow (อนุญาต)
 *   4) คัดลอก URL ที่ลงท้ายด้วย /exec
 *   5) วาง URL นั้นในหน้า "ตั้งค่า Sheet" ของเว็บ แล้วกด "ทดสอบการเชื่อมต่อ"
 *
 * ทดสอบเร็วๆ: เปิด URL /exec ในเบราว์เซอร์ ต้องเห็นข้อความ {"ok":true,...}
 * ถ้าแก้โค้ดภายหลัง ต้อง Deploy > Manage deployments > แก้ไข > Version: New version ทุกครั้ง
 */

var SHEET_ID = '1kdcfefZOt8Tl8n2P9xvlr6Qmgdrh6dmuPyYB2SLheow';
var DEFAULT_SHEET = 'ผลสอบ';

var HEADERS = [
  'วันที่เวลา', 'ชื่อนักเรียน', 'ชั้น', 'ห้อง', 'ประเภทการทดสอบ',
  'คะแนนเต็ม', 'คะแนนที่ได้', 'เปอร์เซ็นต์', 'ระดับ', 'อุปกรณ์'
];

/** เปิด URL เฉยๆ เพื่อตรวจว่าระบบพร้อม */
function doGet() {
  return json_({ ok: true, message: 'SLR save endpoint ready', sheetId: SHEET_ID, sheet: DEFAULT_SHEET });
}

/** รับข้อมูลจากเว็บ (POST) แล้วเพิ่มแถวในชีต */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // กันนักเรียนหลายคนส่งพร้อมกันแล้วแถวทับกัน
  } catch (err) {
    return json_({ ok: false, error: 'ระบบกำลังบันทึกแถวอื่นอยู่ กรุณาลองใหม่อีกครั้ง' });
  }

  try {
    var body = readBody_(e);
    var rows = Array.isArray(body.rows) ? body.rows : (body.record ? [body.record] : [body]);
    var sheetName = String(body.sheet || DEFAULT_SHEET).trim() || DEFAULT_SHEET;

    var ss = openSheet_();
    var sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    ensureHeaders_(sh);

    var values = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i] || {};
      if (!r.name && !r.type) continue; // ข้ามแถวว่าง
      values.push([
        r.time || '', r.name || '', r.grade || '', r.room || '', r.type || '',
        number_(r.total), number_(r.score), number_(r.pct), r.level || '', r.device || ''
      ]);
    }

    if (values.length) {
      sh.getRange(sh.getLastRow() + 1, 1, values.length, HEADERS.length).setValues(values);
    }

    return json_({ ok: true, added: values.length, sheet: sheetName });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function openSheet_() {
  try {
    return SpreadsheetApp.openById(SHEET_ID);
  } catch (err) {
    // กรณีสร้างสคริปต์จากในชีตนี้โดยตรง
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

function ensureHeaders_(sh) {
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
    return;
  }
  var first = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  var empty = true;
  for (var i = 0; i < first.length; i++) {
    if (String(first[i]).trim()) { empty = false; break; }
  }
  if (empty) sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
}

function readBody_(e) {
  var raw = '';
  if (e && e.postData && e.postData.contents) {
    raw = e.postData.contents;
  } else if (e && e.parameter && e.parameter.payload) {
    raw = e.parameter.payload;
  }
  if (!raw) return {};

  try {
    return JSON.parse(raw);
  } catch (err) {}

  // เผื่อกรณีส่งมาเป็น form-data แบบเข้ารหัส
  try {
    return JSON.parse(decodeURIComponent(raw.split('+').join(' ')));
  } catch (err) {
    return {};
  }
}

function number_(v) {
  var n = Number(v);
  return isFinite(n) ? n : 0;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
