/**
 * ระบบจองบ้านพัก (Google Apps Script Web App)
 *
 * ชีทที่ใช้:
 *  - "รายชื่อพนักงาน"           : รายชื่อพนักงาน คอลัมน์ D (เริ่มแถว 3)
 *  - "ฐานข้อมูลรายการห้องพัก"    : ประเภทห้อง คอลัมน์ A, เลขห้อง คอลัมน์ B (เริ่มแถว 4)
 *  - "DB_CheckIn"              : บันทึกการจอง (เริ่มแถว 3)
 */

var SHEET_STAFF = 'รายชื่อพนักงาน';
var SHEET_ROOMS = 'ฐานข้อมูลรายการห้องพัก';
var SHEET_CHECKIN = 'DB_CheckIn';

var STAFF_COL = 4;        // D
var STAFF_FIRST_ROW = 3;
var ROOMS_FIRST_ROW = 4;  // A = ประเภท, B = เลขห้อง
var CHECKIN_FIRST_ROW = 3;

// คอลัมน์ใน DB_CheckIn (1 = A)
var COL = {
  timestamp: 1,    // A ประทับเวลา
  staff: 2,        // B ชื่อผู้รับผิดชอบ
  customerType: 4, // D ประเภทลูกค้า
  roomType: 5,     // E ประเภทห้อง
  roomNo: 6,       // F เลขห้อง
  checkIn: 7,      // G วันที่ Check in
  checkOut: 8,     // H วันที่ Check out
  nights: 9,       // I จำนวนคืน
  status: 10       // J สถานะห้อง
};

var CUSTOMER_TYPES = [
  'เจ้าของบ้านเข้าพัก',
  'ลูกค้าทั่วไปเข้าพัก',
  'ลูกค้าเข้าพัก ใช้บัตรของขวัญ',
  'อินฟลูเอนเซอร์เข้าพัก'
];
var STATUS_CLOSED = 'งดให้บริการ';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ระบบจองบ้านพัก')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSheet_(name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('ไม่พบชีท "' + name + '"');
  return sheet;
}

/** ข้อมูลเริ่มต้นสำหรับหน้าบันทึกการจอง */
function getBookingFormData() {
  var staffSheet = getSheet_(SHEET_STAFF);
  var staff = [];
  var staffRows = staffSheet.getLastRow() - STAFF_FIRST_ROW + 1;
  if (staffRows > 0) {
    staffSheet.getRange(STAFF_FIRST_ROW, STAFF_COL, staffRows, 1).getValues()
      .forEach(function (r) {
        var name = String(r[0]).trim();
        if (name) staff.push(name);
      });
  }

  // กลุ่มห้องอ่านจากชีทแบบไดนามิก: เพิ่มประเภท/ห้องใหม่ในชีทแล้วจะแสดงเองอัตโนมัติ
  var roomSheet = getSheet_(SHEET_ROOMS);
  var groups = {};
  var order = [];
  var roomRows = roomSheet.getLastRow() - ROOMS_FIRST_ROW + 1;
  if (roomRows > 0) {
    roomSheet.getRange(ROOMS_FIRST_ROW, 1, roomRows, 2).getDisplayValues()
      .forEach(function (r) {
        var type = String(r[0]).trim();
        var room = String(r[1]).trim();
        if (!type || !room) return;
        if (!groups[type]) {
          groups[type] = [];
          order.push(type);
        }
        if (groups[type].indexOf(room) === -1) groups[type].push(room);
      });
  }

  return {
    staff: staff,
    customerTypes: CUSTOMER_TYPES,
    roomGroups: order.map(function (t) { return { type: t, rooms: groups[t] }; })
  };
}

/**
 * บันทึกการจอง
 * @param {Object} form {staff, closed, customerType, roomType, roomNo, checkIn, checkOut}
 *   checkIn / checkOut เป็นข้อความรูปแบบ yyyy-mm-dd
 */
function saveBooking(form) {
  if (!form || !form.staff) throw new Error('กรุณาเลือกชื่อผู้รับผิดชอบ');
  if (!form.roomType || !form.roomNo) throw new Error('กรุณาเลือกประเภทห้องและเลขห้อง');
  if (!form.closed && CUSTOMER_TYPES.indexOf(form.customerType) === -1) {
    throw new Error('กรุณาเลือกประเภทลูกค้า');
  }

  var checkIn = parseDate_(form.checkIn);
  var checkOut = parseDate_(form.checkOut);
  if (!checkIn || !checkOut) throw new Error('กรุณาเลือกวันที่ Check in และ Check out');
  var nights = Math.round((checkOut - checkIn) / 86400000);
  if (nights < 1) throw new Error('วันที่ Check out ต้องอยู่หลังวันที่ Check in');

  var status = form.closed ? STATUS_CLOSED : form.customerType;

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(SHEET_CHECKIN);
    var row = nextEmptyRow_(sheet);
    var lastCol = Math.max.apply(null, Object.keys(COL).map(function (k) { return COL[k]; }));
    var values = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

    values[COL.timestamp - 1] = new Date();
    values[COL.staff - 1] = form.staff;
    values[COL.customerType - 1] = form.closed ? '' : form.customerType;
    values[COL.roomType - 1] = form.roomType;
    values[COL.roomNo - 1] = form.roomNo;
    values[COL.checkIn - 1] = checkIn;
    values[COL.checkOut - 1] = checkOut;
    values[COL.nights - 1] = nights;
    values[COL.status - 1] = status;

    sheet.getRange(row, 1, 1, lastCol).setValues([values]);
    sheet.getRange(row, COL.timestamp).setNumberFormat('dd/MM/yyyy HH:mm:ss');
    sheet.getRange(row, COL.checkIn).setNumberFormat('dd/MM/yyyy');
    sheet.getRange(row, COL.checkOut).setNumberFormat('dd/MM/yyyy');
    SpreadsheetApp.flush();

    return { row: row, nights: nights, status: status };
  } finally {
    lock.releaseLock();
  }
}

/** แถวว่างถัดไปโดยดูจากคอลัมน์ประทับเวลา (A) ตั้งแต่แถว 3 */
function nextEmptyRow_(sheet) {
  var last = sheet.getLastRow();
  if (last < CHECKIN_FIRST_ROW) return CHECKIN_FIRST_ROW;
  var colA = sheet.getRange(CHECKIN_FIRST_ROW, COL.timestamp, last - CHECKIN_FIRST_ROW + 1, 1).getValues();
  for (var i = colA.length - 1; i >= 0; i--) {
    if (colA[i][0] !== '') return CHECKIN_FIRST_ROW + i + 1;
  }
  return CHECKIN_FIRST_ROW;
}

function parseDate_(s) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}
