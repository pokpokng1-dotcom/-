/**
 * ===================================================
 *  Content Library - Google Apps Script Backend
 *  (เวอร์ชัน fetch-based สำหรับเรียกจากภายนอก เช่น GitHub Pages)
 * ===================================================
 * โครงสร้างคอลัมน์ใน Google Sheet (แถวที่ 1 เป็น Header):
 *   A: ID
 *   B: Category
 *   C: Title
 *   D: Content
 *   E: Source
 *   F: Timestamp
 *
 * วิธีใช้:
 * 1. เปิด Google Sheet ใหม่ (จะมี Header ให้อัตโนมัติ ไม่ต้องพิมพ์เอง)
 * 2. Extensions > Apps Script แล้ววางไฟล์นี้ทับ Code.gs เดิม
 *    (ไม่ต้องมีไฟล์ index.html ในโปรเจกต์ Apps Script อีกต่อไป
 *     เพราะหน้าเว็บจะไปอยู่บน GitHub Pages แทน)
 * 3. กด Deploy > New deployment > Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. คัดลอก URL ที่ลงท้ายด้วย /exec เอาไปใส่ในตัวแปร API_URL
 *    ของไฟล์ index.html (ฝั่ง GitHub)
 * 5. ทุกครั้งที่แก้โค้ด Code.gs ต้องกด Deploy > Manage deployments
 *    > แก้ไข (ไอคอนดินสอ) > Version: New version > Deploy ใหม่
 *    ไม่งั้น URL เดิมจะยังใช้โค้ดเวอร์ชันเก่าอยู่ (จุดที่คนพลาดบ่อยที่สุด)
 */

const HEADERS = ['ID', 'Category', 'Title', 'Content', 'Source', 'Timestamp'];

/**
 * รับ GET request
 * - ถ้ามี ?action=getData  -> คืนข้อมูลทั้งหมดเป็น JSON
 * - ถ้าไม่มี param อะไรเลย -> คืนข้อความบอกสถานะเฉยๆ (เผื่อเปิด URL เช็คเล่น)
 */
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === 'getData') {
    return jsonResponse_({ success: true, data: getData() });
  }

  return jsonResponse_({
    success: true,
    message: 'Content Library API is running. ใช้ ?action=getData เพื่อดึงข้อมูล หรือส่ง POST เพื่อเพิ่มข้อมูล'
  });
}

/**
 * รับ POST request สำหรับเพิ่มข้อมูลใหม่
 * ฝั่ง client ต้องส่ง body เป็น JSON string ผ่าน Content-Type: text/plain
 * (สำคัญ: ใช้ text/plain ไม่ใช่ application/json เพื่อเลี่ยง CORS preflight
 *  ซึ่ง Apps Script Web App ไม่รองรับ OPTIONS request)
 *
 * รูปแบบ body ที่รองรับ:
 *   { "action": "addData", "title": "...", "category": "...",
 *     "content": "...", "source": "..." }
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    if (payload.action === 'deleteData') {
      return jsonResponse_(deleteData(payload.id));
    }

    // ค่าเริ่มต้น: เพิ่มข้อมูลใหม่
    return jsonResponse_(addData(payload));
  } catch (err) {
    return jsonResponse_({ success: false, message: 'เกิดข้อผิดพลาด: ' + err.message });
  }
}

/**
 * helper: ห่อผลลัพธ์เป็น JSON response
 * หมายเหตุ: Apps Script Web App (deploy แบบ Anyone) จะแนบ
 * Access-Control-Allow-Origin: * ให้อัตโนมัติสำหรับ response แบบนี้อยู่แล้ว
 * ไม่ต้องตั้งค่า header เพิ่มเอง (และ setHeader ก็ไม่รองรับใน ContentService ด้วย)
 */
function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Utility: ดึง Sheet object พร้อมสร้าง Header อัตโนมัติถ้ายังไม่มี
 * ใช้ "sheet แรกที่มีอยู่จริง" ของสเปรดชีตนี้เสมอ (ไม่อิงชื่อ tab)
 */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

/**
 * ดึงข้อมูลทั้งหมดใน Sheet ออกมาเป็น Array ของ Object
 */
function getData() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const range = sheet.getRange(2, 1, lastRow - 1, HEADERS.length);
  const values = range.getValues();

  return values
    .filter(row => row[0] !== '')
    .map(row => ({
      id: row[0],
      category: row[1],
      title: row[2],
      content: row[3],
      source: row[4],
      timestamp: row[5] instanceof Date ? row[5].toISOString() : row[5]
    }))
    .reverse();
}

/**
 * เพิ่มข้อมูลใหม่ลงใน Sheet
 * formObject = { category, title, content, source }
 */
function addData(formObject) {
  try {
    const sheet = getSheet_();
    const newId = Utilities.getUuid();
    const timestamp = new Date();

    sheet.appendRow([
      newId,
      formObject.category || '',
      formObject.title || '',
      formObject.content || '',
      formObject.source || '',
      timestamp
    ]);
    SpreadsheetApp.flush();
    Logger.log('เขียนข้อมูลลง sheet: ' + sheet.getName());

    return {
      success: true,
      message: 'บันทึกข้อมูลเรียบร้อยแล้ว!',
      data: getData()
    };
  } catch (err) {
    return {
      success: false,
      message: 'เกิดข้อผิดพลาด: ' + err.message
    };
  }
}

/**
 * ลบข้อมูลตาม ID
 */
function deleteData(id) {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: 'ไม่พบข้อมูล' };

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) {
      sheet.deleteRow(i + 2);
      return { success: true, data: getData() };
    }
  }
  return { success: false, message: 'ไม่พบข้อมูล' };
}

/**
 * ฟังก์ชันตรวจสอบ (Debug) — รันเองใน Apps Script editor เพื่อเช็คว่า
 * ผูกกับ Sheet ไหน และมีข้อมูลกี่แถว ดูผลได้ที่ Executions log
 */
function checkSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet_();
  Logger.log('ชื่อสเปรดชีต: ' + ss.getName());
  Logger.log('URL สเปรดชีต: ' + ss.getUrl());
  Logger.log('ชื่อแท็บที่ใช้บันทึกข้อมูลจริง: ' + sheet.getName());
  Logger.log('จำนวนแถวข้อมูล (ไม่รวม header): ' + Math.max(0, sheet.getLastRow() - 1));
}
