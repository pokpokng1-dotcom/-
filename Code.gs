/**
 * ===================================================
 *  Content Library - Google Apps Script Backend
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
 * 1. เปิด Google Sheet ใหม่ ตั้งชื่อ Header ตามด้านบน (แถวที่ 1)
 * 2. เปิด Extensions > Apps Script แล้ววางไฟล์นี้เป็น Code.gs
 * 3. สร้างไฟล์ HTML ชื่อ index.html แล้ววาง index.html ที่ให้ไปด้วย
 * 4. กด Deploy > New deployment > Web app
 *    - Execute as: Me
 *    - Who has access: Anyone (หรือปรับตามต้องการ)
 * 5. คัดลอก URL ของ Web App ไปเปิดใช้งาน
 */

const SHEET_NAME = 'Sheet1'; // แก้ชื่อ Sheet ให้ตรงกับของจริงถ้าจำเป็น
const HEADERS = ['ID', 'Category', 'Title', 'Content', 'Source', 'Timestamp'];

/**
 * เสิร์ฟหน้าเว็บหลัก (index.html) เมื่อมีการเปิด Web App URL
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('คลังความรู้ส่วนตัว (Content Library)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Utility: ดึง Sheet object พร้อมสร้าง Header อัตโนมัติถ้ายังไม่มี
 */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

/**
 * ดึงข้อมูลทั้งหมดใน Sheet ออกมาเป็น Array ของ Object
 * เรียกใช้จาก client ผ่าน google.script.run.getData()
 */
function getData() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const range = sheet.getRange(2, 1, lastRow - 1, HEADERS.length);
  const values = range.getValues();

  return values
    .filter(row => row[0] !== '') // ข้ามแถวว่าง
    .map(row => ({
      id: row[0],
      category: row[1],
      title: row[2],
      content: row[3],
      source: row[4],
      timestamp: row[5] instanceof Date ? row[5].toISOString() : row[5]
    }))
    .reverse(); // ให้รายการล่าสุดอยู่บนสุด
}

/**
 * เพิ่มข้อมูลใหม่ลงใน Sheet
 * เรียกใช้จาก client ผ่าน google.script.run.addData(formObject)
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

    return {
      success: true,
      message: 'บันทึกข้อมูลเรียบร้อยแล้ว!',
      data: getData() // ส่งข้อมูลล่าสุดกลับไปเพื่ออัปเดตตารางทันที
    };
  } catch (err) {
    return {
      success: false,
      message: 'เกิดข้อผิดพลาด: ' + err.message
    };
  }
}

/**
 * ลบข้อมูลตาม ID (ฟังก์ชันเสริม เผื่อใช้งาน)
 */
function deleteData(id) {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
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
 * (ทางเลือก) เปิดให้เรียกเป็น JSON API ตรงๆ ผ่าน URL
 * เช่น .../exec?action=json
 * มีไว้เผื่อกรณีต้องการเรียกจากภายนอกด้วย fetch()
 */
function doGetJson_(e) {
  const output = ContentService.createTextOutput(JSON.stringify(getData()));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
