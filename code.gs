/**
 * Smart School AI System - Backend V5.0 (Auto-ID Detection)
 * ระบบจะพยายามค้นหาไฟล์ Spreadsheet อัตโนมัติเพื่อความง่ายในการตั้งค่า
 */

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Smart School AI System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** * ฟังก์ชันดึง Spreadsheet แบบชาญฉลาด
 * 1. ตรวจสอบใน Properties ก่อน
 * 2. ถ้าไม่มี ให้ใช้ไฟล์ที่ Script นี้วางอยู่ (Active)
 */
function getTargetSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('SS_ID');
  
  if (savedId && savedId.length > 5) {
    try {
      return SpreadsheetApp.openById(savedId);
    } catch (e) {
      console.warn("ใช้ ID ที่บันทึกไว้ไม่ได้: " + e.message);
    }
  }
  
  // ถ้าไม่ได้ระบุ ID หรือ ID ผิด ให้ใช้ไฟล์ปัจจุบันที่สคริปต์รันอยู่
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSettings() {
  const props = PropertiesService.getScriptProperties();
  const ss = getTargetSpreadsheet();
  return {
    ssId: props.getProperty('SS_ID') || ss.getId(), // ส่ง ID ปัจจุบันกลับไปโชว์
    ssName: ss.getName(), // ส่งชื่อไฟล์กลับไปโชว์ในหน้า Admin เพื่อความมั่นใจ
    scriptUrl: props.getProperty('SCRIPT_URL') || ""
  };
}

function saveSettings(id, url) {
  try {
    const props = PropertiesService.getScriptProperties();
    if (id) props.setProperty('SS_ID', id.trim());
    if (url) props.setProperty('SCRIPT_URL', url.trim());
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/** บันทึกเวลานักเรียน (Attendance) */
function checkInStudent(studentId, studentName) {
  try {
    const ss = getTargetSpreadsheet();
    let sheet = ss.getSheetByName('Attendance');
    if (!sheet) {
      sheet = ss.insertSheet('Attendance');
      sheet.appendRow(['รหัสนักศึกษา', 'ชื่อ-นามสกุล', 'เวลาบันทึก', 'สถานะ']);
    }
    sheet.appendRow([studentId, studentName, new Date(), 'เข้าเรียน']);
    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/** บันทึกรายชื่อนักเรียน (Students) */
function registerStudent(studentData) {
  try {
    const ss = getTargetSpreadsheet();
    let sheet = ss.getSheetByName('Students');
    if (!sheet) {
      sheet = ss.insertSheet('Students');
      sheet.appendRow(['StudentID', 'FullName', 'Room', 'FaceData', 'Timestamp']);
    }
    sheet.appendRow([
      studentData.id, 
      studentData.name, 
      studentData.room, 
      studentData.faceData, 
      new Date()
    ]);
    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getStudents() {
  try {
    const ss = getTargetSpreadsheet();
    const sheet = ss.getSheetByName('Students');
    if (!sheet) return [];
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return [];
    values.shift();
    return values.map(row => ({ StudentID: row[0], FullName: row[1], Room: row[2] }));
  } catch (e) { return []; }
}
