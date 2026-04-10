/**
 * Smart School AI System - Backend Logic V4.0
 * ระบบนี้จะเชื่อมต่อกับ Google Sheets ที่สคริปต์นี้ฝังตัวอยู่โดยอัตโนมัติ
 */

// ฟังก์ชันหลักในการเปิดใช้งาน Web App
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Smart School AI System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** * ฟังก์ชันสำหรับบันทึกนักเรียนพร้อมรูปภาพใบหน้า 3 มุม 
 * @param {Object} studentData ข้อมูลนักเรียนที่ส่งมาจากหน้าบ้าน
 */
function registerStudent(studentData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Students');
    
    // หากไม่มี Sheet ให้สร้างขึ้นใหม่พร้อมหัวตาราง
    if (!sheet) {
      sheet = ss.insertSheet('Students');
      sheet.appendRow(['StudentID', 'FullName', 'Room', 'FaceFront', 'FaceLeft', 'FaceRight', 'Timestamp']);
    }
    
    sheet.appendRow([
      studentData.id,
      studentData.name,
      studentData.room,
      studentData.faces.front, // เก็บเป็น Base64 String
      studentData.faces.left,
      studentData.faces.right,
      new Date()
    ]);
    
    return { success: true, message: "บันทึกข้อมูลนักเรียนสำเร็จ" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/** ฟังก์ชันเพิ่มวิชาเรียน */
function addSubject(subjectData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Subjects');
    
    if (!sheet) {
      sheet = ss.insertSheet('Subjects');
      sheet.appendRow(['SubjectID', 'SubjectName', 'CreatedAt']);
    }
    
    sheet.appendRow([subjectData.id, subjectData.name, new Date()]);
    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/** ดึงข้อมูลนักเรียนทั้งหมดจาก Sheet */
function getStudents() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Students');
    if (!sheet) return [];
    
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return []; // มีแต่หัวตาราง
    
    values.shift(); // ลบหัวตารางออก
    return values.map(row => ({
      StudentID: row[0],
      FullName: row[1],
      Room: row[2]
    }));
  } catch (error) {
    return [];
  }
}

/** ดึงข้อมูลรายวิชาทั้งหมด */
function getSubjects() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Subjects');
    if (!sheet) return [];
    
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return [];
    
    values.shift();
    return values.map(row => ({
      id: row[0],
      name: row[1]
    }));
  } catch (error) {
    return [];
  }
}

/** ระบบ Admin: ล้างข้อมูลใน Sheet (เฉพาะข้อมูล ไม่ลบหัวตาราง) */
function clearDatabase(type) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(type);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
      }
    }
    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
