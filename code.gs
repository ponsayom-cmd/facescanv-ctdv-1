/**
 * Smart School AI System - Backend Logic V4.4
 * ระบบบันทึกเวลาเรียนและระบบจัดการ Cloud (เพิ่มระบบตรวจสอบ Error)
 */

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Smart School AI System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** * ตรวจสอบและดึง Spreadsheet
 * หากบันทึกไม่ได้ ให้ตรวจสอบว่าได้กด "Authorize" สิทธิ์การเข้าถึง Spreadsheet หรือยัง
 */
function getTargetSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  const ssId = props.getProperty('SS_ID');
  
  if (!ssId) {
    console.warn("ไม่ได้ระบุ Spreadsheet ID ในระบบ Admin");
    return SpreadsheetApp.getActiveSpreadsheet();
  }

  try {
    return SpreadsheetApp.openById(ssId);
  } catch (e) {
    console.error("ไม่สามารถเปิด Spreadsheet ตาม ID ที่ระบุได้: " + e.message);
    // กรณี ID ผิดหรือไม่มีสิทธิ์ จะใช้ไฟล์ที่สคริปต์นี้ฝังตัวอยู่แทน
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

function saveSettings(id, url) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('SS_ID', id.trim());
    props.setProperty('SCRIPT_URL', url.trim());
    
    // ทดสอบเปิดไฟล์หลังจากบันทึก
    const ss = SpreadsheetApp.openById(id.trim());
    return { success: true, name: ss.getName() };
  } catch (e) {
    return { success: false, message: "ID ไม่ถูกต้อง หรือยังไม่ได้กดอนุญาตสิทธิ์: " + e.toString() };
  }
}

function getSettings() {
  const props = PropertiesService.getScriptProperties();
  return {
    ssId: props.getProperty('SS_ID') || "",
    scriptUrl: props.getProperty('SCRIPT_URL') || ""
  };
}

/** ฟังก์ชันบันทึกเวลาเรียน */
function checkInStudent(studentId, studentName) {
  try {
    const ss = getTargetSpreadsheet();
    let sheet = ss.getSheetByName('Attendance');
    if (!sheet) {
      sheet = ss.insertSheet('Attendance');
      sheet.appendRow(['StudentID', 'Name', 'Timestamp', 'Status']);
    }
    sheet.appendRow([studentId, studentName, new Date(), 'Checked-In']);
    SpreadsheetApp.flush(); // บังคับให้เขียนข้อมูลลง Sheet ทันที
    return { success: true };
  } catch (e) {
    console.error("Error in checkInStudent: " + e.toString());
    return { success: false, message: e.toString() };
  }
}

/** บันทึกข้อมูลนักเรียนใหม่ */
function registerStudent(studentData) {
  try {
    const ss = getTargetSpreadsheet();
    let sheet = ss.getSheetByName('Students');
    if (!sheet) {
      sheet = ss.insertSheet('Students');
      sheet.appendRow(['StudentID', 'FullName', 'Room', 'FaceFront', 'FaceLeft', 'FaceRight', 'Timestamp']);
    }
    sheet.appendRow([
      studentData.id, studentData.name, studentData.room,
      studentData.faces.front, studentData.faces.left, studentData.faces.right,
      new Date()
    ]);
    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    console.error("Error in registerStudent: " + e.toString());
    return { success: false, message: e.toString() };
  }
}

function addSubject(subjectData) {
  try {
    const ss = getTargetSpreadsheet();
    let sheet = ss.getSheetByName('Subjects');
    if (!sheet) {
      sheet = ss.insertSheet('Subjects');
      sheet.appendRow(['SubjectID', 'SubjectName', 'CreatedAt']);
    }
    sheet.appendRow([subjectData.id, subjectData.name, new Date()]);
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
  } catch (e) { 
    console.error("Error getStudents: " + e.toString());
    return []; 
  }
}

function getSubjects() {
  try {
    const ss = getTargetSpreadsheet();
    const sheet = ss.getSheetByName('Subjects');
    if (!sheet) return [];
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return [];
    values.shift();
    return values.map(row => ({ id: row[0], name: row[1] }));
  } catch (e) { return []; }
}

function clearDatabase(type) {
  try {
    const ss = getTargetSpreadsheet();
    const sheet = ss.getSheetByName(type);
    if (sheet && sheet.getLastRow() > 1) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
    }
    return { success: true };
  } catch (e) { return { success: false, message: e.toString() }; }
}
