/**
 * Smart School AI Backend - Google Apps Script
 * เชื่อมต่อฐานข้อมูล Google Sheets และจัดการ API Requests
 * พัฒนาให้สอดคล้องกับโครงสร้างหน้า Frontend (index.html) ปัจจุบัน
 */

const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE19nGHAiUm3pQoJqUkm86HulCR6N0IqgBvJJyLSJEDXoA'; // เปลี่ยนเป็น ID ของ Google Sheet ของคุณ

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('RMU CTD Smart School AI')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * ฟังก์ชันสำหรับ Frontend เรียกใช้: บันทึกข้อมูลนักศึกษาใหม่พร้อม Biometrics
 * สัมพันธ์กับฟังก์ชัน handleSaveStudent ใน index.html
 */
function registerStudent(studentData) {
  // ข้อมูลที่ส่งมา: { id: '...', name: '...', room: '...', biometrics: [...] }
  return saveDataToSheet('Students', {
    'Timestamp': new Date(),
    'StudentID': studentData.id,
    'FullName': studentData.name,
    'Room': studentData.room,
    'Biometrics': JSON.stringify(studentData.biometrics) // แปลง Array เป็น String เพื่อเก็บใน Sheet
  });
}

/**
 * ฟังก์ชันสำหรับ Frontend เรียกใช้: บันทึกการเช็คชื่อเข้าเรียน
 * สัมพันธ์กับฟังก์ชัน handleIdentify ใน index.html
 */
function recordAttendance(attendanceData) {
  // ข้อมูลที่ส่งมา: { id: '...', name: '...', subject: '...' }
  return saveDataToSheet('Attendance', {
    'Timestamp': new Date(),
    'StudentID': attendanceData.id,
    'FullName': attendanceData.name,
    'Subject': attendanceData.subject,
    'Status': 'Present'
  });
}

/**
 * ฟังก์ชันสำหรับดึงรายชื่อนักศึกษาทั้งหมด
 */
function getStudents() {
  return getDataFromSheet('Students');
}

/**
 * ฟังก์ชันสำหรับดึงรายวิชาทั้งหมด
 */
function getSubjects() {
  return getDataFromSheet('Subjects');
}

/**
 * ฟังก์ชันหลักในการบันทึกข้อมูลลง Google Sheets (Internal)
 */
function saveDataToSheet(sheetName, data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(sheetName);
    
    // ถ้ายังไม่มีแผ่นงาน ให้สร้างใหม่พร้อม Header
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      const headers = Object.keys(data);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = headers.map(header => data[header] || "");
    
    sheet.appendRow(newRow);
    
    return { status: 'success', message: 'บันทึกข้อมูลลง ' + sheetName + ' สำเร็จ' };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

/**
 * ดึงข้อมูลทั้งหมดจากแผ่นงาน (Internal)
 */
function getDataFromSheet(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return [];
    
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    
    return data.map(row => {
      const obj = {};
      headers.forEach((header, i) => {
        let val = row[i];
        // ถ้าเป็นคอลัมน์ Biometrics ให้แปลงกลับเป็น JSON Object
        if (header === 'Biometrics') {
          try { val = JSON.parse(val); } catch(e) { val = []; }
        }
        obj[header] = val;
      });
      return obj;
    });
  } catch (error) {
    return [];
  }
}

/**
 * ฟังก์ชันสำหรับล้างข้อมูลประวัติ (Maintenance)
 * สัมพันธ์กับส่วน Admin Panel ใน index.html
 */
function clearAttendanceHistory() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Attendance');
    if (sheet && sheet.getLastRow() > 1) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
    }
    return { status: 'success' };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}