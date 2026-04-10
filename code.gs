/**
 * Smart School AI System - Backend Logic V4.3
 * ระบบบันทึกเวลาเรียนและระบบจัดการ Cloud
 */

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Smart School AI System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getTargetSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  const ssId = props.getProperty('SS_ID');
  try {
    if (ssId) return SpreadsheetApp.openById(ssId);
  } catch (e) { console.error("SS ID Error: " + e.message); }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function saveSettings(id, url) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('SS_ID', id);
    props.setProperty('SCRIPT_URL', url);
    const ss = SpreadsheetApp.openById(id);
    return { success: true, name: ss.getName() };
  } catch (e) { return { success: false, message: e.toString() }; }
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
    return { success: true };
  } catch (e) { return { success: false, message: e.toString() }; }
}

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
    return { success: true };
  } catch (e) { return { success: false, message: e.toString() }; }
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
    return { success: true };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function getStudents() {
  try {
    const ss = getTargetSpreadsheet();
    const sheet = ss.getSheetByName('Students');
    if (!sheet) return [];
    const values = sheet.getDataRange().getValues();
    values.shift();
    return values.map(row => ({ StudentID: row[0], FullName: row[1], Room: row[2] }));
  } catch (e) { return []; }
}

function getSubjects() {
  try {
    const ss = getTargetSpreadsheet();
    const sheet = ss.getSheetByName('Subjects');
    if (!sheet) return [];
    const values = sheet.getDataRange().getValues();
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
  } catch (e) { return { success: false }; }
}
