// ==========================================
// การตั้งค่าระบบหลังบ้าน (Google Apps Script)
// สำหรับระบบบันทึกข้อมูลกำลังพล R-POST
// ==========================================

const SPREADSHEET_ID = "1DShEQEn-UN23OLxiuZbnL0ETX9udxNxvCgLy3ZMhAaU"; // <- นำ ID จาก URL ของ Google Sheet มาใส่
const FOLDER_ID = "1iPmZQUHM6S2jYgKfd0ZKFTaP_8OCK2w1"; // <- นำ ID ของ Google Drive Folder สำหรับเก็บไฟล์ PDF และ Excel มาใส่

// ชื่อชีตต่างๆ ใน Google Sheet
const SHEET_NAME = "Sheet1"; // สำหรับเก็บข้อมูลรายบุคคล
const USERS_SHEET_NAME = "Users"; // สำหรับเก็บข้อมูลผู้ใช้งาน (ล็อกอิน)
const BULK_SHEET_NAME = "หน่วย"; // สำหรับเก็บรายชื่อหน่วย

// ฟังก์ชันหลักที่รับ Request จาก HTML Frontend
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const data = payload.data;
    
    let result = {};
    
    switch (action) {
      case 'getUnitList':
        result = getUnitList();
        break;
      case 'loginUser':
        result = loginUser(data);
        break;
      case 'registerUser':
        result = registerUser(data);
        break;
      case 'getUserHistory':
        result = getUserHistory(data);
        break;
      case 'processFormSubmission':
        result = processFormSubmission(data);
        break;
      case 'processBulkSubmission':
        result = processBulkSubmission(data);
        break;
      case 'getReportData':
        result = getReportData(data);
        break;
      case 'getAdminDashboardData':
        result = getAdminDashboardData(); // <- เพิ่มฟังก์ชันสำหรับแอดมิน
        break;
      default:
        result = { success: false, message: "Action not found: " + action };
    }
    
    // ส่งข้อมูลกลับไปให้ Frontend ในรูปแบบ JSON
    return ContentService.createTextOutput(JSON.stringify(result))
                         .setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.toString() }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// ส่วนที่ 1: การจัดการผู้ใช้ และ รายชื่อหน่วย
// ดึงรายชื่อหน่วย
function getUnitList() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(BULK_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return ["หน่วยทดสอบ 1", "หน่วยทดสอบ 2"];
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  return data.map(r => r[0]).filter(String);
}

// ระบบเข้าสู่ระบบ
function loginUser(data) {
  const email = data.email.trim().toLowerCase();
  const password = data.password;
  
  // กำหนดสิทธิ์ Admin หากอีเมลมีคำว่า admin
  const isAdmin = email.includes('admin');
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(USERS_SHEET_NAME);
  
  if (!sheet) return { success: false, message: "ไม่พบฐานข้อมูลผู้ใช้งาน" };
  
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const dbEmail = rows[i][0] ? rows[i][0].toString().trim().toLowerCase() : "";
    if (dbEmail === email && rows[i][1].toString() === password.toString()) {
      return { success: true, message: "เข้าสู่ระบบสำเร็จ", isAdmin: isAdmin, unit: rows[i][2] || "-" };
    }
  }
  
  return { success: false, message: "อีเมลหรือรหัสผ่านไม่ถูกต้อง" };
}

// ระบบสมัครสมาชิก
function registerUser(data) {
  const email = data.email.trim();
  const password = data.password;
  const unit = data.unit;
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(USERS_SHEET_NAME);
  
  // ถ้ายังไม่มี Sheet Users ให้สร้างใหม่
  if (!sheet) {
    sheet = ss.insertSheet(USERS_SHEET_NAME);
    sheet.appendRow(["Email", "Password", "Unit", "Timestamp"]);
  }
  
  // ตรวจสอบว่ามีอีเมลนี้แล้วหรือยัง
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === email) {
      return { success: false, message: "อีเมลนี้ถูกใช้งานแล้ว" };
    }
  }
  
  sheet.appendRow([email, password, unit, new Date()]);
  return { success: true, message: "สมัครสมาชิกสำเร็จ" };
}

// ==========================================
// ส่วนที่ 2: การบันทึกข้อมูลและอัปโหลดไฟล์
// ==========================================

// บันทึกข้อมูลรายบุคคล และอัปโหลดไฟล์ PDF
function processFormSubmission(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  // เตรียมชีต
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "Timestamp", "Email", "ID_Card", "Rank", "FirstName", "LastName", 
      "Unit", "Affiliation", "RegularPosition", "FieldPosition", 
      "Score", "Percentage", "Status", "IssueDate", "ExpireDate", "PDF_Link"
    ]);
  }
  
  let fileUrl = "-";
  
  // อัปโหลดไฟล์ PDF ถ้ามี
  if (data.hasCertificate === 'มีใบประกาศ' && data.pdfFileBase64) {
    try {
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const parts = data.pdfFileBase64.split(',');
      const decoded = Utilities.base64Decode(parts[1]); // แปลง Base64 กลับเป็นไฟล์
      const blob = Utilities.newBlob(decoded, 'application/pdf', `${data.idCard}_${data.firstName}.pdf`);
      const file = folder.createFile(blob);
      fileUrl = file.getUrl();
    } catch (e) {
      fileUrl = "Error: " + e.message;
    }
  }
  
  // บันทึกข้อมูลลงชีต
  sheet.appendRow([
    new Date(),
    data.email,
    data.idCard,
    data.rank,
    data.firstName,
    data.lastName,
    data.unit,
    data.affiliation,
    data.regularPosition,
    data.fieldPosition,
    data.score || "-",
    data.percentage || "-",
    data.status || "ไม่มีใบประกาศ",
    data.issueDate || "-",
    data.expireDate || "-",
    fileUrl
  ]);
  
  return { success: true, message: "บันทึกข้อมูลเรียบร้อยแล้ว" };
}

// บันทึกบัญชีรายชื่อ (Excel)
function processBulkSubmission(data) {
  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const parts = data.excelFileBase64.split(',');
    const decoded = Utilities.base64Decode(parts[1]);
    const blob = Utilities.newBlob(decoded, data.excelFileMimeType, `${data.unit}_${new Date().getTime()}_${data.excelFileName}`);
    folder.createFile(blob);
    
    return { success: true, message: "อัปโหลดไฟล์ Excel เรียบร้อยแล้ว" };
  } catch (e) {
    return { success: false, message: "ไม่สามารถอัปโหลดไฟล์ได้: " + e.message };
  }
}

// ==========================================
// ส่วนที่ 3: การดึงข้อมูลรายงานและ Dashboard
// ==========================================

// ดึงประวัติการบันทึกข้อมูลเฉพาะของผู้ใช้ (User History)
function getUserHistory(email) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const history = [];
  
  // ลูปเช็คข้อมูลแถวต่อแถว ข้ามแถวที่ 1 (Header)
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === email) { // คอลัมน์ Email
      history.push(data[i]);
    }
  }
  
  return history;
}

// ดึงข้อมูลผู้ที่ ผ่าน/ไม่ผ่าน 
function getReportData(status) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const report = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][12] === status || status === 'ทั้งหมด') { // คอลัมน์ Status
      report.push(data[i]);
    }
  }
  
  return report;
}

// ดึงข้อมูลทั้งหมดสำหรับแสดงบน Admin Dashboard
function getAdminDashboardData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const dashboardData = [];
  
  // ดึงข้อมูลทั้งหมดข้ามบรรทัด Header (Index 0)
  for (let i = 1; i < data.length; i++) {
    dashboardData.push(data[i]);
  }
  
  return dashboardData;
}
