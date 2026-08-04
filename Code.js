/**
 * ฟังก์ชันหลักในการให้บริการเว็บแอปพลิเคชัน (Web App)
 * ทำหน้าที่เชื่อมโยง URL และส่งผ่านพารามิเตอร์โหมดระบบ
 */
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  var mode = (e && e.parameter && e.parameter.mode) ? e.parameter.mode : 'edit';
  template.mode = mode;
  
  return template.evaluate()
      .setTitle('ระบบบริหารจัดการโรงเรียนกุดจับประชาสรรค์ - วิชาชุมนุม')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * ฟังก์ชันดึงข้อมูลเริ่มต้นและจัดเตรียมโครงสร้างชีตบน Google Sheets (Current Sheet)
 */
function getDatabaseData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); // ใช้งาน Current Sheet ทันที
    
    // 1. ตรวจสอบหรือสร้างชีต "Students" (เก็บข้อมูลผู้สมัครเรียน)
    var studentSheet = ss.getSheetByName("Students");
    if (!studentSheet) {
      studentSheet = ss.insertSheet("Students");
      studentSheet.appendRow(["ID", "StudentID", "Prefix", "Name", "Level", "Username", "Password", "Role"]);
      
      // บัญชีครูผู้สอนเริ่มต้น (Username: admin, Password: admin1234)
      studentSheet.appendRow([
        "admin-id", 
        "99999", 
        "นาย", 
        "สมชาย รักเรียน", 
        "กลุ่มสาระคอมพิวเตอร์", 
        "admin", 
        "admin1234", 
        "teacher"
      ]);
      studentSheet.appendRow(["1", "10001", "เด็กชาย", "กฤษณะ รักเรียน", "ม.1", "student1", "1234", "student"]);
      studentSheet.appendRow(["2", "10002", "เด็กหญิง", "พิมลดา รักษ์ดี", "ม.2", "student2", "1234", "student"]);
    }
    
    // 2. ตรวจสอบหรือสร้างชีต "Settings" (เก็บค่าตั้งค่าของชุมนุม)
    var settingsSheet = ss.getSheetByName("Settings");
    if (!settingsSheet) {
      settingsSheet = ss.insertSheet("Settings");
      settingsSheet.appendRow(["Key", "Value"]);
      settingsSheet.appendRow(["clubName", "ชุมนุม GAS (Google Apps Script)"]);
      settingsSheet.appendRow(["teacherName", "ครูสมชาย รักเรียน"]);
      settingsSheet.appendRow(["clubLogo", ""]);
      settingsSheet.appendRow(["maxCapacity", "40"]);
    }

    // 3. ตรวจสอบหรือสร้างชีต "Attendance" (โครงสร้างใหม่: 1 วัน = 1 แถว เก็บแบบ JSON Array)
    var attendanceSheet = ss.getSheetByName("Attendance");
    if (!attendanceSheet) {
      attendanceSheet = ss.insertSheet("Attendance");
      attendanceSheet.appendRow(["Date", "AttendanceData"]);
      
      var defaultRecords = {
        "10001": "present",
        "10002": "present"
      };
      attendanceSheet.appendRow(["2026-06-01", JSON.stringify(defaultRecords)]);
    }

    // อ่านข้อมูล Students
    var studentData = studentSheet.getDataRange().getValues();
    var students = [];
    for (var i = 1; i < studentData.length; i++) {
      students.push({
        id: studentData[i][0].toString(),
        studentId: studentData[i][1].toString(),
        prefix: studentData[i][2].toString(),
        name: studentData[i][3].toString(),
        level: studentData[i][4].toString(),
        username: studentData[i][5].toString(),
        role: studentData[i][7].toString()
      });
    }

    // อ่านข้อมูล Settings
    var settingsData = settingsSheet.getDataRange().getValues();
    var settings = {};
    for (var i = 1; i < settingsData.length; i++) {
      settings[settingsData[i][0]] = settingsData[i][1].toString();
    }

    // อ่านประวัติการเข้าเรียน ( AttendanceData )
    var attendanceData = attendanceSheet.getDataRange().getValues();
    var records = {};
    for (var i = 1; i < attendanceData.length; i++) {
      var date = attendanceData[i][0].toString();
      if (attendanceData[i][0] instanceof Date) {
        var d = attendanceData[i][0];
        var year = d.getFullYear();
        var month = ("0" + (d.getMonth() + 1)).slice(-2);
        var day = ("0" + d.getDate()).slice(-2);
        date = year + "-" + month + "-" + day;
      }
      
      var rowJsonString = attendanceData[i][1].toString();
      try {
        records[date] = JSON.parse(rowJsonString);
      } catch(err) {
        records[date] = {};
      }
    }

    return {
      students: students,
      settings: settings,
      records: records,
      success: true
    };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

/**
 * ฟังก์ชันล็อกอินตรวจสอบสิทธิ์ผู้ใช้งานบน Google Sheet
 */
function loginUser(username, password) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Students");
    if (!sheet) return { success: false, message: "ไม่พบฐานข้อมูลรายชื่อ" };
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var dbUser = data[i][5].toString().trim();
      var dbPass = data[i][6].toString().trim();
      if (dbUser === username.trim() && dbPass === password.trim()) {
        return {
          success: true,
          user: {
            id: data[i][0].toString(),
            studentId: data[i][1].toString(),
            prefix: data[i][2].toString(),
            name: data[i][3].toString(),
            level: data[i][4].toString(),
            username: dbUser,
            role: data[i][7].toString()
          }
        };
      }
    }
    return { success: false, message: "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง" };
  } catch(e) {
    return { success: false, message: "ระบบล็อกอินขัดข้อง: " + e.message };
  }
}

/**
 * สมัครเรียนเข้าชุมนุมเองของนักเรียน
 */
function registerUser(studentId, prefix, name, level, username, password, role) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Students");
    if (!sheet) return { success: false, message: "ระบบฐานข้อมูลขัดข้อง" };

    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][5].toString().trim() === username.trim()) {
        return { success: false, message: "ชื่อผู้ใช้นี้มีคนใช้งานในระบบแล้ว" };
      }
      if (role === 'student' && data[i][1].toString().trim() === studentId.trim()) {
        return { success: false, message: "รหัสนักเรียนนี้ได้ทำการสมัครเรียนชุมนุมแล้ว" };
      }
    }

    // Check Max Capacity for students
    if (role === 'student') {
      var settingsSheet = ss.getSheetByName("Settings");
      var maxCapacity = 40; // Default
      if (settingsSheet) {
        var sData = settingsSheet.getDataRange().getValues();
        for (var j = 1; j < sData.length; j++) {
          if (sData[j][0].toString() === "maxCapacity") {
            maxCapacity = parseInt(sData[j][1]) || 40;
            break;
          }
        }
      }
      var studentCount = 0;
      for (var i = 1; i < data.length; i++) {
        if (data[i][7].toString() === 'student') {
          studentCount++;
        }
      }
      if (studentCount >= maxCapacity) {
        return { success: false, message: "ขออภัย ชุมนุมนี้เต็มจำนวน " + maxCapacity + " คนแล้ว ไม่สามารถรับสมัครเพิ่มได้" };
      }
    }

    var newId = Date.now().toString();
    sheet.appendRow([
      newId,
      studentId,
      prefix,
      name,
      level,
      username,
      password,
      role
    ]);

    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * ฟังก์ชันช่วยตรวจสอบสิทธิ์ครูผู้สอนฝั่ง Server-side
 */
function verifyTeacherCredentials(username, password) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Students");
    if (!sheet) return false;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var dbUser = data[i][5].toString().trim();
      var dbPass = data[i][6].toString().trim();
      var dbRole = data[i][7].toString().trim();
      if (dbUser === username.trim() && dbPass === password.trim() && dbRole === "teacher") {
        return true;
      }
    }
  } catch(e) {}
  return false;
}

/**
 * บันทึกการเช็คชื่อแบบแมนนวล (1 คาบ/วัน = 1 แถว ในชีตจบงาน)
 */
function saveAttendanceOnSheet(date, recordsJsonString, authUser, authPass) {
  try {
    if (!verifyTeacherCredentials(authUser, authPass)) return false;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Attendance");
    if (!sheet) return false;
    
    var data = sheet.getDataRange().getValues();
    var existingRowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      var rDate = data[i][0].toString();
      if (data[i][0] instanceof Date) {
        var d = data[i][0];
        var year = d.getFullYear();
        var month = ("0" + (d.getMonth() + 1)).slice(-2);
        var day = ("0" + d.getDate()).slice(-2);
        rDate = year + "-" + month + "-" + day;
      }
      
      if (rDate === date) {
        existingRowIndex = i + 1;
        break;
      }
    }
    
    if (existingRowIndex !== -1) {
      sheet.getRange(existingRowIndex, 2).setValue(recordsJsonString);
    } else {
      sheet.appendRow([date, recordsJsonString]);
    }
    return true;
  } catch(e) {
    return false;
  }
}

/**
 * 🌟 CRUD: ลบประวัติการเช็คชื่อรายวัน
 */
function deleteAttendanceDay(date, authUser, authPass) {
  try {
    if (!verifyTeacherCredentials(authUser, authPass)) return false;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Attendance");
    if (!sheet) return false;
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var rDate = data[i][0].toString();
      if (data[i][0] instanceof Date) {
        var d = data[i][0];
        var year = d.getFullYear();
        var month = ("0" + (d.getMonth() + 1)).slice(-2);
        var day = ("0" + d.getDate()).slice(-2);
        rDate = year + "-" + month + "-" + day;
      }
      if (rDate === date) {
        sheet.deleteRow(i + 1);
        return true;
      }
    }
    return false;
  } catch(e) {
    return false;
  }
}

/**
 * บันทึกประวัตินักเรียนทั้งหมด (ครู CRUD)
 */
function updateStudentsOnSheet(studentsJsonString, authUser, authPass) {
  try {
    if (!verifyTeacherCredentials(authUser, authPass)) return false;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Students");
    if (!sheet) return false;
    
    var existingData = sheet.getDataRange().getValues();
    var passwordMap = {};
    for (var i = 1; i < existingData.length; i++) {
      passwordMap[existingData[i][0].toString()] = existingData[i][6].toString();
    }
    
    sheet.clear();
    sheet.appendRow(["ID", "StudentID", "Prefix", "Name", "Level", "Username", "Password", "Role"]);
    
    var students = JSON.parse(studentsJsonString);
    for (var i = 0; i < students.length; i++) {
      var pId = students[i].id;
      var pass = passwordMap[pId] ? passwordMap[pId] : "1234";
      
      sheet.appendRow([
        students[i].id, 
        students[i].studentId, 
        students[i].prefix || "นาย", 
        students[i].name, 
        students[i].level, 
        students[i].username || students[i].studentId, 
        pass, 
        students[i].role || "student"
      ]);
    }
    return true;
  } catch(e) {
    return false;
  }
}

/**
 * บันทึกสัญญลักษณ์และตั้งค่ากิจกรรมวิชาชุมนุม
 */
function saveSettingsOnSheet(clubName, teacherName, clubLogo, maxCapacity, studyDay, studyTime, authUser, authPass) {
  try {
    if (!verifyTeacherCredentials(authUser, authPass)) return false;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Settings");
    if (!sheet) return false;
    
    sheet.clear();
    sheet.appendRow(["Key", "Value"]);
    sheet.appendRow(["clubName", clubName]);
    sheet.appendRow(["teacherName", teacherName]);
    sheet.appendRow(["clubLogo", clubLogo]);
    sheet.appendRow(["maxCapacity", maxCapacity.toString()]);
    sheet.appendRow(["studyDay", studyDay || "วันพุธ"]);
    sheet.appendRow(["studyTime", studyTime || "13:00 - 15:00"]);
    return true;
  } catch(e) {
    return false;
  }
}

/**
 * นำเข้ารายชื่อนักเรียนจำนวนมากจาก CSV
 */
function importStudentsOnSheet(studentsListJson, authUser, authPass) {
  try {
    if (!verifyTeacherCredentials(authUser, authPass)) {
      return { success: false, message: "ไม่มีสิทธิ์ในการเข้าถึงข้อมูล" };
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Students");
    if (!sheet) return { success: false, message: "ไม่พบฐานข้อมูล" };
    
    var newStudents = JSON.parse(studentsListJson);
    var existingData = sheet.getDataRange().getValues();
    var existingUsernames = {};
    var existingStudentIds = {};
    
    for (var i = 1; i < existingData.length; i++) {
      existingStudentIds[existingData[i][1].toString().trim()] = true;
      existingUsernames[existingData[i][5].toString().trim()] = true;
    }
    
    var count = 0;
    for (var k = 0; k < newStudents.length; k++) {
      var s = newStudents[k];
      var sId = s.studentId.toString().trim();
      var username = s.username.toString().trim();
      
      // Skip duplicate IDs or Usernames
      if (existingStudentIds[sId] || existingUsernames[username]) continue;
      
      sheet.appendRow([
        Date.now().toString() + "-" + k,
        sId,
        s.prefix || "นาย",
        s.name,
        s.level,
        username,
        s.password || "1234",
        "student"
      ]);
      
      existingStudentIds[sId] = true;
      existingUsernames[username] = true;
      count++;
    }
    return { success: true, count: count };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * บันทึกรหัสผ่านเช็คชื่อประจำวันสำหรับสแกน QR Code
 */
function saveCheckinSettings(date, passcode, authUser, authPass) {
  if (!verifyTeacherCredentials(authUser, authPass)) return false;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Settings");
  if (!sheet) return false;
  
  var data = sheet.getDataRange().getValues();
  var hasPasscode = false;
  var hasDate = false;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === "checkinPasscode") {
      sheet.getRange(i + 1, 2).setValue(passcode);
      hasPasscode = true;
    }
    if (data[i][0].toString() === "checkinDate") {
      sheet.getRange(i + 1, 2).setValue(date);
      hasDate = true;
    }
  }
  if (!hasPasscode) sheet.appendRow(["checkinPasscode", passcode]);
  if (!hasDate) sheet.appendRow(["checkinDate", date]);
  return true;
}

/**
 * นักเรียนยืนยันตนเองเพื่อเช็คชื่อผ่าน QR Code
 */
function checkInStudent(studentId, date, passcode, username, password) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // ยืนยันตัวตนนักเรียน
    var sSheet = ss.getSheetByName("Students");
    var sData = sSheet.getDataRange().getValues();
    var isValidStudent = false;
    for (var i = 1; i < sData.length; i++) {
      if (sData[i][5].toString().trim() === username.trim() && 
          sData[i][6].toString().trim() === password.trim() && 
          sData[i][1].toString().trim() === studentId.toString().trim()) {
        isValidStudent = true;
        break;
      }
    }
    if (!isValidStudent) return { success: false, message: "การยืนยันตัวตนล้มเหลว รหัสผ่านอาจจะสลับหรือชื่อผู้ใช้ไม่ถูกต้อง" };

    // ดึงรหัสผ่านลงทะเบียนและวันที่ล่าสุด
    var setSheet = ss.getSheetByName("Settings");
    var setData = setSheet.getDataRange().getValues();
    var dbPasscode = "";
    var dbDate = "";
    for (var j = 1; j < setData.length; j++) {
      if (setData[j][0].toString() === "checkinPasscode") dbPasscode = setData[j][1].toString().trim();
      if (setData[j][0].toString() === "checkinDate") dbDate = setData[j][1].toString().trim();
    }

    if (dbPasscode !== passcode.trim() || dbDate !== date.trim()) {
      return { success: false, message: "รหัสเช็คชื่อหมดอายุหรือไม่ถูกต้อง กรุณาสแกน QR Code ใหม่ล่าสุด" };
    }

    // ทำการเช็คชื่อลงชีต Attendance
    var attSheet = ss.getSheetByName("Attendance");
    var attData = attSheet.getDataRange().getValues();
    var rowIndex = -1;
    var currentRecords = {};

    for (var k = 1; k < attData.length; k++) {
      var rDate = attData[k][0].toString();
      if (attData[k][0] instanceof Date) {
        var d = attData[k][0];
        rDate = d.getFullYear() + "-" + ("0" + (d.getMonth() + 1)).slice(-2) + "-" + ("0" + d.getDate()).slice(-2);
      }
      if (rDate === date) {
        rowIndex = k + 1;
        try {
          currentRecords = JSON.parse(attData[k][1].toString());
        } catch(e) {}
        break;
      }
    }

    currentRecords[studentId] = "present";
    var recordString = JSON.stringify(currentRecords);

    if (rowIndex !== -1) {
      attSheet.getRange(rowIndex, 2).setValue(recordString);
    } else {
      attSheet.appendRow([date, recordString]);
    }
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}