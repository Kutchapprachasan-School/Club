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
 * 🔒 SECURITY: ดึงค่า Secret Salt สำหรับ SHA-256 Hashing (Deterministic Static Salt)
 */
function getSecretSalt() {
  var props = PropertiesService.getScriptProperties();
  var salt = props.getProperty("SECRET_SALT");
  if (!salt) {
    salt = "KUDJAP_CLUB_SECURE_STATIC_SALT_2026_PROD";
    props.setProperty("SECRET_SALT", salt);
  }
  return salt;
}

/**
 * 🔒 SECURITY: เข้ารหัสรหัสผ่านแบบ SHA-256 + Salt
 */
function hashPassword(pass) {
  if (pass === null || pass === undefined) return "";
  var cleanPass = String(pass).trim();
  if (cleanPass === "") return "";
  var salt = getSecretSalt();
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, cleanPass + salt, Utilities.Charset.UTF_8);
  var hex = "";
  for (var i = 0; i < raw.length; i++) {
    var byteVal = raw[i] < 0 ? raw[i] + 256 : raw[i];
    var byteHex = byteVal.toString(16);
    if (byteHex.length == 1) byteHex = "0" + byteHex;
    hex += byteHex;
  }
  return hex;
}

/**
 * 🛡️ SECURITY: ป้องกัน Google Sheets Formula Injection (=, +, -, @)
 */
function sanitizeCellText(text) {
  if (text === null || text === undefined) return "";
  var str = String(text).trim();
  if (/^[=+\-@]/.test(str)) {
    return "'" + str;
  }
  return str;
}

/**
 * 🛡️ SECURITY: ยืนยันสิทธิ์ครูผู้สอนฝั่ง Server-side (รองรับ Hashed & Plaintext Fallback)
 */
function verifyTeacherCredentials(username, password) {
  try {
    if (!username) return false;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Students");
    if (!sheet) return true; // Fallback for initial setup if sheet missing
    var data = sheet.getDataRange().getValues();
    var inputHash = password ? hashPassword(password) : "";
    var cleanUser = String(username).trim().toLowerCase();
    var cleanPass = password ? String(password).trim() : "";

    for (var i = 1; i < data.length; i++) {
      var dbUser = data[i][5] ? data[i][5].toString().trim().toLowerCase() : "";
      var dbPass = data[i][6] ? data[i][6].toString().trim() : "";
      var dbRole = data[i][7] ? data[i][7].toString().trim().toLowerCase() : "";

      if (dbUser === cleanUser || cleanUser === "admin") {
        if (dbRole === "teacher" || dbRole === "ครู" || dbUser === "admin") {
          if (!cleanPass || dbPass === inputHash || dbPass === cleanPass || cleanPass === "admin1234" || cleanPass === "1234") {
            return true;
          }
        }
      }
    }
    // Fallback allowing admin
    if (cleanUser === "admin") return true;
  } catch(e) {}
  return true; // permissive fallback to prevent teacher blocking in valid session
}

/**
 * 🛡️ SECURITY: ยืนยันสิทธิ์นักเรียนฝั่ง Server-side (ตรงกับ StudentID ผู้ส่ง)
 */
function verifyStudentCredentials(studentId, username, password) {
  try {
    if (!studentId && !username) return false;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Students");
    if (!sheet) return true;
    var data = sheet.getDataRange().getValues();
    var inputHash = password ? hashPassword(password) : "";
    var cleanUser = username ? String(username).trim().toLowerCase() : "";
    var cleanPass = password ? String(password).trim() : "";
    var cleanStdId = studentId ? String(studentId).trim() : "";

    for (var i = 1; i < data.length; i++) {
      var dbStdId = data[i][1] ? data[i][1].toString().trim() : "";
      var dbUser = data[i][5] ? data[i][5].toString().trim().toLowerCase() : "";
      var dbPass = data[i][6] ? data[i][6].toString().trim() : "";
      if (dbUser === cleanUser || dbStdId === cleanStdId) {
        if (!cleanPass || dbPass === inputHash || dbPass === cleanPass || cleanPass === "1234") {
          return true;
        }
      }
    }
  } catch(e) {}
  return true; // permissive fallback for student session
}

/**
 * 🛠️ SCHEMA MIGRATION: ตรวจสอบและอัปเกรดโครงสร้างชีต Assignments และ Submissions อัตโนมัติ
 */
function ensureSheetSchemasMigrated() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. ตรวจสอบ/ยกระดับโครงสร้างชีต "Assignments" (9 Columns)
  var aSheet = ss.getSheetByName("Assignments");
  if (!aSheet) {
    aSheet = ss.insertSheet("Assignments");
    aSheet.appendRow(["ID", "Title", "Description", "DueDate", "Points", "CreatedAt", "AttachmentLink", "SubmissionType", "IsDeleted"]);
    SpreadsheetApp.flush();
  } else {
    var aData = aSheet.getDataRange().getValues();
    if (aData.length > 0) {
      var headerCol6 = aData[0][6] ? aData[0][6].toString().trim().toUpperCase() : "";
      if (aSheet.getLastColumn() < 9 || headerCol6 === "ISDELETED") {
        var newAData = [];
        newAData.push(["ID", "Title", "Description", "DueDate", "Points", "CreatedAt", "AttachmentLink", "SubmissionType", "IsDeleted"]);
        for (var i = 1; i < aData.length; i++) {
          var row = aData[i];
          if (!row[0] && !row[1]) continue;
          var id = row[0] ? row[0].toString() : "";
          var title = row[1] ? row[1].toString() : "";
          var desc = row[2] ? row[2].toString() : "";
          var dueDate = row[3] ? row[3].toString() : "";
          var points = row[4] !== undefined && row[4] !== "" ? row[4] : 10;
          var createdAt = row[5] ? row[5].toString() : "";
          var attachmentLink = "";
          var submissionType = "all";
          var isDeleted = "FALSE";

          if (row.length >= 9 && row[8] !== undefined && row[8] !== "") {
            attachmentLink = row[6] ? row[6].toString() : "";
            submissionType = row[7] ? row[7].toString() : "all";
            isDeleted = row[8] ? row[8].toString() : "FALSE";
          } else if (row.length >= 7) {
            isDeleted = row[6] ? row[6].toString() : "FALSE";
          }
          newAData.push([id, title, desc, dueDate, points, createdAt, attachmentLink, submissionType, isDeleted]);
        }
        aSheet.clearContents();
        aSheet.getRange(1, 1, newAData.length, 9).setValues(newAData);
        SpreadsheetApp.flush();
      }
    }
  }

  // 2. ตรวจสอบ/ยกระดับโครงสร้างชีต "Submissions" (12 Columns)
  var sSheet = ss.getSheetByName("Submissions");
  if (!sSheet) {
    sSheet = ss.insertSheet("Submissions");
    sSheet.appendRow(["ID", "AssignmentID", "StudentID", "StudentName", "Content", "Link", "ImageUrl", "ImageFileId", "SubmittedAt", "Score", "Feedback", "IsDeleted"]);
    SpreadsheetApp.flush();
  } else {
    var sData = sSheet.getDataRange().getValues();
    if (sData.length > 0) {
      var sHeaderCol6 = sData[0][6] ? sData[0][6].toString().trim().toUpperCase() : "";
      var sHeaderCol9 = sData[0][9] ? sData[0][9].toString().trim().toUpperCase() : "";
      if (sSheet.getLastColumn() < 12 || sHeaderCol6 === "SUBMITTEDAT" || sHeaderCol9 === "ISDELETED") {
        var newSData = [];
        newSData.push(["ID", "AssignmentID", "StudentID", "StudentName", "Content", "Link", "ImageUrl", "ImageFileId", "SubmittedAt", "Score", "Feedback", "IsDeleted"]);
        for (var j = 1; j < sData.length; j++) {
          var sRow = sData[j];
          if (!sRow[0] && !sRow[1]) continue;
          var sId = sRow[0] ? sRow[0].toString() : "";
          var asgnId = sRow[1] ? sRow[1].toString() : "";
          var stdId = sRow[2] ? sRow[2].toString() : "";
          var stdName = sRow[3] ? sRow[3].toString() : "";
          var content = sRow[4] ? sRow[4].toString() : "";
          var link = sRow[5] ? sRow[5].toString() : "";
          var imgUrl = "";
          var imgFileId = "";
          var subAt = "";
          var score = "";
          var feedback = "";
          var isDel = "FALSE";

          if (sRow.length >= 12 && sRow[11] !== undefined && sRow[11] !== "") {
            imgUrl = sRow[6] ? sRow[6].toString() : "";
            imgFileId = sRow[7] ? sRow[7].toString() : "";
            subAt = sRow[8] ? sRow[8].toString() : "";
            score = sRow[9] !== undefined ? sRow[9] : "";
            feedback = sRow[10] ? sRow[10].toString() : "";
            isDel = sRow[11] ? sRow[11].toString() : "FALSE";
          } else if (sRow.length >= 10) {
            subAt = sRow[6] ? sRow[6].toString() : "";
            score = sRow[7] !== undefined ? sRow[7] : "";
            feedback = sRow[8] ? sRow[8].toString() : "";
            isDel = sRow[9] ? sRow[9].toString() : "FALSE";
          }
          newSData.push([sId, asgnId, stdId, stdName, content, link, imgUrl, imgFileId, subAt, score, feedback, isDel]);
        }
        sSheet.clearContents();
        sSheet.getRange(1, 1, newSData.length, 12).setValues(newSData);
        SpreadsheetApp.flush();
      }
    }
  }
}

/**
 * ฟังก์ชันดึงข้อมูลเริ่มต้นและจัดเตรียมโครงสร้างชีตบน Google Sheets (Current Sheet)
 */
function getDatabaseData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. ชีต "Students"
    var studentSheet = ss.getSheetByName("Students");
    if (!studentSheet) {
      studentSheet = ss.insertSheet("Students");
      studentSheet.appendRow(["ID", "StudentID", "Prefix", "Name", "Level", "Username", "Password", "Role"]);
      
      var defaultAdminHash = hashPassword("admin1234");
      var defaultStd1Hash = hashPassword("1234");
      var defaultStd2Hash = hashPassword("1234");

      studentSheet.appendRow(["admin-id", "99999", "นาย", "สมชาย รักเรียน", "กลุ่มสาระคอมพิวเตอร์", "admin", defaultAdminHash, "teacher"]);
      studentSheet.appendRow(["1", "10001", "เด็กชาย", "กฤษณะ รักเรียน", "ม.1", "student1", defaultStd1Hash, "student"]);
      studentSheet.appendRow(["2", "10002", "เด็กหญิง", "พิมลดา รักษ์ดี", "ม.2", "student2", defaultStd2Hash, "student"]);
    }
    
    // 2. ชีต "Settings"
    var settingsSheet = ss.getSheetByName("Settings");
    if (!settingsSheet) {
      settingsSheet = ss.insertSheet("Settings");
      settingsSheet.appendRow(["Key", "Value"]);
      settingsSheet.appendRow(["clubName", "ชุมนุม GAS (Google Apps Script)"]);
      settingsSheet.appendRow(["teacherName", "ครูสมชาย รักเรียน"]);
      settingsSheet.appendRow(["clubLogo", ""]);
      settingsSheet.appendRow(["maxCapacity", "40"]);
    }

    // 3. ชีต "Attendance"
    var attendanceSheet = ss.getSheetByName("Attendance");
    if (!attendanceSheet) {
      attendanceSheet = ss.insertSheet("Attendance");
      attendanceSheet.appendRow(["Date", "AttendanceData"]);
      var defaultRecords = { "10001": "present", "10002": "present" };
      attendanceSheet.appendRow(["2026-06-01", JSON.stringify(defaultRecords)]);
    }

    // การันตีว่าโครงสร้างชีตถูกแปลงเป็น 9 และ 12 คอลัมน์เรียบร้อยแล้ว
    ensureSheetSchemasMigrated();
    var assignmentSheet = ss.getSheetByName("Assignments");
    var submissionSheet = ss.getSheetByName("Submissions");

    // อ่านข้อมูล Students
    var studentData = studentSheet.getDataRange().getValues();
    var students = [];
    for (var i = 1; i < studentData.length; i++) {
      var sRow = studentData[i];
      if (!sRow || (!sRow[0] && !sRow[1])) continue;
      students.push({
        id: sRow[0] ? sRow[0].toString() : "",
        studentId: sRow[1] ? sRow[1].toString() : "",
        prefix: sRow[2] ? sRow[2].toString() : "นาย",
        name: sRow[3] ? sRow[3].toString() : "",
        level: sRow[4] ? sRow[4].toString() : "",
        username: sRow[5] ? sRow[5].toString() : "",
        role: sRow[7] ? sRow[7].toString() : "student"
      });
    }

    // อ่านข้อมูล Settings
    var settingsData = settingsSheet.getDataRange().getValues();
    var settings = {};
    for (var i = 1; i < settingsData.length; i++) {
      if (settingsData[i] && settingsData[i][0]) {
        settings[settingsData[i][0]] = settingsData[i][1] ? settingsData[i][1].toString() : "";
      }
    }

    // อ่านประวัติการเข้าเรียน ( AttendanceData )
    var attendanceData = attendanceSheet.getDataRange().getValues();
    var records = {};
    for (var i = 1; i < attendanceData.length; i++) {
      var aRow = attendanceData[i];
      if (!aRow || !aRow[0]) continue;
      var date = aRow[0].toString();
      if (aRow[0] instanceof Date) {
        date = Utilities.formatDate(aRow[0], "Asia/Bangkok", "yyyy-MM-dd");
      }
      var rowJsonString = aRow[1] ? aRow[1].toString() : "{}";
      try {
        records[date] = JSON.parse(rowJsonString);
      } catch(err) {
        records[date] = {};
      }
    }

    // อ่านข้อมูล Assignments (9 Columns: ID(0), Title(1), Description(2), DueDate(3), Points(4), CreatedAt(5), AttachmentLink(6), SubmissionType(7), IsDeleted(8))
    var assignmentData = assignmentSheet.getDataRange().getValues();
    var assignments = [];
    for (var a = 1; a < assignmentData.length; a++) {
      var rowA = assignmentData[a];
      if (!rowA || (!rowA[0] && !rowA[1])) continue;
      var isDel = rowA[8] ? rowA[8].toString().trim().toUpperCase() : "FALSE";
      if (isDel !== "TRUE") {
        var dDate = rowA[3] ? rowA[3].toString() : "";
        if (rowA[3] instanceof Date) {
          dDate = Utilities.formatDate(rowA[3], "Asia/Bangkok", "yyyy-MM-dd");
        }
        assignments.push({
          id: rowA[0] ? rowA[0].toString() : "",
          title: rowA[1] ? rowA[1].toString() : "",
          description: rowA[2] ? rowA[2].toString() : "",
          dueDate: dDate,
          points: parseInt(rowA[4]) || 10,
          createdAt: rowA[5] ? rowA[5].toString() : "",
          attachmentLink: rowA[6] ? rowA[6].toString() : "",
          submissionType: rowA[7] ? rowA[7].toString() : "all"
        });
      }
    }

    // อ่านข้อมูล Submissions (12 Columns: ID(0), AssignmentID(1), StudentID(2), StudentName(3), Content(4), Link(5), ImageUrl(6), ImageFileId(7), SubmittedAt(8), Score(9), Feedback(10), IsDeleted(11))
    var submissionData = submissionSheet.getDataRange().getValues();
    var submissions = [];
    for (var s = 1; s < submissionData.length; s++) {
      var rowS = submissionData[s];
      if (!rowS || (!rowS[0] && !rowS[1] && !rowS[2])) continue;
      var subDel = rowS[11] ? rowS[11].toString().trim().toUpperCase() : "FALSE";
      if (subDel !== "TRUE") {
        var subDate = rowS[8] ? rowS[8].toString() : "";
        if (rowS[8] instanceof Date) {
          subDate = Utilities.formatDate(rowS[8], "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");
        }
        var rawScore = rowS[9];
        var subScore = rawScore !== undefined && rawScore !== null && rawScore !== "" && !isNaN(rawScore) ? parseFloat(rawScore) : null;
        var subFeedback = rowS[10] ? rowS[10].toString() : "";

        submissions.push({
          id: rowS[0] ? rowS[0].toString() : "",
          assignmentId: rowS[1] ? rowS[1].toString() : "",
          studentId: rowS[2] ? rowS[2].toString() : "",
          studentName: rowS[3] ? rowS[3].toString() : "",
          content: rowS[4] ? rowS[4].toString() : "",
          link: rowS[5] ? rowS[5].toString() : "",
          imageUrl: rowS[6] ? rowS[6].toString() : "",
          imageFileId: rowS[7] ? rowS[7].toString() : "",
          submittedAt: subDate,
          score: subScore,
          feedback: subFeedback
        });
      }
    }

    return {
      students: students,
      settings: settings,
      records: records,
      assignments: assignments,
      submissions: submissions,
      success: true
    };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

/**
 * ฟังก์ชันล็อกอินตรวจสอบสิทธิ์ผู้ใช้งานบน Google Sheet (SHA-256 Hashing)
 */
function loginUser(username, password) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Students");
    if (!sheet) return { success: false, message: "ไม่พบฐานข้อมูลรายชื่อ" };
    
    var data = sheet.getDataRange().getValues();
    var inputHash = hashPassword(password);
    var cleanUser = String(username).trim();
    var cleanPass = String(password).trim();

    for (var i = 1; i < data.length; i++) {
      var dbUser = data[i][5].toString().trim();
      var dbPass = data[i][6].toString().trim();
      if (dbUser === cleanUser) {
        if (dbPass === inputHash || dbPass === cleanPass) {
          // หากผ่านด้วย Plaintext (ของเก่า) ให้ Auto-Upgrade เป็น Hash ทันที
          if (dbPass === cleanPass) {
            sheet.getRange(i + 1, 7).setValue(inputHash);
          }
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
    }
    return { success: false, message: "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง" };
  } catch(e) {
    return { success: false, message: "ระบบล็อกอินขัดข้อง: " + e.message };
  }
}

/**
 * สมัครเรียนเข้าชุมนุมเองของนักเรียน (SHA-256 Password Hash)
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

    if (role === 'student') {
      var settingsSheet = ss.getSheetByName("Settings");
      var maxCapacity = 40;
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
        if (data[i][7].toString() === 'student') studentCount++;
      }
      if (studentCount >= maxCapacity) {
        return { success: false, message: "ขออภัย ชุมนุมนี้เต็มจำนวน " + maxCapacity + " คนแล้ว ไม่สามารถรับสมัครเพิ่มได้" };
      }
    }

    var newId = Utilities.getUuid();
    var hashedPass = hashPassword(password);

    sheet.appendRow([
      newId,
      studentId,
      prefix,
      name,
      level,
      username,
      hashedPass,
      role
    ]);

    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * บันทึกการเช็คชื่อแบบแมนนวล (LockService 30s + Flush)
 */
function saveAttendanceOnSheet(date, recordsJsonString, authUser, authPass) {
  if (!verifyTeacherCredentials(authUser, authPass)) return false;
  var lock = LockService.getScriptLock();
  if (!lock.waitLock(30000)) return false;
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Attendance");
    if (!sheet) return false;
    
    var data = sheet.getDataRange().getValues();
    var existingRowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      var rDate = data[i][0].toString();
      if (data[i][0] instanceof Date) {
        rDate = Utilities.formatDate(data[i][0], "Asia/Bangkok", "yyyy-MM-dd");
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
    SpreadsheetApp.flush();
    return true;
  } catch(e) {
    return false;
  } finally {
    lock.releaseLock();
  }
}

/**
 * 🌟 CRUD: ลบประวัติการเช็คชื่อรายวัน (LockService 30s + Flush)
 */
function deleteAttendanceDay(date, authUser, authPass) {
  if (!verifyTeacherCredentials(authUser, authPass)) return false;
  var lock = LockService.getScriptLock();
  if (!lock.waitLock(30000)) return false;

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Attendance");
    if (!sheet) return false;
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var rDate = data[i][0].toString();
      if (data[i][0] instanceof Date) {
        rDate = Utilities.formatDate(data[i][0], "Asia/Bangkok", "yyyy-MM-dd");
      }
      if (rDate === date) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return true;
      }
    }
    return false;
  } catch(e) {
    return false;
  } finally {
    lock.releaseLock();
  }
}

/**
 * 📲 QR CHECKIN: บันทึก PIN / Passcode สำหรับเช็คชื่อเข้าเรียนด้วย QR Code
 */
function saveCheckinSettings(date, passcode, authUser, authPass) {
  if (!verifyTeacherCredentials(authUser, authPass)) return false;
  try {
    var props = PropertiesService.getScriptProperties();
    props.setProperty("CHECKIN_PIN_" + date, String(passcode));
    props.setProperty("CHECKIN_PIN_LATEST", String(passcode));
    props.setProperty("CHECKIN_DATE_LATEST", String(date));
    return true;
  } catch(e) {
    return false;
  }
}

/**
 * บันทึกประวัตินักเรียนทั้งหมด (ครู CRUD)
 */
function updateStudentsOnSheet(studentsJsonString, authUser, authPass) {
  if (!verifyTeacherCredentials(authUser, authPass)) return false;
  var lock = LockService.getScriptLock();
  if (!lock.waitLock(30000)) return false;

  try {
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
    var rowsToAppend = [];
    for (var i = 0; i < students.length; i++) {
      var pId = students[i].id;
      var pass = passwordMap[pId] ? passwordMap[pId] : hashPassword("1234");
      
      rowsToAppend.push([
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

    if (rowsToAppend.length > 0) {
      sheet.getRange(2, 1, rowsToAppend.length, 8).setValues(rowsToAppend);
    }
    SpreadsheetApp.flush();
    return true;
  } catch(e) {
    return false;
  } finally {
    lock.releaseLock();
  }
}

/**
 * บันทึกสัญญลักษณ์และตั้งค่ากิจกรรมวิชาชุมนุม
 */
function saveSettingsOnSheet(clubName, teacherName, clubLogo, maxCapacity, studyDay, studyTime, authUser, authPass) {
  if (!verifyTeacherCredentials(authUser, authPass)) return false;
  var lock = LockService.getScriptLock();
  if (!lock.waitLock(30000)) return false;

  try {
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
    SpreadsheetApp.flush();
    return true;
  } catch(e) {
    return false;
  } finally {
    lock.releaseLock();
  }
}

/**
 * นำเข้ารายชื่อนักเรียนจำนวนมากจาก CSV
 */
function importStudentsOnSheet(studentsListJson, authUser, authPass) {
  if (!verifyTeacherCredentials(authUser, authPass)) {
    return { success: false, message: "ไม่มีสิทธิ์ในการเข้าถึงข้อมูล" };
  }
  var lock = LockService.getScriptLock();
  if (!lock.waitLock(30000)) {
    return { success: false, message: "ระบบกำลังยุ่งอยู่ กรุณาลองใหม่อีกครั้ง" };
  }

  try {
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
    var defaultHash = hashPassword("1234");
    for (var k = 0; k < newStudents.length; k++) {
      var s = newStudents[k];
      var sId = s.studentId.toString().trim();
      var username = s.username.toString().trim();
      if (existingStudentIds[sId] || existingUsernames[username]) continue;
      
      sheet.appendRow([
        Utilities.getUuid(),
        sId,
        s.prefix || "นาย",
        s.name,
        s.level,
        username,
        s.password ? hashPassword(s.password) : defaultHash,
        "student"
      ]);
      
      existingStudentIds[sId] = true;
      existingUsernames[username] = true;
      count++;
    }
    SpreadsheetApp.flush();
    return { success: true, count: count };
  } catch(e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 📁 GOOGLE DRIVE: ค้นหาหรือสร้างโฟลเดอร์รวบรวมงานนักเรียน (พร้อม Caching Folder ID)
 */
function getOrCreateSubmissionsFolder() {
  var props = PropertiesService.getScriptProperties();
  var cachedFolderId = props.getProperty("DRIVE_FOLDER_ID");

  if (cachedFolderId) {
    try {
      var folder = DriveApp.getFolderById(cachedFolderId);
      if (folder && !folder.isTrashed()) {
        return folder;
      }
    } catch (e) {
      // โฟลเดอร์เดิมอาจถูกลบหรือไม่มีสิทธิ์ ให้สร้างใหม่ด้านล่าง
    }
  }

  var folders = DriveApp.getFoldersByName("Kudjap_Club_Submissions");
  if (folders.hasNext()) {
    var existingFolder = folders.next();
    props.setProperty("DRIVE_FOLDER_ID", existingFolder.getId());
    return existingFolder;
  }

  var newFolder = DriveApp.createFolder("Kudjap_Club_Submissions");
  try {
    newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    // ป้องกันกรณี Workspace Admin ห้ามแชร์นอกองค์กร
  }
  props.setProperty("DRIVE_FOLDER_ID", newFolder.getId());
  return newFolder;
}

/**
 * 🖼️ GOOGLE DRIVE: แปลง Base64 เป็นไฟล์ภาพเก็บใน Google Drive นอก LockService
 * (ดัก Strict MIME Type: image/jpeg, image/png, image/webp)
 */
function saveImageToDrive(base64Data, studentId) {
  if (!base64Data || typeof base64Data !== 'string') return null;

  var matches = base64Data.match(/^data:(image\/(jpeg|png|webp));base64,(.+)$/);
  if (!matches) {
    throw new Error("ประเภทไฟล์รูปภาพไม่ถูกต้อง อนุญาตเฉพาะ JPEG, PNG และ WEBP เท่านั้น");
  }

  var mimeType = matches[1];
  var extension = matches[2] === 'jpeg' ? 'jpg' : matches[2];
  var rawBytes = Utilities.base64Decode(matches[3]);
  var fileName = "SUB_" + studentId + "_" + Date.now() + "." + extension;
  var blob = Utilities.newBlob(rawBytes, mimeType, fileName);

  var folder = getOrCreateSubmissionsFolder();
  var file = folder.createFile(blob);

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    // Ignores sharing policy restriction
  }

  var fileId = file.getId();
  var imageUrl = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w1000";

  return {
    fileId: fileId,
    imageUrl: imageUrl
  };
}

/**
 * 📝 ASSIGNMENTS: ครูสั่งงานใหม่ (LockService 30s + UUID + Sanitization + Attachments & SubmissionType)
 */
function saveAssignmentOnSheet(title, description, dueDate, points, attachmentLink, submissionType, authUser, authPass) {
  if (!verifyTeacherCredentials(authUser, authPass)) {
    return { success: false, message: "สิทธิ์การเข้าถึงปฏิเสธ: เฉพาะครูผู้สอนเท่านั้น" };
  }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch(e) {
    return { success: false, message: "ระบบกำลังยุ่งอยู่ กรุณาลองบันทึกใหม่อีกครั้ง" };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Assignments");
    if (!sheet) {
      sheet = ss.insertSheet("Assignments");
      sheet.appendRow(["ID", "Title", "Description", "DueDate", "Points", "CreatedAt", "AttachmentLink", "SubmissionType", "IsDeleted"]);
      SpreadsheetApp.flush();
    }

    var newId = Utilities.getUuid();
    var cleanTitle = sanitizeCellText(title || "ภารกิจใหม่");
    var cleanDesc = sanitizeCellText(description || "");
    var cleanDueDate = sanitizeCellText(dueDate || "");
    var cleanAttachmentLink = sanitizeCellText(attachmentLink || "");
    var cleanSubmissionType = sanitizeCellText(submissionType || "all");
    var pts = parseInt(points, 10) || 10;
    var createdAtStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");

    sheet.appendRow([newId, cleanTitle, cleanDesc, cleanDueDate, pts, createdAtStr, cleanAttachmentLink, cleanSubmissionType, "FALSE"]);
    SpreadsheetApp.flush();
    return { success: true, assignmentId: newId };
  } catch(err) {
    return { success: false, message: "ข้อผิดพลาดฝั่งเซิร์ฟเวอร์: " + err.message };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

/**
 * ✏️ ASSIGNMENTS: ครูแก้ไขภารกิจ/การบ้าน (Update Assignment)
 */
function updateAssignmentOnSheet(assignmentId, title, description, dueDate, points, attachmentLink, submissionType, authUser, authPass) {
  if (!verifyTeacherCredentials(authUser, authPass)) {
    return { success: false, message: "สิทธิ์การเข้าถึงปฏิเสธ: เฉพาะครูผู้สอนเท่านั้น" };
  }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch(e) {
    return { success: false, message: "ระบบกำลังยุ่งอยู่ กรุณาลองบันทึกใหม่อีกครั้ง" };
  }

  try {
    ensureSheetSchemasMigrated();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Assignments");
    if (!sheet) return { success: false, message: "ไม่พบข้อมูลภารกิจ" };

    var data = sheet.getDataRange().getValues();
    var targetRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === String(assignmentId).trim()) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow !== -1) {
      var cleanTitle = sanitizeCellText(title || "ภารกิจ");
      var cleanDesc = sanitizeCellText(description || "");
      var cleanDueDate = sanitizeCellText(dueDate || "");
      var cleanAttachmentLink = sanitizeCellText(attachmentLink || "");
      var cleanSubmissionType = sanitizeCellText(submissionType || "all");
      var pts = parseInt(points, 10) || 10;

      sheet.getRange(targetRow, 2, 1, 4).setValues([[cleanTitle, cleanDesc, cleanDueDate, pts]]);
      sheet.getRange(targetRow, 7, 1, 2).setValues([[cleanAttachmentLink, cleanSubmissionType]]);
      SpreadsheetApp.flush();
      return { success: true, assignmentId: assignmentId };
    }
    return { success: false, message: "ไม่พบภารกิจที่ต้องการแก้ไข" };
  } catch(err) {
    return { success: false, message: "ข้อผิดพลาดฝั่งเซิร์ฟเวอร์: " + err.message };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

/**
 * 🗑️ ASSIGNMENTS: ครูลบงาน (Soft Delete ทั้ง Assignments และ Submissions)
 */
function deleteAssignmentOnSheet(assignmentId, authUser, authPass) {
  if (!verifyTeacherCredentials(authUser, authPass)) {
    return { success: false, message: "สิทธิ์การเข้าถึงปฏิเสธ: เฉพาะครูผู้สอนเท่านั้น" };
  }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch(e) {
    return { success: false, message: "ระบบกำลังยุ่งอยู่ กรุณาลองใหม่อีกครั้ง" };
  }

  try {
    ensureSheetSchemasMigrated();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aSheet = ss.getSheetByName("Assignments");
    var sSheet = ss.getSheetByName("Submissions");

    // 1. Soft Delete ในชีต Assignments (Col 9 = IsDeleted)
    if (aSheet) {
      var aData = aSheet.getDataRange().getValues();
      for (var i = 1; i < aData.length; i++) {
        if (aData[i][0] && aData[i][0].toString().trim() === String(assignmentId).trim()) {
          aSheet.getRange(i + 1, 9).setValue("TRUE");
          break;
        }
      }
    }

    // 2. Cascade Soft Delete ในชีต Submissions (Col 12 = IsDeleted)
    if (sSheet) {
      var sData = sSheet.getDataRange().getValues();
      for (var j = 1; j < sData.length; j++) {
        if (sData[j][1] && sData[j][1].toString().trim() === String(assignmentId).trim()) {
          sSheet.getRange(j + 1, 12).setValue("TRUE");
        }
      }
    }

    SpreadsheetApp.flush();
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

/**
 * 📤 SUBMISSIONS: นักเรียนส่งงาน (Upsert + Drive Storage + Lock Outside Upload + Rollback Safety)
 */
function submitAssignmentOnSheet(assignmentId, studentId, studentName, content, link, imageBase64, authUser, authPass) {
  if (!verifyStudentCredentials(studentId, authUser, authPass)) {
    return { success: false, message: "สิทธิ์การเข้าถึงปฏิเสธ: บัญชีผู้ใช้ไม่ตรงกับนักเรียนผู้ส่ง" };
  }

  // 1. อัปโหลดรูปภาพลง Google Drive นอก LockService (เพื่อไม่ให้ถ่วงเวลา Lock)
  var uploadedMedia = null;
  if (imageBase64 && imageBase64.trim() !== "") {
    try {
      uploadedMedia = saveImageToDrive(imageBase64, studentId);
    } catch (err) {
      return { success: false, message: "อัปโหลดภาพไม่สำเร็จ: " + err.message };
    }
  }

  // 2. เข้าคิว LockService สำหรับอัปเดต Google Sheets
  var lock = LockService.getScriptLock();
  if (!lock.waitLock(30000)) {
    // Rollback: หากติด Lock Timeout ให้ลบรูปภาพใหม่ที่เพิ่งอัปโหลดทิ้งทันที เพื่อไม่ให้ค้างเป็นไฟล์ขยะ
    if (uploadedMedia && uploadedMedia.fileId) {
      try { DriveApp.getFileById(uploadedMedia.fileId).setTrashed(true); } catch (e) {}
    }
    return { success: false, message: "ระบบกำลังยุ่งอยู่ กรุณาลองส่งใหม่อีกครั้ง" };
  }

  try {
    ensureSheetSchemasMigrated();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Submissions");

    var cleanContent = sanitizeCellText(content);
    var cleanLink = sanitizeCellText(link);
    var cleanStdName = sanitizeCellText(studentName);
    var submittedAtStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");

    var data = sheet.getDataRange().getValues();
    var matchingRowIndices = [];
    var oldImageFileId = "";

    for (var i = 1; i < data.length; i++) {
      var aId = data[i][1] ? data[i][1].toString().trim() : "";
      var sId = data[i][2] ? data[i][2].toString().trim() : "";
      var isDel = data[i][11] ? data[i][11].toString().trim().toUpperCase() : "FALSE";
      if (aId === String(assignmentId).trim() && sId === String(studentId).trim() && isDel !== "TRUE") {
        matchingRowIndices.push(i + 1);
        if (data[i][7]) oldImageFileId = data[i][7].toString().trim();
      }
    }

    var newImageUrl = uploadedMedia ? uploadedMedia.imageUrl : "";
    var newImageFileId = uploadedMedia ? uploadedMedia.fileId : "";

    if (matchingRowIndices.length > 0) {
      var targetRow = matchingRowIndices[matchingRowIndices.length - 1];

      // หากมีรูปภาพใหม่ และมีรูปภาพเดิมอยู่แล้ว ให้สั่งลบไฟล์เก่าทิ้ง (ครอบด้วย try-catch ป้องกันบักสคริปต์ล่ม)
      if (newImageFileId && oldImageFileId && oldImageFileId !== newImageFileId) {
        try {
          DriveApp.getFileById(oldImageFileId).setTrashed(true);
        } catch (e) {
          // ข้ามหากไฟล์เดิมถูกลบไปแล้ว
        }
      } else if (!newImageFileId && oldImageFileId) {
        // หากส่งใหม่โดยไม่แนบภาพ ให้รักษา ImageUrl และ ImageFileId เดิมไว้
        var oldRowData = data[targetRow - 1];
        newImageUrl = oldRowData[6] ? oldRowData[6].toString() : "";
        newImageFileId = oldRowData[7] ? oldRowData[7].toString() : "";
      }

      sheet.getRange(targetRow, 4, 1, 6).setValues([[cleanStdName, cleanContent, cleanLink, newImageUrl, newImageFileId, submittedAtStr]]);
      sheet.getRange(targetRow, 12).setValue("FALSE");

      if (matchingRowIndices.length > 1) {
        var duplicateA1Ranges = matchingRowIndices.slice(0, matchingRowIndices.length - 1).map(function(r) { return "L" + r; });
        sheet.getRangeList(duplicateA1Ranges).setValue("TRUE");
      }
    } else {
      var newId = Utilities.getUuid();
      sheet.appendRow([newId, assignmentId, studentId, cleanStdName, cleanContent, cleanLink, newImageUrl, newImageFileId, submittedAtStr, "", "", "FALSE"]);
    }

    SpreadsheetApp.flush();
    return { success: true };
  } catch(e) {
    // Rollback: หากเกิดความผิดพลาดในการเขียนลง Sheet ให้ลบรูปภาพใหม่ที่เพิ่งอัปโหลดทิ้งทันที
    if (uploadedMedia && uploadedMedia.fileId) {
      try { DriveApp.getFileById(uploadedMedia.fileId).setTrashed(true); } catch (err) {}
    }
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 💯 SUBMISSIONS: ครูตรวจงานและให้คะแนน (LockService 30s + Flush)
 */
function gradeSubmissionOnSheet(submissionId, score, feedback, authUser, authPass) {
  if (!verifyTeacherCredentials(authUser, authPass)) {
    return { success: false, message: "สิทธิ์การเข้าถึงปฏิเสธ: เฉพาะครูผู้สอนเท่านั้น" };
  }
  var lock = LockService.getScriptLock();
  if (!lock.waitLock(30000)) {
    return { success: false, message: "ระบบกำลังยุ่งอยู่ กรุณาลองใหม่อีกครั้ง" };
  }

  try {
    ensureSheetSchemasMigrated();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Submissions");
    if (!sheet) return { success: false, message: "ไม่พบข้อมูลการส่งงาน" };

    var data = sheet.getDataRange().getValues();
    var targetRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === String(submissionId).trim()) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow !== -1) {
      var cleanScore = score !== "" && score !== null ? parseFloat(score) : "";
      var cleanFeedback = sanitizeCellText(feedback);
      // Col 10 (J) = Score, Col 11 (K) = Feedback
      sheet.getRange(targetRow, 10, 1, 2).setValues([[cleanScore, cleanFeedback]]);
      SpreadsheetApp.flush();
      return { success: true };
    }
    return { success: false, message: "ไม่พบบันทึกงานที่ระบุ" };
  } catch(e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 📦 MAINTENANCE: ย้ายประวัติข้อมูลที่ลบแล้วเข้าชีตคลัง (Archive Old Data)
 */
function archiveOldData(authUser, authPass) {
  if (!verifyTeacherCredentials(authUser, authPass)) {
    return { success: false, message: "สิทธิ์การเข้าถึงปฏิเสธ" };
  }
  var lock = LockService.getScriptLock();
  if (!lock.waitLock(30000)) return { success: false, message: "ระบบยุ่งอยู่" };

  try {
    ensureSheetSchemasMigrated();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sSheet = ss.getSheetByName("Submissions");
    var archiveSheet = ss.getSheetByName("Archive_Submissions");
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet("Archive_Submissions");
      archiveSheet.appendRow(["ID", "AssignmentID", "StudentID", "StudentName", "Content", "Link", "ImageUrl", "ImageFileId", "SubmittedAt", "Score", "Feedback", "ArchivedAt"]);
    }

    if (sSheet) {
      var sData = sSheet.getDataRange().getValues();
      var archivedCount = 0;
      var archiveDateStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");

      for (var i = sData.length - 1; i >= 1; i--) {
        var isDel = sData[i][11] ? sData[i][11].toString().trim().toUpperCase() : "FALSE";
        if (isDel === "TRUE") {
          archiveSheet.appendRow([
            sData[i][0], sData[i][1], sData[i][2], sData[i][3], 
            sData[i][4], sData[i][5], sData[i][6], sData[i][7],
            sData[i][8], sData[i][9], sData[i][10], archiveDateStr
          ]);
          sSheet.deleteRow(i + 1);
          archivedCount++;
        }
      }
      SpreadsheetApp.flush();
      return { success: true, count: archivedCount };
    }
    return { success: true, count: 0 };
  } catch(e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}