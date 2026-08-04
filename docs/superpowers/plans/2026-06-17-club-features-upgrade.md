# Kudjap School Club Features Upgrade Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Upgrade Kudjap School Club management system with server-side authorization, member registration limit, CSV bulk import, QR Code self check-in, and premium glassmorphism UI.

**Architecture:** 
- Secure write operations by verifying teacher credentials on the server side in `Code.gs`.
- Add `maxCapacity`, `checkinPasscode`, and `checkinDate` to the `Settings` sheet.
- Parse CSV files client-side using JavaScript to preview data before batch saving.
- Generate dynamic check-in QR codes via public APIs and read query parameters in the React frontend.
- Refine Tailwind CSS visual styles with soft gradients, transitions, and clear typography.

**Tech Stack:** 
- Google Apps Script (GAS)
- Google Sheets (Database)
- React 18 (Client-side)
- Tailwind CSS (CDN)
- QR Server API (`https://api.qrserver.com`)

## Global Constraints
- Do not introduce external server side dependencies (all code must run natively in Apps Script V8).
- Do not use client side package managers for libraries (use browser native APIs where possible).
- Maintain backward compatibility with the existing sheets schema (`Students`, `Settings`, `Attendance`).
- Maintain offline/demo mode functionality when `google.script.run` is unavailable.

---

### Task 1: Server-side Security Verification & Club Limits

**Files:**
- Modify: `Code.gs`

**Interfaces:**
- Produces: Role verification helper function `verifyTeacherCredentials(username, password)`.
- Modifies signatures of:
  - `saveAttendanceOnSheet(date, recordsJsonString)` ➡️ `saveAttendanceOnSheet(date, recordsJsonString, authUser, authPass)`
  - `deleteAttendanceDay(date)` ➡️ `deleteAttendanceDay(date, authUser, authPass)`
  - `updateStudentsOnSheet(studentsJsonString)` ➡️ `updateStudentsOnSheet(studentsJsonString, authUser, authPass)`
  - `saveSettingsOnSheet(clubName, teacherName, clubLogo)` ➡️ `saveSettingsOnSheet(clubName, teacherName, clubLogo, authUser, authPass)`
  - `registerUser(studentId, prefix, name, level, username, password, role)`: adds check for `maxCapacity`.

- [ ] **Step 1: Write helper function to verify teacher credentials**
  Open [Code.gs](file:///c:/dev/05%20ระบบชุมนุม/Code.gs). Implement the role verification logic.
  ```javascript
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
  ```

- [ ] **Step 2: Update database write functions to verify credentials**
  Modify the existing write functions in [Code.gs](file:///c:/dev/05%20ระบบชุมนุม/Code.gs) to check credentials before making modifications.
  ```javascript
  function saveAttendanceOnSheet(date, recordsJsonString, authUser, authPass) {
    if (!verifyTeacherCredentials(authUser, authPass)) return false;
    // ... existing save logic ...
  }
  function deleteAttendanceDay(date, authUser, authPass) {
    if (!verifyTeacherCredentials(authUser, authPass)) return false;
    // ... existing delete logic ...
  }
  function updateStudentsOnSheet(studentsJsonString, authUser, authPass) {
    if (!verifyTeacherCredentials(authUser, authPass)) return false;
    // ... existing update logic ...
  }
  function saveSettingsOnSheet(clubName, teacherName, clubLogo, authUser, authPass) {
    if (!verifyTeacherCredentials(authUser, authPass)) return false;
    // ... existing save logic ...
  }
  ```

- [ ] **Step 3: Update `registerUser` with `maxCapacity` constraint**
  Modify `registerUser` to verify student count against the `maxCapacity` stored in Settings sheet before appending rows.
  ```javascript
  // Read maxCapacity from Settings
  var settingsSheet = ss.getSheetByName("Settings");
  var maxCapacity = 999; // Default
  if (settingsSheet) {
    var sData = settingsSheet.getDataRange().getValues();
    for (var j = 1; j < sData.length; j++) {
      if (sData[j][0].toString() === "maxCapacity") {
        maxCapacity = parseInt(sData[j][1]) || 999;
      }
    }
  }
  // Count current students (role === 'student')
  var currentStudentCount = 0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][7].toString() === 'student') {
      currentStudentCount++;
    }
  }
  if (role === 'student' && currentStudentCount >= maxCapacity) {
    return { success: false, message: "ขออภัย ชุมนุมนี้เต็มจำนวน " + maxCapacity + " คนแล้ว ไม่สามารถรับสมัครเพิ่มได้" };
  }
  ```

- [ ] **Step 4: Verify task works by testing fake inputs**
  Check that student registration fails when the limit is exceeded. Verify write functions block unauthorized requests.

---

### Task 2: CSV Bulk Import Feature (Client & Server)

**Files:**
- Modify: `Code.gs`
- Modify: `index.html`

- [ ] **Step 1: Create server-side bulk insert function**
  Open [Code.gs](file:///c:/dev/05%20ระบบชุมนุม/Code.gs). Add a server function `importStudentsOnSheet(studentsListJson, authUser, authPass)`:
  ```javascript
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
  ```

- [ ] **Step 2: Add CSV file uploader and parser in HTML frontend**
  Open [index.html](file:///c:/dev/05%20ระบบชุมนุม/index.html). In the "Manage Students" tab, add an import button, a CSV file parser, and preview table container.
  ```javascript
  const [csvPreview, setCsvPreview] = useState([]);
  const [importStatus, setImportStatus] = useState("");

  const handleCsvUpload = (e) => {
      const file = e.target.files[0];
      if (!file) return;
      
      const reader = new FileReader();
      reader.onload = (evt) => {
          const text = evt.target.result;
          const rows = text.split("\n").map(r => r.trim()).filter(r => r);
          if (rows.length <= 1) {
              showToast("ไฟล์ CSV ไม่มีข้อมูลเพียงพอ", "error");
              return;
          }
          
          // Parse headers and rows
          // Expected headers: รหัสนักเรียน, คำนำหน้า, ชื่อ-นามสกุล, ระดับชั้น
          const parsed = [];
          for (let i = 1; i < rows.length; i++) {
              const cols = rows[i].split(",").map(c => c.trim().replace(/^["']|["']$/g, ''));
              if (cols.length >= 4) {
                  parsed.push({
                      studentId: cols[0],
                      prefix: cols[1],
                      name: cols[2],
                      level: cols[3],
                      username: cols[0], // Default username as student ID
                      password: "1234"    // Default password
                  });
              }
          }
          setCsvPreview(parsed);
      };
      reader.readAsText(file, "UTF-8");
  };
  ```

- [ ] **Step 3: Render CSV Import Preview UI**
  Add JSX block to show the preview modal and import confirmation.
  ```jsx
  {csvPreview.length > 0 && (
      <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-3xl p-6 max-w-2xl w-full max-h-[80vh] flex flex-col shadow-2xl">
              <h3 className="text-lg font-black text-slate-800 mb-2">ตรวจสอบรายชื่อนำเข้า ({csvPreview.length} รายการ)</h3>
              <p className="text-xs text-slate-400 mb-4">ข้อมูลรหัสนักเรียนหรือชื่อผู้ใช้ที่ซ้ำในระบบจะถูกข้ามโดยอัตโนมัติ</p>
              <div className="overflow-y-auto flex-1 border rounded-2xl mb-4 text-xs">
                  <table className="w-full text-left">
                      <thead className="bg-slate-50 border-b">
                          <tr className="font-bold text-slate-600">
                              <th className="p-3">รหัส</th>
                              <th className="p-3">คำนำหน้า</th>
                              <th className="p-3">ชื่อ-นามสกุล</th>
                              <th className="p-3">ระดับชั้น</th>
                          </tr>
                      </thead>
                      <tbody className="divide-y">
                          {csvPreview.map((st, idx) => (
                              <tr key={idx} className="hover:bg-slate-50">
                                  <td className="p-3 font-mono font-bold">{st.studentId}</td>
                                  <td className="p-3">{st.prefix}</td>
                                  <td className="p-3 font-bold">{st.name}</td>
                                  <td className="p-3">{st.level}</td>
                              </tr>
                          ))}
                      </tbody>
                  </table>
              </div>
              <div className="flex gap-2">
                  <button 
                      onClick={() => {
                          setIsSaving(true);
                          if (isOfflineMode) {
                              setStudents(prev => [...prev, ...csvPreview]);
                              showToast(`[จำลอง] นำเข้าสำเร็จ ${csvPreview.length} คน`);
                              setCsvPreview([]);
                              setIsSaving(false);
                          } else {
                              google.script.run
                                  .withSuccessHandler((res) => {
                                      setIsSaving(false);
                                      if (res.success) {
                                          showToast(`นำเข้าสำเร็จ ${res.count} คน`);
                                          loadAllData();
                                          setCsvPreview([]);
                                      } else {
                                          showToast(res.message, "error");
                                      }
                                  })
                                  .withFailureHandler(() => { setIsSaving(false); })
                                  .importStudentsOnSheet(JSON.stringify(csvPreview), currentUser.username, loginPassword);
                          }
                      }}
                      className="flex-1 bg-emerald-600 text-white font-bold py-2.5 rounded-xl text-sm"
                  >
                      ยืนยันนำเข้าข้อมูล
                  </button>
                  <button onClick={() => setCsvPreview([])} className="flex-1 bg-slate-100 text-slate-700 font-bold py-2.5 rounded-xl text-sm">
                      ยกเลิก
                  </button>
              </div>
          </div>
      </div>
  )}
  ```

---

### Task 3: QR Code Self Check-in System

**Files:**
- Modify: `Code.gs`
- Modify: `index.html`

- [ ] **Step 1: Create server-side QR Code setup and validation**
  Open [Code.gs](file:///c:/dev/05%20ระบบชุมนุม/Code.gs). Add:
  - `saveCheckinSettings(date, passcode, authUser, authPass)`
  - `checkInStudent(studentId, date, passcode, username, password)`
  ```javascript
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

  function checkInStudent(studentId, date, passcode, username, password) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      // Verify student credentials
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
      if (!isValidStudent) return { success: false, message: "การยืนยันตัวตนล้มเหลว" };

      // Verify QR Code config
      var setSheet = ss.getSheetByName("Settings");
      var setData = setSheet.getDataRange().getValues();
      var dbPasscode = "";
      var dbDate = "";
      for (var j = 1; j < setData.length; j++) {
        if (setData[j][0].toString() === "checkinPasscode") dbPasscode = setData[j][1].toString().trim();
        if (setData[j][0].toString() === "checkinDate") dbDate = setData[j][1].toString().trim();
      }

      if (dbPasscode !== passcode.trim() || dbDate !== date.trim()) {
        return { success: false, message: "รหัสเช็คชื่อหมดอายุหรือไม่ถูกต้อง" };
      }

      // Record present in Attendance
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
  ```

- [ ] **Step 2: Add dynamic QR Code Generator in Teacher UI**
  Open [index.html](file:///c:/dev/05%20ระบบชุมนุม/index.html). In "Check Attendance" tab, add a QR code generation button and modal.
  ```javascript
  const [qrCodeUrl, setQrCodeUrl] = useState("");
  const [activePasscode, setActivePasscode] = useState("");

  const handleGenerateQR = () => {
      const pin = Math.floor(100000 + Math.random() * 900000).toString();
      setActivePasscode(pin);
      
      const baseUrl = window.location.href.split('?')[0];
      const checkinUrl = `${baseUrl}?checkin=true&date=${currentDate}&passcode=${pin}`;
      const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=250x250&data=${encodeURIComponent(checkinUrl)}`;
      
      setQrCodeUrl(qrUrl);
      
      if (isOfflineMode) {
          showToast(`[จำลอง] รหัสเช็คชื่อคือ ${pin}`);
      } else {
          google.script.run
              .withSuccessHandler((res) => {
                  if (res) showToast("เปิดรับลงทะเบียน QR Code สำเร็จ");
              })
              .saveCheckinSettings(currentDate, pin, currentUser.username, loginPassword);
      }
  };
  ```

- [ ] **Step 3: Handle URL check-in redirection for students**
  In the `useEffect` initialization, detect `checkin=true` query parameters.
  ```javascript
  const [checkinParams, setCheckinParams] = useState(null);

  useEffect(() => {
      const params = new URLSearchParams(window.location.search);
      if (params.get("checkin") === "true") {
          setCheckinParams({
              date: params.get("date"),
              passcode: params.get("passcode")
          });
      }
      loadAllData();
  }, []);
  ```

- [ ] **Step 4: Build Self Check-in confirmation Screen**
  If `checkinParams` is present and the student logs in, redirect them immediately to a check-in confirmation view:
  ```jsx
  {checkinParams && currentUser && currentUser.role === 'student' && (
      <div className="bg-white rounded-3xl p-6 border shadow-xl max-w-md w-full text-center space-y-4">
          <span className="text-emerald-500 font-bold bg-emerald-50 px-3 py-1 rounded-full text-xs">Self Check-in</span>
          <h2 className="text-xl font-black text-slate-800">ยืนยันการเช็คชื่อเข้าเรียน</h2>
          <p className="text-sm text-slate-500">วิชาชุมนุม: <strong className="text-slate-800">{clubName}</strong><br/>ประจำวันที่: <strong className="text-slate-800">{checkinParams.date}</strong></p>
          <button 
              onClick={() => {
                  setIsSaving(true);
                  if (isOfflineMode) {
                      setRecords(prev => {
                          const today = prev[checkinParams.date] || {};
                          today[currentUser.studentId] = "present";
                          return { ...prev, [checkinParams.date]: today };
                      });
                      setIsSaving(false);
                      showToast("เช็คชื่อสำเร็จ (โหมดทดสอบ)");
                      setCheckinParams(null);
                  } else {
                      google.script.run
                          .withSuccessHandler((res) => {
                              setIsSaving(false);
                              if (res.success) {
                                  showToast("เช็คชื่อเข้าเรียนสำเร็จแล้ว!");
                                  setCheckinParams(null);
                                  loadAllData();
                              } else {
                                  showToast(res.message, "error");
                              }
                          })
                          .withFailureHandler(() => { setIsSaving(false); })
                          .checkInStudent(currentUser.studentId, checkinParams.date, checkinParams.passcode, currentUser.username, loginPassword);
                  }
              }}
              className="w-full py-3 bg-emerald-600 text-white font-black rounded-2xl shadow-lg hover:shadow-xl transition-all"
          >
              ยืนยันการเข้าเรียน
          </button>
          <button onClick={() => setCheckinParams(null)} className="w-full text-slate-400 font-bold text-xs">เข้าสู่หน้าหลัก</button>
      </div>
  )}
  ```

---

### Task 4: Premium Glassmorphism UI & Config updates

**Files:**
- Modify: `index.html`

- [ ] **Step 1: Set up HSL tailwind colors and outfit typography**
  Add deep gradients and layout updates using backdrop-filters. Set up custom HSL variables in Tailwind configuration or custom inline classes inside [index.html](file:///c:/dev/05%20ระบบชุมนุม/index.html).
  Use soft purple/indigo gradients: `bg-gradient-to-tr from-slate-900 via-indigo-950 to-slate-900` for login backdrop, and pure white cards with `backdrop-blur-md bg-white/85 border-white/20` for modern glassmorphism.

- [ ] **Step 2: Add maxCapacity input to Settings screen**
  Let teachers configure the club member limits:
  ```jsx
  <div className="bg-white p-6 rounded-3xl border shadow-sm space-y-4">
      <h3 className="text-md font-extrabold text-slate-800">จำกัดจำนวนสมาชิก</h3>
      <div>
          <label className="text-xs font-bold text-slate-400 block mb-1">จำนวนรับสมัครสูงสุด (คน)</label>
          <input 
              type="number"
              value={maxCapacity}
              onChange={(e) => setMaxCapacity(parseInt(e.target.value) || 0)}
              className="w-full px-4 py-2 border rounded-xl"
          />
      </div>
  </div>
  ```
  Ensure `maxCapacity` settings is saved via `saveSettingsOnSheet`.

- [ ] **Step 3: Smooth UI layout polish**
  Refine CSS animations for Toast notifications. Polish student card components and modals for high-fidelity interactive feedback.

---

## Verification Plan

### Automated Tests
*Verify syntax and build correctness locally using clasp before deployment:*
- `npm run status`

### Manual Verification
1. **Security**: Try to trigger `saveAttendanceOnSheet` via console using custom credentials and check that it rejects unauthorized calls.
2. **Limit check**: Set `maxCapacity` to 2 in Settings, try to register a 3rd student, verify registration gets rejected with warning.
3. **CSV Bulk Import**: Upload a sample CSV containing 3 students, review the preview modal, confirm import, check that students are added.
4. **QR Code Check-in**: Open teacher view, click "Generate QR Code", scan/click generated link with student account, confirm check-in, check attendance status updates.
