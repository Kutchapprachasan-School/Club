# Google Apps Script (GAS) Production Standards

1. **NO BASE64 IN SHEETS:** 
   - Never write Base64 image strings to cells (max 50k chars limit).
   - Always upload images to Google Drive first using `DriveApp`, then save only the file ID and `https://drive.google.com/thumbnail?id={FILE_ID}&sz=w1000` URL to the sheet.
2. **OUTSIDE-LOCK I/O:** 
   - Perform slow operations (like `DriveApp.createFile()`) OUTSIDE/BEFORE calling `LockService.getScriptLock()`.
   - Keep `LockService` execution time under milliseconds.
3. **MANDATORY CONCURRENCY PROTECTION:** 
   - Wrap all CUD sheet operations inside `LockService.getScriptLock().waitLock(30000)`.
   - Always call `SpreadsheetApp.flush()` before `lock.releaseLock()`.
4. **UPSERT & CLEANUP:** 
   - On submission update, wrap old Drive file deletion (`file.setTrashed(true)`) in a `try-catch` block to prevent crashes if the file was already deleted.
5. **FORMULA INJECTION PROTECTION:** 
   - Prepend `'` to cell text if it starts with `=`, `+`, `-`, or `@`.
6. **STRICT DATE FORMAT:** 
   - Send dates across `google.script.run` only as `YYYY-MM-DD` (Asia/Bangkok) strings.
