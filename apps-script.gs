// ============================================================
//  APSIPA Conference — Registration Backend (Google Apps Script)
//
//  วิธีติดตั้ง / อัพเดท:
//  1. เปิด Google Sheet → Extensions → Apps Script
//  2. วาง code นี้ทั้งหมดลงใน Code.gs แล้ว Save (Ctrl+S)
//  3. Deploy → Manage deployments → Edit (ดินสอ) → New version → Deploy
//     (ครั้งแรก: Deploy → New deployment → Web app
//       Execute as: Me | Who has access: Anyone, even anonymous)
//  4. Copy Web app URL → วางใน .env ที่ SCRIPT_URL=...
//
//  ตรวจสอบ Google Sheet:
//  - ต้องมี column ชื่อ "Picture" (ถ้าไม่มี Drive upload จะถูกข้าม)
//  - column headers ต้องอยู่ใน row แรก (row 1)
// ============================================================

const SHEET_ID     = '19TspRNs1fkeY89CP-lJW8sdXaTCSGUeD8wh8FPpKoww';
const DRIVE_FOLDER = 'APSIPA_Registration_Photos';

// ── doGet ──────────────────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action;

  // action=register : update Status + Last_status_cheanged
  if (action === 'register') {
    const result = registerPaper(e.parameter.paperId);
    return respond(result);
  }

  // action=setPicture : write Drive URL into Picture column (called after upload)
  if (action === 'setPicture') {
    const result = setPictureUrl(e.parameter.paperId, e.parameter.driveUrl || '');
    return respond(result);
  }

  // action=info : show Drive folder URL for easy navigation
  if (action === 'info') {
    const folders = DriveApp.getFoldersByName(DRIVE_FOLDER);
    const folder  = folders.hasNext() ? folders.next() : null;
    return respond({
      status:    'ok',
      folder:    DRIVE_FOLDER,
      folderUrl: folder
        ? 'https://drive.google.com/drive/folders/' + folder.getId()
        : '(not created yet — upload a photo first)',
      sheetId:   SHEET_ID,
    });
  }

  return respond({ status: 'ok', message: 'APSIPA Registration API' });
}

// ── doPost — รับ base64 image แล้วอัพโหลดขึ้น Drive ────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    if (payload.action === 'uploadImage') {
      const result = uploadBase64ToDrive(payload.paperId, payload.imageBase64 || '');
      return respond(result);
    }

    return respond({ success: false, message: 'Unknown action' });
  } catch (err) {
    return respond({ success: false, message: 'doPost error: ' + err.message });
  }
}

// ── registerPaper ──────────────────────────────────────────────
function registerPaper(paperId) {
  if (!paperId) return { success: false, message: 'paperId is required' };

  try {
    const { sheet, data, headers } = getSheet();
    const col = getColumns(headers);

    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][col.paperId]).trim();
      if (rowId !== String(paperId).trim()) continue;

      if (data[i][col.status] === 'Registed') {
        return { success: false, alreadyRegistered: true, message: 'Already registered' };
      }

      const timestamp = now();
      sheet.getRange(i + 1, col.status + 1).setValue('Registed');
      sheet.getRange(i + 1, col.lastChanged + 1).setValue(timestamp);

      return {
        success:   true,
        timestamp,
        name:      data[i][col.name],
        title:     data[i][col.title],
        paperId:   rowId,
      };
    }

    return { success: false, message: 'Paper ID not found: ' + paperId };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

// ── setPictureUrl ──────────────────────────────────────────────
function setPictureUrl(paperId, driveUrl) {
  if (!paperId || !driveUrl) return { success: false, message: 'paperId and driveUrl required' };

  try {
    const { sheet, data, headers } = getSheet();
    const col = getColumns(headers);

    if (col.picture === -1) return { success: false, message: 'No "Picture" column in sheet' };

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][col.paperId]).trim() !== String(paperId).trim()) continue;
      sheet.getRange(i + 1, col.picture + 1).setValue(driveUrl);
      return { success: true };
    }

    return { success: false, message: 'Paper ID not found' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ── uploadBase64ToDrive ────────────────────────────────────────
function uploadBase64ToDrive(paperId, base64) {
  if (!paperId || !base64) return { success: false, error: 'paperId and imageBase64 required' };

  try {
    const decoded = Utilities.base64Decode(base64);
    const blob    = Utilities.newBlob(decoded, 'image/jpeg', paperId + '_' + Date.now() + '.jpg');

    const folders = DriveApp.getFoldersByName(DRIVE_FOLDER);
    const folder  = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER);

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const driveUrl = 'https://drive.google.com/file/d/' + file.getId() + '/view';
    return { success: true, driveUrl };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── Helpers ────────────────────────────────────────────────────
function getSheet() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheets()[0];
  const data  = sheet.getDataRange().getValues();
  return { sheet, data, headers: data[0] };
}

function getColumns(headers) {
  return {
    name:        headers.indexOf('Name'),
    title:       headers.indexOf('Peper-Title'),
    paperId:     headers.indexOf('Paper-ID'),
    status:      headers.indexOf('Status'),
    lastChanged: headers.indexOf('Last_status_cheanged'),
    picture:     headers.indexOf('Picture'),
  };
}

function now() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
