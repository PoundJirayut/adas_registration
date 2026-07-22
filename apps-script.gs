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

// ── doPost ─────────────────────────────────────────────────────
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'register') {
      const result = registerPaper(
        body.paperId,
        body.imageBase64 || '',
        body.imageUrl    || ''
      );
      return respond(result);
    }

    return respond({ status: 'ok', message: 'APSIPA Registration API' });
  } catch (err) {
    return respond({ success: false, message: 'doPost error: ' + err.message });
  }
}

// ── doGet ──────────────────────────────────────────────────────
function doGet(e) {
  const action = (e && e.parameter) ? e.parameter.action : undefined;

  if (action === 'register') {
    const result = registerPaper(
      e.parameter.paperId,
      '',
      e.parameter.imageUrl || ''
    );
    return respond(result);
  }

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

// ── registerPaper ──────────────────────────────────────────────
function registerPaper(paperId, imageBase64, imageUrl) {
  if (!paperId) return { success: false, message: 'paperId is required' };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: 'Server busy, please try again' };
  }

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
      sheet.getRange(i + 1, col.status + 1, 1, 2).setValues([['Registed', timestamp]]);

      // Upload photo to Google Drive
      let driveUrl   = '';
      let driveError = '';
      if (col.picture === -1) {
        driveError = 'Picture column not found in sheet (expected column named "Picture")';
      } else if (!imageBase64 && !imageUrl) {
        driveError = 'no image provided';
      } else {
        let result;
        if (imageBase64) {
          result = uploadToDriveBase64(paperId, imageBase64);
        } else {
          result = uploadToDriveUrl(paperId, imageUrl);
        }
        if (result.url) {
          driveUrl = result.url;
          sheet.getRange(i + 1, col.picture + 1).setValue(driveUrl);
        } else {
          driveError = result.error;
        }
      }

      return {
        success:    true,
        timestamp,
        name:       data[i][col.name],
        title:      data[i][col.title],
        paperId:    rowId,
        driveUrl,
        driveError: driveError || undefined,
      };
    }

    return { success: false, message: 'Paper ID not found: ' + paperId };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

// ── uploadToDriveBase64 ────────────────────────────────────────
// Decodes base64 image and saves directly to Google Drive.
// Works from any server including localhost.
function uploadToDriveBase64(paperId, imageBase64) {
  try {
    const bytes  = Utilities.base64Decode(imageBase64);
    const blob   = Utilities.newBlob(bytes, 'image/jpeg', paperId + '_' + Date.now() + '.jpg');
    const folder = getOrCreateFolder();
    const file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return { url: 'https://drive.google.com/file/d/' + file.getId() + '/view' };
  } catch (err) {
    return { error: err.message };
  }
}

// ── uploadToDriveUrl ───────────────────────────────────────────
// Fetches the photo from a public URL and saves it to Google Drive.
// Requires the server to be publicly accessible (Railway, not localhost).
function uploadToDriveUrl(paperId, sourceUrl) {
  try {
    const response = UrlFetchApp.fetch(sourceUrl, {
      muteHttpExceptions: true,
      followRedirects:    true,
    });

    const code = response.getResponseCode();
    if (code !== 200) {
      return { error: 'Fetch returned HTTP ' + code + ' — server must be publicly accessible' };
    }

    const blob   = response.getBlob().setName(paperId + '_' + Date.now() + '.jpg');
    const folder = getOrCreateFolder();
    const file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return { url: 'https://drive.google.com/file/d/' + file.getId() + '/view' };
  } catch (err) {
    return { error: err.message };
  }
}

// ── Helpers ────────────────────────────────────────────────────
function getOrCreateFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER);
}

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
