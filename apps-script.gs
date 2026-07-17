// ============================================================
//  APSIPA Conference — Registration Backend (Google Apps Script)
//
//  วิธีติดตั้ง / อัพเดท:
//  1. เปิด Google Sheet → Extensions → Apps Script
//  2. วาง code นี้ทั้งหมดลงใน Code.gs แล้ว Save (Ctrl+S)
//  3. Deploy → Manage deployments → Edit (ดินสอ) → New version → Deploy
//     (ครั้งแรก: Deploy → New deployment → Web app
//       Execute as: Me | Who has access: Anyone, even anonymous)
//  4. Copy Web app URL → วางใน .env / Railway env vars ที่ SCRIPT_URL=...
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

  if (action === 'register') {
    const result = registerPaper(
      e.parameter.paperId,
      e.parameter.imageUrl || ''
    );
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'info') {
    const folders = DriveApp.getFoldersByName(DRIVE_FOLDER);
    const folder  = folders.hasNext() ? folders.next() : null;
    return ContentService
      .createTextOutput(JSON.stringify({
        status:     'ok',
        folder:     DRIVE_FOLDER,
        folderUrl:  folder ? 'https://drive.google.com/drive/folders/' + folder.getId() : '(not created yet — upload a photo first)',
        sheetId:    SHEET_ID,
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Health check — useful for testing if the deployment is live
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'APSIPA Registration API' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── registerPaper ──────────────────────────────────────────────
function registerPaper(paperId, imageUrl) {
  if (!paperId) return { success: false, message: 'paperId is required' };

  try {
    const ss      = SpreadsheetApp.openById(SHEET_ID);
    const sheet   = ss.getSheets()[0];
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];

    const col = {
      name:        headers.indexOf('Name'),
      title:       headers.indexOf('Peper-Title'),
      paperId:     headers.indexOf('Paper-ID'),
      status:      headers.indexOf('Status'),
      lastChanged: headers.indexOf('Last_status_cheanged'),
      picture:     headers.indexOf('Picture'),
    };

    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][col.paperId]).trim();
      if (rowId !== String(paperId).trim()) continue;

      if (data[i][col.status] === 'Registed') {
        return { success: false, alreadyRegistered: true, message: 'Already registered' };
      }

      const tz        = Session.getScriptTimeZone();
      const timestamp = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');

      sheet.getRange(i + 1, col.status + 1).setValue('Registed');
      sheet.getRange(i + 1, col.lastChanged + 1).setValue(timestamp);

      // Upload photo to Google Drive
      let driveUrl    = '';
      let driveError  = '';
      if (imageUrl) {
        if (col.picture === -1) {
          driveError = 'No "Picture" column found in sheet — skipping Drive upload';
        } else {
          const uploadResult = uploadToDrive(paperId, imageUrl);
          if (uploadResult.url) {
            driveUrl = uploadResult.url;
            sheet.getRange(i + 1, col.picture + 1).setValue(driveUrl);
          } else {
            driveError = uploadResult.error;
            // Still store the original URL as fallback
            sheet.getRange(i + 1, col.picture + 1).setValue(imageUrl + ' (Drive upload failed: ' + uploadResult.error + ')');
          }
        }
      }

      return {
        success:    true,
        timestamp,
        name:       data[i][col.name],
        title:      data[i][col.title],
        paperId:    rowId,
        imageUrl:   driveUrl || imageUrl,
        driveUrl,
        driveError: driveError || undefined,
      };
    }

    return { success: false, message: 'Paper ID not found: ' + paperId };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

// ── uploadToDrive ──────────────────────────────────────────────
// Fetches the photo from the Railway server URL and saves it to Google Drive.
// Returns { url } on success or { error } on failure.
function uploadToDrive(paperId, sourceUrl) {
  try {
    // sourceUrl must be a publicly accessible URL (works on Railway, not on localhost)
    const response = UrlFetchApp.fetch(sourceUrl, {
      muteHttpExceptions: true,
      followRedirects:    true,
    });

    const code = response.getResponseCode();
    if (code !== 200) {
      return { error: 'Fetch failed with HTTP ' + code + ' — is the server publicly accessible?' };
    }

    const filename = paperId + '_' + Date.now() + '.jpg';
    const blob     = response.getBlob().setName(filename);

    const folders = DriveApp.getFoldersByName(DRIVE_FOLDER);
    const folder  = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER);

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return { url: 'https://drive.google.com/file/d/' + file.getId() + '/view' };
  } catch (err) {
    return { error: err.message };
  }
}
