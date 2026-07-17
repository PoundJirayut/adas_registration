// ============================================================
//  ADAS Conference — Registration Backend (Google Apps Script)
//
//  วิธีติดตั้ง / อัพเดท:
//  1. เปิด Google Sheet → Extensions → Apps Script
//  2. วาง code นี้ทั้งหมดลงใน Code.gs แล้ว Save (Ctrl+S)
//  3. Deploy → Manage deployments → Edit (ดินสอ) → New version → Deploy
//     (ครั้งแรก: Deploy → New deployment → Web app
//       Execute as: Me | Who has access: Anyone, even anonymous)
//  4. Copy Web app URL → วางใน .env ที่ SCRIPT_URL=...
//
//  ต้องมี column "Picture" ใน Google Sheet
// ============================================================

const SHEET_ID     = '19TspRNs1fkeY89CP-lJW8sdXaTCSGUeD8wh8FPpKoww';
const DRIVE_FOLDER = 'ADAS_Registration_Photos';

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

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'ADAS Registration API' }))
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

      // ดาวน์โหลดรูปจาก server แล้วอัพโหลดขึ้น Google Drive
      let driveUrl = '';
      if (imageUrl && col.picture !== -1) {
        driveUrl = uploadToDrive(paperId, imageUrl);
        sheet.getRange(i + 1, col.picture + 1).setValue(driveUrl);
      }

      return {
        success:  true,
        timestamp,
        name:     data[i][col.name],
        title:    data[i][col.title],
        paperId:  rowId,
        imageUrl: driveUrl || imageUrl,
      };
    }

    return { success: false, message: 'Paper ID not found' };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

// ── uploadToDrive ──────────────────────────────────────────────
// Apps Script fetch รูปจาก Railway URL แล้วบันทึกใน Google Drive
function uploadToDrive(paperId, sourceUrl) {
  try {
    const response = UrlFetchApp.fetch(sourceUrl, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      throw new Error('Image fetch failed: HTTP ' + response.getResponseCode());
    }

    const filename = paperId + '_' + Date.now() + '.jpg';
    const blob     = response.getBlob().setName(filename);

    const folders = DriveApp.getFoldersByName(DRIVE_FOLDER);
    const folder  = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER);

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return 'https://drive.google.com/file/d/' + file.getId() + '/view';
  } catch (err) {
    // ถ้า upload ไม่สำเร็จ ใช้ URL เดิมแทน
    console.error('Drive upload failed:', err.message);
    return sourceUrl;
  }
}
