// ============================================================
// Village Capital & Investment — Branch Hub Apps Script
// Paste this entire file into: Branch Pipeline Tracker
//   → Extensions → Apps Script → Code.gs
// Then: Deploy → New Deployment → Web App
//   Execute as: Me
//   Who has access: Anyone
// Copy the deployment URL into each HTML tool file where you
// see: PASTE_YOUR_APPS_SCRIPT_URL_HERE
// ============================================================

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = data.sheet || 'LOA Pipeline';
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return respond({ success: false, error: 'Sheet not found: ' + sheetName });
    }

    var row = data.row || [];

    // ── File upload handling (Competing Offer) ────────────────
    if (data.file) {
      var fileInfo   = data.file;
      var folderId   = fileInfo.folderId;
      var linkColIdx = fileInfo.linkRowIndex || 7; // 0-based column index for the Drive link

      var folder;
      try {
        // Try as a standard Drive folder first
        folder = DriveApp.getFolderById(folderId);
      } catch(folderErr) {
        // May be a Shared Drive — use Drive API advanced service if enabled
        // Fallback: save to root and note the error
        folder = DriveApp.getRootFolder();
        Logger.log('Folder not found, saving to root: ' + folderErr.message);
      }

      var blob     = Utilities.newBlob(
        Utilities.base64Decode(fileInfo.data),
        fileInfo.mimeType,
        fileInfo.name
      );
      var uploaded = folder.createFile(blob);
      uploaded.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      var driveLink = uploaded.getUrl();

      // Replace the placeholder in the row with the real Drive link
      if (row[linkColIdx] === '{{DRIVE_LINK}}') {
        row[linkColIdx] = driveLink;
      }
    }

    // ── Append the row ────────────────────────────────────────
    sheet.appendRow(row);

    return respond({ success: true, driveLink: (data.file ? row[data.file.linkRowIndex || 7] : null) });

  } catch(err) {
    Logger.log(err);
    return respond({ success: false, error: err.message });
  }
}

function doGet(e) {
  return ContentService.createTextOutput('Branch Hub Apps Script is live.');
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// SHARED DRIVE NOTE:
// If your Competing Offer folder is in a Shared Drive, the
// standard DriveApp.getFolderById() may fail. In that case:
//  1. Enable the Drive API advanced service in this project
//     (Services → Drive API → Add)
//  2. Replace the folder lookup above with:
//     var folder = { createFile: function(blob) {
//       var meta = { name: blob.getName(), parents: [folderId] };
//       var file = Drive.Files.create(meta, blob, { supportsAllDrives: true });
//       Drive.Permissions.create({ role:'reader', type:'anyone' },
//         file.id, { supportsAllDrives: true });
//       return { getUrl: function(){ return 'https://drive.google.com/file/d/'+file.id+'/view'; } };
//     }};
// ============================================================
