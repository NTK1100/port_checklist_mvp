/**
 * ============================================================================
 * PORTCALL DOCS - GOOGLE APPS SCRIPT WEB APP BRIDGE (google_drive_bridge.gs)
 * ============================================================================
 * CẬP NHẬT MỚI:
 * 1. Lưu file Database (port_checklist_backup.db) vào thư mục hệ thống ẩn riêng biệt:
 *    ".PortCall_System_DB_DoNotDelete" (tách rời hoàn toàn khỏi PortCall_Checklist_Storage
 *    để tránh người dùng xóa nhầm, tự động dọn file .db cũ trong thư mục chính).
 * 2. Hỗ trợ Đồng bộ 2 chiều (SYNC_FOLDER_FROM_DRIVE):
 *    Quét các file do người dùng tải trực tiếp lên thư mục tàu trên Google Drive,
 *    đọc nội dung về Server để phân loại và tự động đổi tên chuẩn ngay trên Drive.
 * ============================================================================
 */

var BRIDGE_SECRET_KEY = "PORTCALL_SECRET_2026";
var ROOT_FOLDER_NAME = "PortCall_Checklist_Storage";
var SYSTEM_DB_FOLDER_NAME = ".PortCall_System_DB_DoNotDelete";
var DB_BACKUP_FILENAME = "port_checklist_backup.db";

/**
 * Kiểm tra trạng thái hoạt động khi mở link Web App trên trình duyệt
 */
function doGet(e) {
  var root = getOrCreateRootFolder_();
  var sysFolder = getOrCreateSystemDbFolder_();
  return jsonResponse_({
    ok: true,
    service: "PortCall Docs - Google Drive 2-Way Bridge v2",
    rootFolderName: root.getName(),
    rootFolderId: root.getId(),
    rootFolderUrl: root.getUrl(),
    systemDbFolderName: sysFolder.getName(),
    message: "Bridge v2 (2-Way Sync & Hidden System DB) đang hoạt động."
  });
}

/**
 * Xử lý các yêu cầu đồng bộ từ FastAPI Backend
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse_({ ok: false, error: "Empty request body" });
    }

    var data = JSON.parse(e.postData.contents);

    // 1. Kiểm tra khóa bảo mật BRIDGE_SECRET_KEY (không phân biệt hoa thường)
    if (!data.secretKey || String(data.secretKey).trim().toUpperCase() !== BRIDGE_SECRET_KEY.toUpperCase()) {
      return jsonResponse_({ ok: false, error: "Unauthorized: Invalid BRIDGE_SECRET_KEY" });
    }

    var action = String(data.action || "").toUpperCase();
    var rootFolder = getOrCreateRootFolder_();

    // 2. Action: PING
    if (action === "PING") {
      var sysDbFolder = getOrCreateSystemDbFolder_();
      cleanupLegacyDbInRootFolder_(rootFolder);
      return jsonResponse_({
        ok: true,
        rootFolderId: rootFolder.getId(),
        rootFolderUrl: rootFolder.getUrl(),
        rootFolderName: rootFolder.getName(),
        systemDbFolderId: sysDbFolder.getId()
      });
    }

    // 3. Action: SYNC_DB (Lưu port_checklist_backup.db vào thư mục ẩn .PortCall_System_DB_DoNotDelete)
    if (action === "SYNC_DB") {
      if (!data.dbBase64) {
        return jsonResponse_({ ok: false, error: "Missing dbBase64" });
      }
      var sysFolder = getOrCreateSystemDbFolder_();
      cleanupLegacyDbInRootFolder_(rootFolder);

      var bytes = Utilities.base64Decode(data.dbBase64);
      var blob = Utilities.newBlob(bytes, "application/x-sqlite3", DB_BACKUP_FILENAME);

      var existingFiles = sysFolder.getFilesByName(DB_BACKUP_FILENAME);
      while (existingFiles.hasNext()) {
        var oldFile = existingFiles.next();
        if (!oldFile.isTrashed()) {
          oldFile.setTrashed(true);
        }
      }

      var dbFile = sysFolder.createFile(blob);
      return jsonResponse_({
        ok: true,
        action: "SYNC_DB",
        dbFileId: dbFile.getId(),
        systemDbFolderName: SYSTEM_DB_FOLDER_NAME,
        updatedAt: new Date().toISOString()
      });
    }

    // 4. Action: LOAD_DB (Tải port_checklist_backup.db từ .PortCall_System_DB_DoNotDelete, có fallback)
    if (action === "LOAD_DB") {
      var sysFolderLoad = getOrCreateSystemDbFolder_();
      var files = sysFolderLoad.getFilesByName(DB_BACKUP_FILENAME);
      while (files.hasNext()) {
        var f = files.next();
        if (!f.isTrashed()) {
          var dbBytes = f.getBlob().getBytes();
          var b64 = Utilities.base64Encode(dbBytes);
          return jsonResponse_({
            ok: true,
            exists: true,
            dbFileId: f.getId(),
            dbBase64: b64,
            updatedAt: f.getLastUpdated().toISOString()
          });
        }
      }

      // Nếu trong thư mục hệ thống mới chưa có nhưng thư mục gốc cũ có thì di chuyển sang thư mục ẩn
      var legacyFiles = rootFolder.getFilesByName(DB_BACKUP_FILENAME);
      while (legacyFiles.hasNext()) {
        var lf = legacyFiles.next();
        if (!lf.isTrashed()) {
          var lBytes = lf.getBlob().getBytes();
          var lB64 = Utilities.base64Encode(lBytes);
          var lBlob = Utilities.newBlob(lBytes, "application/x-sqlite3", DB_BACKUP_FILENAME);
          var migratedFile = sysFolderLoad.createFile(lBlob);
          lf.setTrashed(true);
          return jsonResponse_({
            ok: true,
            exists: true,
            dbFileId: migratedFile.getId(),
            dbBase64: lB64,
            updatedAt: migratedFile.getLastUpdated().toISOString()
          });
        }
      }

      return jsonResponse_({ ok: true, exists: false });
    }

    // 5. Action: CREATE_VESSEL_FOLDER
    if (action === "CREATE_VESSEL_FOLDER") {
      var folderName = data.vesselFolderName || "UNKNOWN_VESSEL";
      var vesselFolder = getOrCreateSubFolder_(rootFolder, folderName);
      return jsonResponse_({
        ok: true,
        folderId: vesselFolder.getId(),
        folderUrl: vesselFolder.getUrl(),
        folderName: vesselFolder.getName()
      });
    }

    // 6. Action: RENAME_VESSEL_FOLDER (Khi sửa tên tàu hoặc ngày ETA)
    if (action === "RENAME_VESSEL_FOLDER") {
      var targetFolder = null;
      if (data.folderId) {
        try {
          targetFolder = DriveApp.getFolderById(data.folderId);
        } catch (err) {}
      }
      if (!targetFolder && data.oldFolderName) {
        targetFolder = findSubFolder_(rootFolder, data.oldFolderName);
      }
      if (targetFolder && data.newFolderName) {
        targetFolder.setName(data.newFolderName);
        return jsonResponse_({
          ok: true,
          folderId: targetFolder.getId(),
          folderUrl: targetFolder.getUrl(),
          folderName: targetFolder.getName()
        });
      }
      return jsonResponse_({ ok: false, error: "Folder not found" });
    }

    // 7. Action: UPLOAD_DOC
    if (action === "UPLOAD_DOC") {
      var vFolderName = data.vesselFolderName || "UNKNOWN_VESSEL";
      var fileName = data.fileName || "document.pdf";
      var mimeType = data.mimeType || "application/octet-stream";

      if (!data.fileBase64) {
        return jsonResponse_({ ok: false, error: "Missing fileBase64" });
      }

      var vFolder = getOrCreateSubFolder_(rootFolder, vFolderName);
      var existingDocs = vFolder.getFilesByName(fileName);
      while (existingDocs.hasNext()) {
        var oldDoc = existingDocs.next();
        if (!oldDoc.isTrashed()) {
          oldDoc.setTrashed(true);
        }
      }

      var docBytes = Utilities.base64Decode(data.fileBase64);
      var docBlob = Utilities.newBlob(docBytes, mimeType, fileName);
      var savedFile = vFolder.createFile(docBlob);

      return jsonResponse_({
        ok: true,
        fileId: savedFile.getId(),
        viewUrl: savedFile.getUrl(),
        folderId: vFolder.getId(),
        folderUrl: vFolder.getUrl()
      });
    }

    // 8. Action: SYNC_FOLDER_FROM_DRIVE (Đồng bộ 2 chiều: Quét file trong thư mục tàu trên Google Drive)
    if (action === "SYNC_FOLDER_FROM_DRIVE") {
      var syncSub = null;
      if (data.folderId) {
        try {
          syncSub = DriveApp.getFolderById(data.folderId);
          if (syncSub.isTrashed()) syncSub = null;
        } catch (err) {
          syncSub = null;
        }
      }
      if (!syncSub && data.vesselFolderName) {
        syncSub = getOrCreateSubFolder_(rootFolder, data.vesselFolderName);
      }
      if (!syncSub) {
        return jsonResponse_({ ok: false, error: "Vessel folder not found on Google Drive" });
      }

      var knownIds = data.knownFileIds || [];
      var knownMap = {};
      for (var k = 0; k < knownIds.length; k++) {
        knownMap[String(knownIds[k])] = true;
      }

      var resultFiles = [];
      var folderFiles = syncSub.getFiles();
      while (folderFiles.hasNext()) {
        var df = folderFiles.next();
        if (df.isTrashed()) continue;
        var fId = df.getId();
        var fName = df.getName();
        var fSize = df.getSize();
        var item = {
          fileId: fId,
          fileName: fName,
          mimeType: df.getMimeType(),
          size: fSize,
          viewUrl: df.getUrl(),
          updatedAt: df.getLastUpdated().toISOString()
        };

        // Nếu là file mới người dùng vừa bỏ trực tiếp lên Drive (chưa có trong DB) và <= 15MB
        if (!knownMap[fId] && fSize <= 15 * 1024 * 1024) {
          try {
            item.fileBase64 = Utilities.base64Encode(df.getBlob().getBytes());
          } catch (readErr) {
            item.readError = String(readErr);
          }
        }
        resultFiles.push(item);
      }

      return jsonResponse_({
        ok: true,
        folderId: syncSub.getId(),
        folderUrl: syncSub.getUrl(),
        folderName: syncSub.getName(),
        files: resultFiles
      });
    }

    // 9. Action: DOWNLOAD_DOC
    if (action === "DOWNLOAD_DOC") {
      if (!data.fileId) {
        return jsonResponse_({ ok: false, error: "Missing fileId" });
      }
      var dlFile = DriveApp.getFileById(data.fileId);
      var dlB64 = Utilities.base64Encode(dlFile.getBlob().getBytes());
      return jsonResponse_({
        ok: true,
        fileId: dlFile.getId(),
        fileName: dlFile.getName(),
        viewUrl: dlFile.getUrl(),
        fileBase64: dlB64
      });
    }

    // 10. Action: RENAME_DOC (Đổi tên trực tiếp file trên Google Drive, không tạo bản trùng lặp)
    if (action === "RENAME_DOC") {
      if (!data.fileId || !data.newFileName) {
        return jsonResponse_({ ok: false, error: "Missing fileId or newFileName" });
      }
      var rnFile = DriveApp.getFileById(data.fileId);
      // Nếu trong cùng thư mục đã có file khác trùng với tên mới và yêu cầu ghi đè
      if (data.replaceExistingInFolder) {
        var parents = rnFile.getParents();
        while (parents.hasNext()) {
          var pFolder = parents.next();
          var dupFiles = pFolder.getFilesByName(data.newFileName);
          while (dupFiles.hasNext()) {
            var dup = dupFiles.next();
            if (dup.getId() !== rnFile.getId() && !dup.isTrashed()) {
              dup.setTrashed(true);
            }
          }
        }
      }
      rnFile.setName(data.newFileName);
      return jsonResponse_({
        ok: true,
        fileId: rnFile.getId(),
        fileName: rnFile.getName(),
        viewUrl: rnFile.getUrl()
      });
    }

    // 11. Action: DELETE_DOC
    if (action === "DELETE_DOC") {
      if (data.fileId) {
        try {
          DriveApp.getFileById(data.fileId).setTrashed(true);
          return jsonResponse_({ ok: true, deletedFileId: data.fileId });
        } catch (err) {}
      }
      if (data.vesselFolderName && data.fileName) {
        var subDel = findSubFolder_(rootFolder, data.vesselFolderName);
        if (subDel) {
          var delFiles = subDel.getFilesByName(data.fileName);
          while (delFiles.hasNext()) {
            delFiles.next().setTrashed(true);
          }
        }
      }
      return jsonResponse_({ ok: true });
    }

    // 12. Action: CLEAR_VESSEL_FOLDER
    if (action === "CLEAR_VESSEL_FOLDER") {
      var subClear = data.folderId
        ? DriveApp.getFolderById(data.folderId)
        : findSubFolder_(rootFolder, data.vesselFolderName || "");
      if (subClear) {
        var cFiles = subClear.getFiles();
        while (cFiles.hasNext()) {
          cFiles.next().setTrashed(true);
        }
      }
      return jsonResponse_({ ok: true });
    }

    // 13. Action: DELETE_VESSEL_FOLDER
    if (action === "DELETE_VESSEL_FOLDER") {
      if (data.folderId) {
        try {
          DriveApp.getFolderById(data.folderId).setTrashed(true);
          return jsonResponse_({ ok: true });
        } catch (err) {}
      }
      if (data.vesselFolderName) {
        var subFolderDel = findSubFolder_(rootFolder, data.vesselFolderName);
        if (subFolderDel) {
          subFolderDel.setTrashed(true);
        }
      }
      return jsonResponse_({ ok: true });
    }

    return jsonResponse_({ ok: false, error: "Unsupported action: " + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Tìm hoặc tạo thư mục gốc PortCall_Checklist_Storage trên Google Drive
 */
function getOrCreateRootFolder_() {
  var folders = DriveApp.getFoldersByName(ROOT_FOLDER_NAME);
  while (folders.hasNext()) {
    var f = folders.next();
    if (!f.isTrashed()) {
      return f;
    }
  }
  return DriveApp.createFolder(ROOT_FOLDER_NAME);
}

/**
 * Tìm hoặc tạo thư mục hệ thống ẩn .PortCall_System_DB_DoNotDelete để lưu port_checklist_backup.db
 */
function getOrCreateSystemDbFolder_() {
  var folders = DriveApp.getFoldersByName(SYSTEM_DB_FOLDER_NAME);
  while (folders.hasNext()) {
    var f = folders.next();
    if (!f.isTrashed()) {
      return f;
    }
  }
  return DriveApp.createFolder(SYSTEM_DB_FOLDER_NAME);
}

/**
 * Dọn dẹp file port_checklist_backup.db cũ nếu còn nằm trong PortCall_Checklist_Storage
 */
function cleanupLegacyDbInRootFolder_(rootFolder) {
  try {
    var oldDbs = rootFolder.getFilesByName(DB_BACKUP_FILENAME);
    while (oldDbs.hasNext()) {
      var f = oldDbs.next();
      if (!f.isTrashed()) {
        f.setTrashed(true);
      }
    }
  } catch (e) {}
}

/**
 * Tìm thư mục con theo tên trong thư mục gốc
 */
function findSubFolder_(parentFolder, subName) {
  if (!subName) return null;
  var folders = parentFolder.getFoldersByName(subName);
  while (folders.hasNext()) {
    var f = folders.next();
    if (!f.isTrashed()) {
      return f;
    }
  }
  return null;
}

/**
 * Tìm hoặc tạo thư mục con theo tên tàu (VD: 2026-10-27_MV_BINH_HOA)
 */
function getOrCreateSubFolder_(parentFolder, subName) {
  var existing = findSubFolder_(parentFolder, subName);
  if (existing) return existing;
  return parentFolder.createFolder(subName);
}

/**
 * Trả về JSON Output chuẩn cho FastAPI
 */
function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
