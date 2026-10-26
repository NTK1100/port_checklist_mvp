/**
 * ============================================================================
 * PORTCALL DOCS - GOOGLE APPS SCRIPT WEB APP BRIDGE (google_drive_bridge.gs)
 * ============================================================================
 * HƯỚNG DẪN CÀI ĐẶT TRONG 1 PHÚT (MIỄN PHÍ 100%, KHÔNG CẦN GOOGLE CLOUD CONSOLE):
 *
 * Bước 1: Mở https://script.google.com -> Bấm "Dự án mới" (New project).
 * Bước 2: Xóa hết code mặc định, copy toàn bộ nội dung file này dán vào.
 * Bước 3: Bấm nút "Triển khai" (Deploy) ở góc trên bên phải -> "Tùy chọn triển khai mới" (New deployment).
 * Bước 4: Ở mục "Chọn loại" (hình bánh răng), chọn "Ứng dụng web" (Web app):
 *         - Thực thi dưới dạng (Execute as): "Tôi" (Me)
 *         - Ai có quyền truy cập (Who has access): "Bất kỳ ai" (Anyone)
 * Bước 5: Bấm "Triển khai" (Deploy) -> Cấp quyền truy cập Google Drive -> Copy đường link "URL ứng dụng web"
 *         (có dạng https://script.google.com/macros/s/.../exec) và dán vào biến môi trường GDRIVE_BRIDGE_URL
 *         hoặc dán trực tiếp trong mục Cấu hình Google Drive trên giao diện Admin của PortCall Docs!
 * ============================================================================
 */

const BRIDGE_SECRET_KEY = "PortCall_Secret_2026";
const ROOT_FOLDER_NAME = "PortCall_Checklist_Storage";
const DB_BACKUP_FILENAME = "port_checklist_backup.db";

/**
 * Kiểm tra trạng thái hoạt động khi mở link Web App trên trình duyệt
 */
function doGet(e) {
  const root = getOrCreateRootFolder_();
  return jsonResponse_({
    ok: true,
    service: "PortCall Docs - Google Drive Bridge",
    rootFolderName: root.getName(),
    rootFolderId: root.getId(),
    rootFolderUrl: root.getUrl(),
    message: "Bridge đang hoạt động. Hãy dùng URL này cho biến môi trường GDRIVE_BRIDGE_URL."
  });
}

/**
 * Xử lý các yêu cầu đồng bộ từ FastAPI Backend (SYNC_DB, LOAD_DB, UPLOAD_DOC, DELETE_DOC, DELETE_VESSEL_FOLDER...)
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse_({ ok: false, error: "Empty request body" });
    }

    const data = JSON.parse(e.postData.contents);

    // 1. Kiểm tra khóa bảo mật BRIDGE_SECRET_KEY
    if (!data.secretKey || data.secretKey !== BRIDGE_SECRET_KEY) {
      return jsonResponse_({ ok: false, error: "Unauthorized: Invalid BRIDGE_SECRET_KEY" });
    }

    const action = (data.action || "").toUpperCase();
    const rootFolder = getOrCreateRootFolder_();

    // 2. Action: PING (Kiểm tra kết nối)
    if (action === "PING") {
      return jsonResponse_({
        ok: true,
        rootFolderId: rootFolder.getId(),
        rootFolderUrl: rootFolder.getUrl(),
        rootFolderName: rootFolder.getName()
      });
    }

    // 3. Action: SYNC_DB (Lưu / ghi đè file port_checklist_backup.db trên Google Drive)
    if (action === "SYNC_DB") {
      if (!data.dbBase64) {
        return jsonResponse_({ ok: false, error: "Missing dbBase64" });
      }
      const bytes = Utilities.base64Decode(data.dbBase64);
      const blob = Utilities.newBlob(bytes, "application/x-sqlite3", DB_BACKUP_FILENAME);

      // Xóa các bản backup cũ trùng tên trong thư mục gốc
      const existingFiles = rootFolder.getFilesByName(DB_BACKUP_FILENAME);
      while (existingFiles.hasNext()) {
        const oldFile = existingFiles.next();
        if (!oldFile.isTrashed()) {
          oldFile.setTrashed(true);
        }
      }

      const dbFile = rootFolder.createFile(blob);
      return jsonResponse_({
        ok: true,
        action: "SYNC_DB",
        dbFileId: dbFile.getId(),
        updatedAt: new Date().toISOString()
      });
    }

    // 4. Action: LOAD_DB (Tải file port_checklist_backup.db từ Google Drive về Server khi khởi động)
    if (action === "LOAD_DB") {
      const files = rootFolder.getFilesByName(DB_BACKUP_FILENAME);
      while (files.hasNext()) {
        const f = files.next();
        if (!f.isTrashed()) {
          const bytes = f.getBlob().getBytes();
          const b64 = Utilities.base64Encode(bytes);
          return jsonResponse_({
            ok: true,
            exists: true,
            dbFileId: f.getId(),
            dbBase64: b64,
            updatedAt: f.getLastUpdated().toISOString()
          });
        }
      }
      return jsonResponse_({ ok: true, exists: false });
    }

    // 5. Action: CREATE_VESSEL_FOLDER (Tạo thư mục con cho tàu)
    if (action === "CREATE_VESSEL_FOLDER") {
      const folderName = data.vesselFolderName || "UNKNOWN_VESSEL";
      const vesselFolder = getOrCreateSubFolder_(rootFolder, folderName);
      return jsonResponse_({
        ok: true,
        folderId: vesselFolder.getId(),
        folderUrl: vesselFolder.getUrl(),
        folderName: vesselFolder.getName()
      });
    }

    // 6. Action: UPLOAD_DOC (Tạo thư mục tàu & lưu file chứng từ đã đổi tên chuẩn)
    if (action === "UPLOAD_DOC") {
      const folderName = data.vesselFolderName || "UNKNOWN_VESSEL";
      const fileName = data.fileName || "document.pdf";
      const mimeType = data.mimeType || "application/octet-stream";

      if (!data.fileBase64) {
        return jsonResponse_({ ok: false, error: "Missing fileBase64" });
      }

      const vesselFolder = getOrCreateSubFolder_(rootFolder, folderName);

      // Nếu trong thư mục tàu đã có file trùng tên thì chuyển bản cũ vào thùng rác
      const existingDocs = vesselFolder.getFilesByName(fileName);
      while (existingDocs.hasNext()) {
        const oldDoc = existingDocs.next();
        if (!oldDoc.isTrashed()) {
          oldDoc.setTrashed(true);
        }
      }

      const bytes = Utilities.base64Decode(data.fileBase64);
      const blob = Utilities.newBlob(bytes, mimeType, fileName);
      const savedFile = vesselFolder.createFile(blob);

      return jsonResponse_({
        ok: true,
        fileId: savedFile.getId(),
        viewUrl: savedFile.getUrl(),
        folderId: vesselFolder.getId(),
        folderUrl: vesselFolder.getUrl()
      });
    }

    // 7. Action: DOWNLOAD_DOC (Tải nội dung file từ Drive theo fileId khi container Render vừa restart)
    if (action === "DOWNLOAD_DOC") {
      if (!data.fileId) {
        return jsonResponse_({ ok: false, error: "Missing fileId" });
      }
      const f = DriveApp.getFileById(data.fileId);
      const b64 = Utilities.base64Encode(f.getBlob().getBytes());
      return jsonResponse_({
        ok: true,
        fileId: f.getId(),
        fileName: f.getName(),
        fileBase64: b64
      });
    }

    // 8. Action: RENAME_DOC (Đổi tên file trên Google Drive khi phân loại 1-Click)
    if (action === "RENAME_DOC") {
      if (!data.fileId || !data.newFileName) {
        return jsonResponse_({ ok: false, error: "Missing fileId or newFileName" });
      }
      const f = DriveApp.getFileById(data.fileId);
      f.setName(data.newFileName);
      return jsonResponse_({
        ok: true,
        fileId: f.getId(),
        fileName: f.getName(),
        viewUrl: f.getUrl()
      });
    }

    // 9. Action: DELETE_DOC (Xóa file chứng từ trên Google Drive)
    if (action === "DELETE_DOC") {
      if (data.fileId) {
        try {
          DriveApp.getFileById(data.fileId).setTrashed(true);
          return jsonResponse_({ ok: true, deletedFileId: data.fileId });
        } catch (err) {}
      }
      if (data.vesselFolderName && data.fileName) {
        const sub = findSubFolder_(rootFolder, data.vesselFolderName);
        if (sub) {
          const files = sub.getFilesByName(data.fileName);
          while (files.hasNext()) {
            files.next().setTrashed(true);
          }
        }
      }
      return jsonResponse_({ ok: true });
    }

    // 10. Action: CLEAR_VESSEL_FOLDER (Làm trống các file trong thư mục tàu)
    if (action === "CLEAR_VESSEL_FOLDER") {
      const sub = data.folderId
        ? DriveApp.getFolderById(data.folderId)
        : findSubFolder_(rootFolder, data.vesselFolderName || "");
      if (sub) {
        const files = sub.getFiles();
        while (files.hasNext()) {
          files.next().setTrashed(true);
        }
      }
      return jsonResponse_({ ok: true });
    }

    // 11. Action: DELETE_VESSEL_FOLDER (Chuyển toàn bộ thư mục tàu vào thùng rác)
    if (action === "DELETE_VESSEL_FOLDER") {
      if (data.folderId) {
        try {
          DriveApp.getFolderById(data.folderId).setTrashed(true);
          return jsonResponse_({ ok: true });
        } catch (err) {}
      }
      if (data.vesselFolderName) {
        const sub = findSubFolder_(rootFolder, data.vesselFolderName);
        if (sub) {
          sub.setTrashed(true);
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
  const folders = DriveApp.getFoldersByName(ROOT_FOLDER_NAME);
  while (folders.hasNext()) {
    const f = folders.next();
    if (!f.isTrashed()) {
      return f;
    }
  }
  return DriveApp.createFolder(ROOT_FOLDER_NAME);
}

/**
 * Tìm thư mục con theo tên trong thư mục gốc
 */
function findSubFolder_(parentFolder, subName) {
  if (!subName) return null;
  const folders = parentFolder.getFoldersByName(subName);
  while (folders.hasNext()) {
    const f = folders.next();
    if (!f.isTrashed()) {
      return f;
    }
  }
  return null;
}

/**
 * Tìm hoặc tạo thư mục con theo tên tàu (VD: 2026-10-26_MV_BINH_HOA)
 */
function getOrCreateSubFolder_(parentFolder, subName) {
  const existing = findSubFolder_(parentFolder, subName);
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
