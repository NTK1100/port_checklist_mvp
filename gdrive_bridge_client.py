import os
import json
import base64
import sqlite3
from datetime import datetime
from typing import Dict, Any, Optional

import requests

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ENV_FILE_PATH = os.path.join(BASE_DIR, ".env")
DB_PATH = os.path.join(BASE_DIR, "port_checklist.db")

_last_sync_info: Dict[str, Any] = {
    "last_sync_at": None,
    "last_sync_status": "IDLE",
    "last_error": "",
    "root_folder_url": "",
}


def _load_dotenv_if_exists() -> None:
    """Đọc file .env cục bộ nếu có (không ghi đè biến môi trường hệ thống trên Render)."""
    if not os.path.exists(ENV_FILE_PATH):
        return
    try:
        with open(ENV_FILE_PATH, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                k = k.strip()
                v = v.strip().strip('"').strip("'")
                if k and k not in os.environ:
                    os.environ[k] = v
    except Exception:
        pass


_load_dotenv_if_exists()


def get_bridge_url() -> str:
    return (os.environ.get("GDRIVE_BRIDGE_URL") or "").strip()


def get_bridge_secret() -> str:
    return (os.environ.get("BRIDGE_SECRET_KEY") or "PortCall_Secret_2026").strip()


def is_bridge_enabled() -> bool:
    url = get_bridge_url()
    return bool(url and url.startswith("https://script.google.com/"))


def set_bridge_config(bridge_url: str, secret_key: str = "PortCall_Secret_2026") -> None:
    """Cho phép Admin cập nhật GDRIVE_BRIDGE_URL và BRIDGE_SECRET_KEY ngay trên giao diện và lưu vào .env."""
    clean_url = (bridge_url or "").strip()
    clean_secret = (secret_key or "PortCall_Secret_2026").strip()
    os.environ["GDRIVE_BRIDGE_URL"] = clean_url
    os.environ["BRIDGE_SECRET_KEY"] = clean_secret

    env_lines = []
    existing_keys = set()
    if os.path.exists(ENV_FILE_PATH):
        try:
            with open(ENV_FILE_PATH, "r", encoding="utf-8") as f:
                for line in f:
                    s = line.strip()
                    if s.startswith("GDRIVE_BRIDGE_URL="):
                        env_lines.append(f'GDRIVE_BRIDGE_URL="{clean_url}"\n')
                        existing_keys.add("GDRIVE_BRIDGE_URL")
                    elif s.startswith("BRIDGE_SECRET_KEY="):
                        env_lines.append(f'BRIDGE_SECRET_KEY="{clean_secret}"\n')
                        existing_keys.add("BRIDGE_SECRET_KEY")
                    else:
                        env_lines.append(line if line.endswith("\n") else line + "\n")
        except Exception:
            pass

    if "GDRIVE_BRIDGE_URL" not in existing_keys:
        env_lines.append(f'GDRIVE_BRIDGE_URL="{clean_url}"\n')
    if "BRIDGE_SECRET_KEY" not in existing_keys:
        env_lines.append(f'BRIDGE_SECRET_KEY="{clean_secret}"\n')

    try:
        with open(ENV_FILE_PATH, "w", encoding="utf-8") as f:
            f.writelines(env_lines)
    except Exception:
        pass


def get_bridge_status() -> Dict[str, Any]:
    url = get_bridge_url()
    enabled = is_bridge_enabled()
    masked_url = ""
    if url:
        masked_url = url[:48] + "..." + url[-10:] if len(url) > 62 else url
    return {
        "enabled": enabled,
        "mode": "GDRIVE_CLOUD" if enabled else "LOCAL_STORAGE",
        "bridge_url": url,
        "masked_url": masked_url,
        "last_sync_at": _last_sync_info["last_sync_at"],
        "last_sync_status": _last_sync_info["last_sync_status"],
        "last_error": _last_sync_info["last_error"],
        "root_folder_url": _last_sync_info["root_folder_url"],
    }


def _call_bridge(payload: Dict[str, Any], timeout: int = 45) -> Dict[str, Any]:
    """Gửi HTTP POST tới Google Apps Script Web App Bridge (tự động theo dõi 302 Redirect của Google)."""
    url = get_bridge_url()
    if not url:
        return {"ok": False, "error": "GDRIVE_BRIDGE_URL chưa được cấu hình."}

    body = dict(payload)
    body["secretKey"] = get_bridge_secret()

    try:
        resp = requests.post(
            url,
            data=json.dumps(body),
            headers={"Content-Type": "application/json"},
            timeout=timeout,
            allow_redirects=True,
        )
        if resp.status_code != 200:
            err_msg = f"HTTP {resp.status_code} từ Google Apps Script Bridge"
            _last_sync_info["last_sync_status"] = "ERROR"
            _last_sync_info["last_error"] = err_msg
            return {"ok": False, "error": err_msg}

        data = resp.json()
        if data.get("ok"):
            _last_sync_info["last_sync_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            _last_sync_info["last_sync_status"] = "OK"
            _last_sync_info["last_error"] = ""
            if data.get("rootFolderUrl"):
                _last_sync_info["root_folder_url"] = data["rootFolderUrl"]
        else:
            _last_sync_info["last_sync_status"] = "ERROR"
            _last_sync_info["last_error"] = data.get("error", "Unknown error")
        return data
    except Exception as exc:
        err_str = str(exc)
        _last_sync_info["last_sync_status"] = "ERROR"
        _last_sync_info["last_error"] = err_str
        print(f"[GDriveBridge] Lỗi kết nối Bridge ({payload.get('action')}): {err_str}")
        return {"ok": False, "error": err_str}


def ping_bridge() -> Dict[str, Any]:
    return _call_bridge({"action": "PING"}, timeout=20)


def load_db_from_gdrive(db_path: str = DB_PATH) -> bool:
    """
    Gọi LOAD_DB khi Server vừa khởi động (startup):
    - Nếu trên Google Drive đã có file port_checklist_backup.db hợp lệ, tải về ghi đè làm DB hiện hành.
    - Trả về True nếu đã khôi phục từ Google Drive, False nếu trên Drive chưa có bản backup.
    """
    if not is_bridge_enabled():
        return False

    print("[GDriveBridge] Đang kiểm tra và tải port_checklist_backup.db từ Google Drive...")
    res = _call_bridge({"action": "LOAD_DB"}, timeout=45)
    if not res.get("ok") or not res.get("exists") or not res.get("dbBase64"):
        print("[GDriveBridge] Chưa có bản backup DB trên Google Drive, sẽ khởi tạo DB mới và đồng bộ lên Drive.")
        return False

    try:
        raw_bytes = base64.b64decode(res["dbBase64"])
        if not raw_bytes.startswith(b"SQLite format 3\x00"):
            print("[GDriveBridge] Dữ liệu tải về không phải SQLite hợp lệ, bỏ qua.")
            return False

        os.makedirs(os.path.dirname(os.path.abspath(db_path)), exist_ok=True)
        with open(db_path, "wb") as f_out:
            f_out.write(raw_bytes)
        print(f"[GDriveBridge] Đã khôi phục thành công port_checklist.db ({len(raw_bytes)} bytes) từ Google Drive!")
        return True
    except Exception as exc:
        print(f"[GDriveBridge] Lỗi giải mã DB từ Google Drive: {exc}")
        return False


def sync_db_to_gdrive(db_path: str = DB_PATH) -> Dict[str, Any]:
    """Đẩy toàn bộ file SQLite port_checklist.db lên Google Drive dưới dạng port_checklist_backup.db."""
    if not is_bridge_enabled() or not os.path.exists(db_path):
        return {"ok": False, "skipped": True}

    try:
        with open(db_path, "rb") as f_in:
            raw_bytes = f_in.read()
        b64_str = base64.b64encode(raw_bytes).decode("ascii")
        res = _call_bridge({"action": "SYNC_DB", "dbBase64": b64_str}, timeout=60)
        if res.get("ok"):
            print(f"[GDriveBridge] Đồng bộ port_checklist.db lên Google Drive thành công (FileID: {res.get('dbFileId')}).")
        return res
    except Exception as exc:
        print(f"[GDriveBridge] Lỗi đồng bộ DB lên Google Drive: {exc}")
        return {"ok": False, "error": str(exc)}


def _update_vessel_gdrive_meta(vessel_id: int, folder_id: str, folder_url: str) -> None:
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.execute(
            "UPDATE vessels SET gdrive_folder_id = ?, gdrive_folder_url = ? WHERE id = ?",
            (folder_id or "", folder_url or "", vessel_id),
        )
        conn.commit()
        conn.close()
    except Exception as exc:
        print(f"[GDriveBridge] Lỗi cập nhật metadata folder tàu: {exc}")


def _update_file_gdrive_meta(file_id: int, gdrive_file_id: str, gdrive_view_url: str) -> None:
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.execute(
            """
            UPDATE uploaded_files
            SET gdrive_file_id = ?, gdrive_view_url = ?, sync_status = 'SYNCED'
            WHERE id = ?
            """,
            (gdrive_file_id or "", gdrive_view_url or "", file_id),
        )
        conn.commit()
        conn.close()
    except Exception as exc:
        print(f"[GDriveBridge] Lỗi cập nhật metadata file Drive: {exc}")


def bg_create_vessel_folder_and_sync(vessel_id: int, vessel_folder_name: str) -> None:
    """BackgroundTask: Tạo thư mục tàu trên Google Drive và đồng bộ DB."""
    if not is_bridge_enabled():
        return
    res = _call_bridge(
        {
            "action": "CREATE_VESSEL_FOLDER",
            "vesselFolderName": vessel_folder_name,
        },
        timeout=30,
    )
    if res.get("ok"):
        _update_vessel_gdrive_meta(vessel_id, res.get("folderId", ""), res.get("folderUrl", ""))
    sync_db_to_gdrive()


def bg_upload_doc_and_sync(
    file_id: int,
    vessel_id: int,
    vessel_folder_name: str,
    local_file_path: str,
    stored_filename: str,
    mime_type: str,
) -> None:
    """BackgroundTask: Đẩy file chứng từ đã đổi tên chuẩn lên thư mục tàu trên Google Drive và đồng bộ DB."""
    if not is_bridge_enabled():
        return
    if not os.path.exists(local_file_path):
        sync_db_to_gdrive()
        return

    try:
        with open(local_file_path, "rb") as f_in:
            file_bytes = f_in.read()
        b64_data = base64.b64encode(file_bytes).decode("ascii")
        res = _call_bridge(
            {
                "action": "UPLOAD_DOC",
                "vesselFolderName": vessel_folder_name,
                "fileName": stored_filename,
                "mimeType": mime_type,
                "fileBase64": b64_data,
            },
            timeout=60,
        )
        if res.get("ok"):
            _update_file_gdrive_meta(file_id, res.get("fileId", ""), res.get("viewUrl", ""))
            if res.get("folderId"):
                _update_vessel_gdrive_meta(vessel_id, res.get("folderId", ""), res.get("folderUrl", ""))
    except Exception as exc:
        print(f"[GDriveBridge] Lỗi upload chứng từ '{stored_filename}' lên Drive: {exc}")

    sync_db_to_gdrive()


def bg_rename_doc_and_sync(
    file_id: int,
    vessel_id: int,
    gdrive_file_id: str,
    new_stored_filename: str,
    vessel_folder_name: str,
    local_file_path: str,
    mime_type: str,
) -> None:
    """BackgroundTask: Đổi tên file trên Google Drive khi phân loại 1-Click và đồng bộ DB."""
    if not is_bridge_enabled():
        return
    if gdrive_file_id:
        res = _call_bridge(
            {
                "action": "RENAME_DOC",
                "fileId": gdrive_file_id,
                "newFileName": new_stored_filename,
            },
            timeout=30,
        )
        if res.get("ok"):
            _update_file_gdrive_meta(file_id, res.get("fileId", gdrive_file_id), res.get("viewUrl", ""))
            sync_db_to_gdrive()
            return

    # Nếu file chưa có trên Drive thì upload mới
    bg_upload_doc_and_sync(
        file_id=file_id,
        vessel_id=vessel_id,
        vessel_folder_name=vessel_folder_name,
        local_file_path=local_file_path,
        stored_filename=new_stored_filename,
        mime_type=mime_type,
    )


def bg_delete_doc_and_sync(gdrive_file_id: str, vessel_folder_name: str, stored_filename: str) -> None:
    """BackgroundTask: Xóa file chứng từ trên Google Drive và đồng bộ DB."""
    if not is_bridge_enabled():
        return
    _call_bridge(
        {
            "action": "DELETE_DOC",
            "fileId": gdrive_file_id,
            "vesselFolderName": vessel_folder_name,
            "fileName": stored_filename,
        },
        timeout=30,
    )
    sync_db_to_gdrive()


def bg_clear_vessel_folder_and_sync(gdrive_folder_id: str, vessel_folder_name: str) -> None:
    """BackgroundTask: Làm trống toàn bộ chứng từ trong thư mục tàu trên Google Drive và đồng bộ DB."""
    if not is_bridge_enabled():
        return
    _call_bridge(
        {
            "action": "CLEAR_VESSEL_FOLDER",
            "folderId": gdrive_folder_id,
            "vesselFolderName": vessel_folder_name,
        },
        timeout=30,
    )
    sync_db_to_gdrive()


def bg_delete_vessel_folder_and_sync(gdrive_folder_id: str, vessel_folder_name: str) -> None:
    """BackgroundTask: Xóa thư mục tàu trên Google Drive và đồng bộ DB."""
    if not is_bridge_enabled():
        return
    _call_bridge(
        {
            "action": "DELETE_VESSEL_FOLDER",
            "folderId": gdrive_folder_id,
            "vesselFolderName": vessel_folder_name,
        },
        timeout=30,
    )
    sync_db_to_gdrive()


def download_doc_from_gdrive_if_missing(gdrive_file_id: str, dest_abs_path: str) -> bool:
    """
    Nếu container Render vừa khởi động lại làm mất file vật lý trong ./storage/...,
    tự động kéo file từ Google Drive về theo gdrive_file_id để phục vụ tải/xem file.
    """
    if not is_bridge_enabled() or not gdrive_file_id:
        return False
    res = _call_bridge({"action": "DOWNLOAD_DOC", "fileId": gdrive_file_id}, timeout=45)
    if not res.get("ok") or not res.get("fileBase64"):
        return False
    try:
        os.makedirs(os.path.dirname(os.path.abspath(dest_abs_path)), exist_ok=True)
        with open(dest_abs_path, "wb") as f_out:
            f_out.write(base64.b64decode(res["fileBase64"]))
        return True
    except Exception as exc:
        print(f"[GDriveBridge] Lỗi tải file từ Drive: {exc}")
        return False
