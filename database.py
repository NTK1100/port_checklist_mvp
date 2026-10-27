import os
import re
import json
import sqlite3
from datetime import datetime
from typing import List, Dict, Any, Optional

from auth import hash_password
from parser_engine import build_standardized_filename, sanitize_token

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "port_checklist.db")
STORAGE_ROOT = os.path.join(BASE_DIR, "storage")

DOC_TYPE_INFO = {
    "MANIFEST": {
        "code": "MANIFEST",
        "name": "Cargo Manifest",
        "vn_name": "Bản lược khai hàng hóa",
    },
    "BL": {
        "code": "BL",
        "name": "Bill of Lading",
        "vn_name": "Vận đơn đường biển",
    },
    "VGM": {
        "code": "VGM",
        "name": "Verified Gross Mass",
        "vn_name": "Phiếu xác nhận khối lượng VGM",
    },
    "CREW_LIST": {
        "code": "CREW_LIST",
        "name": "Crew List (FAL Form 5)",
        "vn_name": "Danh sách thuyền viên",
    },
    "PORT_CLEARANCE": {
        "code": "PORT_CLEARANCE",
        "name": "Port Clearance",
        "vn_name": "Giấy phép rời cảng cuối",
    },
    "STOWAGE_PLAN": {
        "code": "STOWAGE_PLAN",
        "name": "Stowage Plan",
        "vn_name": "Sơ đồ xếp hàng",
    },
    "MATES_RECEIPT": {
        "code": "MATES_RECEIPT",
        "name": "Mate's Receipt",
        "vn_name": "Biên lai thuyền phó",
    },
    "UNCLASSIFIED": {
        "code": "UNCLASSIFIED",
        "name": "Unclassified Document",
        "vn_name": "Chứng từ chưa phân loại",
    },
}

VESSEL_TYPES = {
    "CONTAINER_INTL": {
        "code": "CONTAINER_INTL",
        "label": "Tàu Container Quốc tế",
        "checklist": [
            {"doc_type": "MANIFEST", "doc_name": "Cargo Manifest", "requires_count_match": 0},
            {"doc_type": "BL", "doc_name": "Bill of Lading", "requires_count_match": 1},
            {"doc_type": "VGM", "doc_name": "Verified Gross Mass", "requires_count_match": 1},
            {"doc_type": "CREW_LIST", "doc_name": "Danh sách thuyền viên", "requires_count_match": 0},
            {"doc_type": "PORT_CLEARANCE", "doc_name": "Giấy phép rời cảng cuối", "requires_count_match": 0},
        ],
    },
    "BULK": {
        "code": "BULK",
        "label": "Tàu Hàng rời (Bulk)",
        "checklist": [
            {"doc_type": "MANIFEST", "doc_name": "Cargo Manifest", "requires_count_match": 0},
            {"doc_type": "STOWAGE_PLAN", "doc_name": "Sơ đồ xếp hàng", "requires_count_match": 0},
            {"doc_type": "MATES_RECEIPT", "doc_name": "Biên lai thuyền phó", "requires_count_match": 0},
            {"doc_type": "CREW_LIST", "doc_name": "Danh sách thuyền viên", "requires_count_match": 0},
            {"doc_type": "PORT_CLEARANCE", "doc_name": "Giấy phép rời cảng cuối", "requires_count_match": 0},
        ],
    },
    "DOMESTIC": {
        "code": "DOMESTIC",
        "label": "Tàu Nội địa",
        "checklist": [
            {"doc_type": "MANIFEST", "doc_name": "Bản lược khai hàng hóa", "requires_count_match": 0},
            {"doc_type": "CREW_LIST", "doc_name": "Danh sách thuyền viên", "requires_count_match": 0},
            {"doc_type": "PORT_CLEARANCE", "doc_name": "Giấy phép rời cảng cuối", "requires_count_match": 0},
        ],
    },
}

DEFAULT_USERS = [
    {
        "username": "admin",
        "full_name": "Quản trị viên Hệ thống",
        "password": "Admin@123",
        "role": "ADMIN",
    },
    {
        "username": "ops_kien",
        "full_name": "Nguyễn Trung Kiên (Chuyên viên Chứng từ)",
        "password": "Ops@123",
        "role": "OPS",
    },
    {
        "username": "viewer",
        "full_name": "Ban Giám đốc / Cảng vụ",
        "password": "View@123",
        "role": "VIEWER",
    },
]


def normalize_vessel_type(vessel_type_input: str) -> str:
    if not vessel_type_input:
        return "CONTAINER_INTL"
    val = vessel_type_input.strip()
    if val in VESSEL_TYPES:
        return val
    val_upper = val.upper()
    if "CONTAINER" in val_upper or "QUỐC TẾ" in val_upper or "QUOC TE" in val_upper:
        return "CONTAINER_INTL"
    if "BULK" in val_upper or "HÀNG RỜI" in val_upper or "HANG ROI" in val_upper:
        return "BULK"
    if "DOMESTIC" in val_upper or "NỘI ĐỊA" in val_upper or "NOI DIA" in val_upper:
        return "DOMESTIC"
    return "CONTAINER_INTL"


def sanitize_token(text: str) -> str:
    if not text:
        return "UNKNOWN"
    cleaned = re.sub(r"[^\w\-]", "_", text.strip())
    cleaned = re.sub(r"_+", "_", cleaned).strip("_")
    return cleaned.upper() if cleaned else "UNKNOWN"


def extract_date_str(eta_str: str) -> str:
    if eta_str:
        m = re.search(r"(\d{4}-\d{2}-\d{2})", eta_str)
        if m:
            return m.group(1)
        m2 = re.search(r"(\d{2})/(\d{2})/(\d{4})", eta_str)
        if m2:
            return f"{m2.group(3)}-{m2.group(2)}-{m2.group(1)}"
    return datetime.now().strftime("%Y-%m-%d")


def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def init_db() -> None:
    """Khởi tạo các bảng SQLite (vessels, checklist_items, uploaded_files, users, audit_logs) và seed 3 user mặc định."""
    os.makedirs(STORAGE_ROOT, exist_ok=True)
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            full_name TEXT NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL,
            is_active INTEGER DEFAULT 1,
            created_at TEXT NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS audit_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT NOT NULL,
            username TEXT NOT NULL,
            role TEXT NOT NULL,
            action TEXT NOT NULL,
            details TEXT NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS vessels (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            vessel_name TEXT NOT NULL,
            voyage_no TEXT NOT NULL,
            eta TEXT NOT NULL,
            eta_date TEXT NOT NULL,
            vessel_type TEXT NOT NULL,
            vessel_type_label TEXT NOT NULL,
            storage_folder TEXT NOT NULL,
            expected_bl_count INTEGER DEFAULT 0,
            status_color TEXT DEFAULT 'RED',
            status_summary TEXT DEFAULT 'Chưa có chứng từ',
            created_at TEXT NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS checklist_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            vessel_id INTEGER NOT NULL,
            doc_type TEXT NOT NULL,
            doc_name TEXT NOT NULL,
            is_required INTEGER DEFAULT 1,
            requires_count_match INTEGER DEFAULT 0,
            expected_count INTEGER DEFAULT 0,
            actual_count INTEGER DEFAULT 0,
            file_count INTEGER DEFAULT 0,
            status TEXT DEFAULT 'MISSING',
            note TEXT DEFAULT 'Chưa nộp',
            FOREIGN KEY (vessel_id) REFERENCES vessels(id) ON DELETE CASCADE
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS uploaded_files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            vessel_id INTEGER NOT NULL,
            original_filename TEXT NOT NULL,
            stored_filename TEXT NOT NULL,
            file_path TEXT NOT NULL,
            file_ext TEXT NOT NULL,
            file_size INTEGER DEFAULT 0,
            doc_type TEXT NOT NULL,
            detected_method TEXT DEFAULT 'CONTENT_REGEX',
            detected_keyword TEXT DEFAULT '',
            detected_vessel_name TEXT DEFAULT '',
            is_wrong_vessel INTEGER DEFAULT 0,
            error_code TEXT DEFAULT '',
            error_message TEXT DEFAULT '',
            extracted_count INTEGER DEFAULT 0,
            extracted_codes_json TEXT DEFAULT '[]',
            extracted_text_preview TEXT DEFAULT '',
            uploaded_by TEXT DEFAULT 'ops_kien',
            uploaded_at TEXT NOT NULL,
            FOREIGN KEY (vessel_id) REFERENCES vessels(id) ON DELETE CASCADE
        )
    """)

    # Đảm bảo các cột mở rộng tồn tại nếu nâng cấp từ DB cũ
    cursor.execute("PRAGMA table_info(vessels)")
    vessel_cols = {row["name"] for row in cursor.fetchall()}
    if "gdrive_folder_id" not in vessel_cols:
        cursor.execute("ALTER TABLE vessels ADD COLUMN gdrive_folder_id TEXT DEFAULT ''")
    if "gdrive_folder_url" not in vessel_cols:
        cursor.execute("ALTER TABLE vessels ADD COLUMN gdrive_folder_url TEXT DEFAULT ''")

    cursor.execute("PRAGMA table_info(checklist_items)")
    chk_cols = {row["name"] for row in cursor.fetchall()}
    if "is_manual_verified" not in chk_cols:
        cursor.execute("ALTER TABLE checklist_items ADD COLUMN is_manual_verified INTEGER DEFAULT 0")

    cursor.execute("PRAGMA table_info(uploaded_files)")
    existing_cols = {row["name"] for row in cursor.fetchall()}
    if "uploaded_by" not in existing_cols:
        cursor.execute("ALTER TABLE uploaded_files ADD COLUMN uploaded_by TEXT DEFAULT 'ops_kien'")
    if "gdrive_file_id" not in existing_cols:
        cursor.execute("ALTER TABLE uploaded_files ADD COLUMN gdrive_file_id TEXT DEFAULT ''")
    if "gdrive_view_url" not in existing_cols:
        cursor.execute("ALTER TABLE uploaded_files ADD COLUMN gdrive_view_url TEXT DEFAULT ''")
    if "sync_status" not in existing_cols:
        cursor.execute("ALTER TABLE uploaded_files ADD COLUMN sync_status TEXT DEFAULT 'LOCAL'")

    # Seed 3 tài khoản mặc định
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for u in DEFAULT_USERS:
        cursor.execute("SELECT id FROM users WHERE username = ?", (u["username"],))
        if not cursor.fetchone():
            cursor.execute(
                """
                INSERT INTO users (username, full_name, password_hash, role, is_active, created_at)
                VALUES (?, ?, ?, ?, 1, ?)
                """,
                (u["username"], u["full_name"], hash_password(u["password"]), u["role"], now_str),
            )

    conn.commit()
    conn.close()


# ==================== USER & AUDIT LOG HELPERS ====================

def log_audit_event(username: str, role: str, action: str, details: str) -> None:
    """Ghi nhật ký bảo mật (Audit Log) cho các hành động LOGIN, CREATE_VESSEL, UPLOAD_FILE, DELETE_FILE, USER_MGMT."""
    conn = get_connection()
    cursor = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute(
        """
        INSERT INTO audit_logs (timestamp, username, role, action, details)
        VALUES (?, ?, ?, ?, ?)
        """,
        (now_str, username, role, action, details),
    )
    conn.commit()
    conn.close()


def get_audit_logs(limit: int = 200) -> List[Dict[str, Any]]:
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM audit_logs ORDER BY id DESC LIMIT ?", (limit,))
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows


def get_user_by_username(username: str) -> Optional[Dict[str, Any]]:
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE LOWER(username) = LOWER(?)", (username.strip(),))
    row = cursor.fetchone()
    conn.close()
    return dict(row) if row else None


def get_user_by_id(user_id: int) -> Optional[Dict[str, Any]]:
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE id = ?", (user_id,))
    row = cursor.fetchone()
    conn.close()
    return dict(row) if row else None


def list_all_users() -> List[Dict[str, Any]]:
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, username, full_name, role, is_active, created_at FROM users ORDER BY id ASC")
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows


# ==================== INTERNAL STORAGE PATH HELPERS (NO PATH DISCLOSURE TO CLIENT) ====================

def get_vessel_storage_abs_path(vessel_id: int) -> Optional[str]:
    """Chỉ dùng nội bộ trên Backend để lấy thư mục vật lý của tàu, tuyệt đối không trả về qua API."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT storage_folder FROM vessels WHERE id = ?", (vessel_id,))
    row = cursor.fetchone()
    conn.close()
    if not row:
        return None
    rel = row["storage_folder"].replace("./", "").strip("/")
    abs_p = os.path.abspath(os.path.join(BASE_DIR, rel))
    if not abs_p.startswith(os.path.abspath(STORAGE_ROOT)):
        return None
    os.makedirs(abs_p, exist_ok=True)
    return abs_p


def get_vessel_folder_meta_internal(vessel_id: int) -> Dict[str, str]:
    """Lấy tên thư mục chuẩn (VD: 2026-10-26_MV_BINH_HOA) và gdrive_folder_id phục vụ đồng bộ Google Drive."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT storage_folder, eta_date, vessel_name, gdrive_folder_id, gdrive_folder_url FROM vessels WHERE id = ?",
        (vessel_id,),
    )
    row = cursor.fetchone()
    conn.close()
    if not row:
        return {"folder_name": "UNKNOWN_VESSEL", "gdrive_folder_id": "", "gdrive_folder_url": ""}
    rel = (row["storage_folder"] or "").replace("./storage/", "").strip("/")
    folder_name = rel or f"{row['eta_date']}_{sanitize_token(row['vessel_name'])}"
    return {
        "folder_name": folder_name,
        "gdrive_folder_id": row["gdrive_folder_id"] or "",
        "gdrive_folder_url": row["gdrive_folder_url"] or "",
    }


def get_file_record_internal(file_id: int) -> Optional[Dict[str, Any]]:
    """Lấy bản ghi file đầy đủ kèm đường dẫn vật lý phục vụ xử lý nội bộ trên server."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM uploaded_files WHERE id = ?", (file_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        return None
    f_dict = dict(row)
    cursor.execute(
        "SELECT storage_folder, vessel_name, voyage_no, eta_date, gdrive_folder_id FROM vessels WHERE id = ?",
        (f_dict["vessel_id"],),
    )
    v_row = cursor.fetchone()
    conn.close()
    if v_row:
        rel_folder = v_row["storage_folder"].replace("./", "").strip("/")
        f_dict["abs_vessel_folder"] = os.path.abspath(os.path.join(BASE_DIR, rel_folder))
        f_dict["abs_file_path"] = os.path.abspath(
            os.path.join(BASE_DIR, rel_folder, os.path.basename(f_dict["stored_filename"]))
        )
        f_dict["vessel_name"] = v_row["vessel_name"]
        f_dict["voyage_no"] = v_row["voyage_no"]
        f_dict["vessel_folder_name"] = v_row["storage_folder"].replace("./storage/", "").strip("/") or (
            f"{v_row['eta_date']}_{sanitize_token(v_row['vessel_name'])}"
        )
    return f_dict


# ==================== VESSEL & CHECKLIST OPERATIONS ====================

def create_vessel(
    vessel_name: str,
    eta: str,
    vessel_type: str,
    voyage_no: str = "",
    eta_date_override: Optional[str] = None,
) -> Dict[str, Any]:
    """Tạo tàu mới, khởi tạo thư mục lưu trữ vật lý và danh sách Checklist động."""
    v_type_code = normalize_vessel_type(vessel_type)
    v_type_info = VESSEL_TYPES[v_type_code]
    v_name_clean = vessel_name.strip().upper()
    voyage_clean = (voyage_no or "").strip().upper()
    eta_date = eta_date_override or extract_date_str(eta)

    folder_name = f"{eta_date}_{sanitize_token(v_name_clean)}"
    rel_storage_folder = f"./storage/{folder_name}/"
    abs_storage_folder = os.path.join(STORAGE_ROOT, folder_name)
    os.makedirs(abs_storage_folder, exist_ok=True)

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        """
        INSERT INTO vessels (
            vessel_name, voyage_no, eta, eta_date,
            vessel_type, vessel_type_label, storage_folder,
            expected_bl_count, status_color, status_summary, created_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, 0, 'RED', 'Chưa có chứng từ', ?)
        """,
        (
            v_name_clean,
            voyage_clean,
            eta.strip(),
            eta_date,
            v_type_code,
            v_type_info["label"],
            rel_storage_folder,
            now_str,
        ),
    )
    vessel_id = cursor.lastrowid

    for item in v_type_info["checklist"]:
        cursor.execute(
            """
            INSERT INTO checklist_items (
                vessel_id, doc_type, doc_name, is_required,
                requires_count_match, expected_count, actual_count,
                file_count, status, note
            ) VALUES (?, ?, ?, 1, ?, 0, 0, 0, 'MISSING', 'Chưa nộp')
            """,
            (vessel_id, item["doc_type"], item["doc_name"], item["requires_count_match"]),
        )

    conn.commit()
    conn.close()
    return get_vessel_detail(vessel_id)


def update_vessel_info(
    vessel_id: int,
    vessel_name: str,
    eta: str,
    vessel_type: Optional[str] = None,
) -> Dict[str, Any]:
    """Cập nhật Tên tàu, Giờ ETA và Loại tàu (kèm đổi tên thư mục lưu trữ nếu thay đổi)."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM vessels WHERE id = ?", (vessel_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        return {}

    old_vessel = dict(row)
    v_name_clean = (vessel_name or old_vessel["vessel_name"]).strip().upper()
    eta_clean = (eta or old_vessel["eta"]).strip().replace("T", " ")
    eta_date = extract_date_str(eta_clean)
    v_type_code = normalize_vessel_type(vessel_type) if vessel_type else old_vessel["vessel_type"]
    v_type_info = VESSEL_TYPES[v_type_code]

    old_folder_meta = get_vessel_folder_meta_internal(vessel_id)
    old_abs_folder = get_vessel_storage_abs_path(vessel_id)

    new_folder_name = f"{eta_date}_{sanitize_token(v_name_clean)}"
    new_rel_storage_folder = f"./storage/{new_folder_name}/"
    new_abs_folder = os.path.join(STORAGE_ROOT, new_folder_name)

    if old_abs_folder and os.path.exists(old_abs_folder) and os.path.abspath(old_abs_folder) != os.path.abspath(new_abs_folder):
        try:
            if not os.path.exists(new_abs_folder):
                os.rename(old_abs_folder, new_abs_folder)
            else:
                os.makedirs(new_abs_folder, exist_ok=True)
        except Exception:
            os.makedirs(new_abs_folder, exist_ok=True)
    else:
        os.makedirs(new_abs_folder, exist_ok=True)

    cursor.execute(
        """
        UPDATE vessels
        SET vessel_name = ?, eta = ?, eta_date = ?, vessel_type = ?, vessel_type_label = ?, storage_folder = ?
        WHERE id = ?
        """,
        (v_name_clean, eta_clean, eta_date, v_type_code, v_type_info["label"], new_rel_storage_folder, vessel_id),
    )

    # Nếu thay đổi loại tàu thì đồng bộ lại danh sách checklist_items theo loại tàu mới
    if v_type_code != old_vessel["vessel_type"]:
        cursor.execute("SELECT doc_type, is_manual_verified FROM checklist_items WHERE vessel_id = ?", (vessel_id,))
        existing_verified = {r["doc_type"]: r["is_manual_verified"] for r in cursor.fetchall()}
        cursor.execute("DELETE FROM checklist_items WHERE vessel_id = ?", (vessel_id,))
        for item in v_type_info["checklist"]:
            mv = existing_verified.get(item["doc_type"], 0)
            cursor.execute(
                """
                INSERT INTO checklist_items (
                    vessel_id, doc_type, doc_name, is_required,
                    requires_count_match, expected_count, actual_count,
                    file_count, status, note, is_manual_verified
                ) VALUES (?, ?, ?, 1, ?, 0, 0, 0, 'MISSING', 'Chưa nộp', ?)
                """,
                (vessel_id, item["doc_type"], item["doc_name"], item["requires_count_match"], mv),
            )

    conn.commit()
    conn.close()

    detail = recalculate_vessel_status(vessel_id)
    detail["_old_folder_name"] = old_folder_meta["folder_name"]
    detail["_new_folder_name"] = new_folder_name
    return detail


def get_next_sequence_number(
    vessel_id: int,
    doc_type: str,
    file_ext: str = "",
    conn: Optional[sqlite3.Connection] = None,
) -> int:
    """
    Tính số thứ tự cho file mới theo (vessel_id, doc_type, file_ext).
    Nếu 2 file cùng loại chứng từ nhưng KHÁC định dạng (VD: 1 file .docx và 1 file .pdf cho CREW_LIST),
    mỗi định dạng sẽ có seq_num = 1 -> [CREW_LIST]_TEN_TAU.docx và [CREW_LIST]_TEN_TAU.pdf (không gắn _(2)).
    """
    should_close = False
    if conn is None:
        conn = get_connection()
        should_close = True
    cursor = conn.cursor()
    if file_ext:
        ext_norm = file_ext.lower() if file_ext.startswith(".") else f".{file_ext.lower()}"
        cursor.execute(
            """
            SELECT COUNT(*) AS cnt FROM uploaded_files
            WHERE vessel_id = ? AND doc_type = ? AND LOWER(file_ext) = ? AND is_wrong_vessel = 0
            """,
            (vessel_id, doc_type, ext_norm),
        )
    else:
        cursor.execute(
            "SELECT COUNT(*) AS cnt FROM uploaded_files WHERE vessel_id = ? AND doc_type = ? AND is_wrong_vessel = 0",
            (vessel_id, doc_type),
        )
    row = cursor.fetchone()
    count = row["cnt"] if row else 0
    if should_close:
        conn.close()
    return count + 1


def recalculate_vessel_status(vessel_id: int) -> Dict[str, Any]:
    """
    Tính toán lại trạng thái Checklist, đối chiếu số lượng B/L & VGM với Manifest,
    hỗ trợ cờ xác nhận thủ công (is_manual_verified) và cập nhật trạng thái màu của tàu (RED / ORANGE / GREEN).
    """
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM vessels WHERE id = ?", (vessel_id,))
    vessel = cursor.fetchone()
    if not vessel:
        conn.close()
        return {}

    cursor.execute("SELECT * FROM uploaded_files WHERE vessel_id = ? ORDER BY id ASC", (vessel_id,))
    files = [dict(r) for r in cursor.fetchall()]

    cursor.execute("SELECT * FROM checklist_items WHERE vessel_id = ? ORDER BY id ASC", (vessel_id,))
    checklist_items = [dict(r) for r in cursor.fetchall()]

    total_files = len(files)
    wrong_vessel_files = [f for f in files if f["is_wrong_vessel"] == 1]
    unclassified_files = [f for f in files if f["doc_type"] == "UNCLASSIFIED" and f["is_wrong_vessel"] == 0]
    valid_files = [f for f in files if f["is_wrong_vessel"] == 0 and f["doc_type"] != "UNCLASSIFIED"]

    # 1. Tính expected_count từ các file MANIFEST hợp lệ
    manifest_files = [f for f in valid_files if f["doc_type"] == "MANIFEST"]
    expected_bl_count = 0
    manifest_codes = set()
    for mf in manifest_files:
        try:
            codes = json.loads(mf["extracted_codes_json"] or "[]")
            for c in codes:
                manifest_codes.add(c)
        except Exception:
            pass
        if mf["extracted_count"] > expected_bl_count:
            expected_bl_count = mf["extracted_count"]

    if len(manifest_codes) > expected_bl_count:
        expected_bl_count = len(manifest_codes)

    # 2. Cập nhật từng hạng mục trong Checklist
    all_required_completed = True
    missing_count = 0
    mismatch_notes = []

    for item in checklist_items:
        d_type = item["doc_type"]
        is_manual = int(item.get("is_manual_verified") or 0)
        matching_valid_files = [f for f in valid_files if f["doc_type"] == d_type]
        f_count = len(matching_valid_files)

        unique_codes = set()
        sum_extracted_count = 0
        for mf in matching_valid_files:
            sum_extracted_count += mf["extracted_count"] or 0
            try:
                codes = json.loads(mf["extracted_codes_json"] or "[]")
                for c in codes:
                    unique_codes.add(c)
            except Exception:
                pass

        actual_count = len(unique_codes) if len(unique_codes) > 0 else sum_extracted_count

        if is_manual == 1:
            item_expected = expected_bl_count if (d_type == "MANIFEST" or item["requires_count_match"] == 1) else 1
            status = "COMPLETED"
            if d_type == "MANIFEST":
                note = f"✓ Đạt ({expected_bl_count} B/L)" if expected_bl_count > 0 else "✓ Đạt"
            elif item["requires_count_match"] == 1 and expected_bl_count > 0:
                disp_cnt = actual_count if actual_count > 0 else expected_bl_count
                note = f"✓ Đạt ({disp_cnt}/{expected_bl_count} B/L)"
            else:
                note = "✓ Đạt"
        elif d_type == "MANIFEST":
            item_expected = expected_bl_count
            if f_count == 0:
                status = "MISSING"
                note = "✕ Chưa nộp"
                all_required_completed = False
                missing_count += 1
            else:
                status = "COMPLETED"
                note = f"✓ Đạt ({expected_bl_count} B/L)" if expected_bl_count > 0 else "✓ Đạt"
        elif item["requires_count_match"] == 1:
            item_expected = expected_bl_count
            if f_count == 0:
                status = "MISSING"
                note = "✕ Chưa nộp"
                all_required_completed = False
                missing_count += 1
            else:
                if item_expected > 0:
                    if actual_count < item_expected:
                        status = "INCOMPLETE"
                        note = f"— Thiếu ({actual_count}/{item_expected} B/L)"
                        all_required_completed = False
                        mismatch_notes.append(f"{d_type} ({actual_count}/{item_expected} B/L)")
                    else:
                        status = "COMPLETED"
                        note = f"✓ Đạt ({actual_count}/{item_expected} B/L)"
                else:
                    if len(manifest_files) == 0:
                        status = "INCOMPLETE"
                        note = "— Chờ Manifest"
                        all_required_completed = False
                        mismatch_notes.append(f"{d_type} (Chờ Manifest)")
                    else:
                        status = "COMPLETED"
                        note = f"✓ Đạt ({actual_count} B/L)" if actual_count > 0 else "✓ Đạt"
        else:
            item_expected = 1
            if f_count == 0:
                status = "MISSING"
                note = "✕ Chưa nộp"
                all_required_completed = False
                missing_count += 1
            else:
                status = "COMPLETED"
                note = "✓ Đạt"

        cursor.execute(
            """
            UPDATE checklist_items
            SET expected_count = ?, actual_count = ?, file_count = ?, status = ?, note = ?
            WHERE id = ?
            """,
            (item_expected, actual_count, f_count, status, note, item["id"]),
        )

    # 3. Tính màu trạng thái tổng thể của tàu (RED / ORANGE / GREEN)
    if total_files == 0 and all(int(it.get("is_manual_verified") or 0) == 0 for it in checklist_items):
        status_color = "RED"
        status_summary = "Cần xử lý · Chưa có chứng từ"
    elif len(wrong_vessel_files) > 0:
        status_color = "RED"
        status_summary = f"Cần xử lý · {len(wrong_vessel_files)} chứng từ sai tên tàu"
    elif not all_required_completed or len(unclassified_files) > 0:
        status_color = "ORANGE"
        parts = []
        if missing_count > 0:
            parts.append(f"Thiếu {missing_count} hạng mục")
        if len(mismatch_notes) > 0:
            parts.append(", ".join(mismatch_notes))
        if len(unclassified_files) > 0:
            parts.append(f"{len(unclassified_files)} file chờ phân loại")
        status_summary = "Thiếu chứng từ · " + " · ".join(parts) if parts else "Thiếu chứng từ"
    else:
        status_color = "GREEN"
        status_summary = "Đủ hồ sơ"

    cursor.execute(
        """
        UPDATE vessels
        SET expected_bl_count = ?, status_color = ?, status_summary = ?
        WHERE id = ?
        """,
        (expected_bl_count, status_color, status_summary, vessel_id),
    )

    conn.commit()
    conn.close()
    return get_vessel_detail(vessel_id)


def _sanitize_file_for_response(r: sqlite3.Row) -> Dict[str, Any]:
    """Chuyển đổi bản ghi file sang JSON-safe dict, loại bỏ tuyệt đối file_path vật lý (No Path Disclosure)."""
    f_dict = dict(r)
    f_dict.pop("file_path", None)
    try:
        f_dict["extracted_codes"] = json.loads(f_dict.pop("extracted_codes_json", "[]") or "[]")
    except Exception:
        f_dict["extracted_codes"] = []
    f_dict["doc_type_name"] = DOC_TYPE_INFO.get(f_dict["doc_type"], {}).get("name", f_dict["doc_type"])
    f_dict["download_url"] = f"/api/files/{f_dict['id']}/download"
    return f_dict


def _build_standardized_match_info(item: Dict[str, Any], wrong_vessel_files: List[Dict[str, Any]]) -> Dict[str, str]:
    """
    Quy toàn bộ cột ĐỐI CHIẾU về đúng 3 trạng thái chuẩn:
    1. GREEN (#10B981) + '✓': '✓ Đạt', '✓ Đạt (5/5 B/L)', '✓ Đạt (5 B/L)'
    2. ORANGE (#F59E0B) + '—': '— Thiếu (3/5 B/L)', '— Chờ Manifest'
    3. RED (#EF4444) + '✕': '✕ Chưa nộp', '✕ Sai tàu'
    """
    d_type = item.get("doc_type", "")
    status = item.get("status", "MISSING")
    req_match = int(item.get("requires_count_match") or 0)
    exp = int(item.get("expected_count") or 0)
    act = int(item.get("actual_count") or 0)

    has_wrong_vessel_for_type = any(wf.get("doc_type") == d_type for wf in wrong_vessel_files)

    if status == "COMPLETED":
        if d_type == "MANIFEST":
            label = f"✓ Đạt ({exp} B/L)" if exp > 0 else "✓ Đạt"
        elif req_match == 1:
            if exp > 0:
                shown_act = act if act >= exp else (act if act > 0 else exp)
                label = f"✓ Đạt ({shown_act}/{exp} B/L)"
            elif act > 0:
                label = f"✓ Đạt ({act} B/L)"
            else:
                label = "✓ Đạt"
        else:
            label = "✓ Đạt"
        return {"match_state": "GREEN", "match_label": label}

    if status == "INCOMPLETE":
        if req_match == 1:
            if exp <= 0:
                label = "— Chờ Manifest"
            else:
                label = f"— Thiếu ({act}/{exp} B/L)"
        else:
            label = "— Thiếu"
        return {"match_state": "ORANGE", "match_label": label}

    # status == 'MISSING'
    if has_wrong_vessel_for_type:
        return {"match_state": "RED", "match_label": "✕ Sai tàu"}
    return {"match_state": "RED", "match_label": "✕ Chưa nộp"}


def get_vessel_detail(vessel_id: int) -> Dict[str, Any]:
    """
    Lấy thông tin chi tiết của 1 tàu kèm Checklist và danh sách file.
    Tuyệt đối KHÔNG trả về storage_folder hoặc file_path vật lý (No Path Disclosure).
    """
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM vessels WHERE id = ?", (vessel_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        return {}

    vessel = dict(row)
    vessel.pop("storage_folder", None)

    cursor.execute("SELECT * FROM checklist_items WHERE vessel_id = ? ORDER BY id ASC", (vessel_id,))
    checklist = [dict(r) for r in cursor.fetchall()]

    cursor.execute("SELECT * FROM uploaded_files WHERE vessel_id = ? ORDER BY id ASC", (vessel_id,))
    files = [_sanitize_file_for_response(r) for r in cursor.fetchall()]

    conn.close()

    wrong_vessel_files = [f for f in files if f["is_wrong_vessel"] == 1]
    unclassified_files = [f for f in files if f["doc_type"] == "UNCLASSIFIED" and f["is_wrong_vessel"] == 0]

    for item in checklist:
        item["files"] = [
            f for f in files if f["doc_type"] == item["doc_type"] and f["is_wrong_vessel"] == 0
        ]
        m_info = _build_standardized_match_info(item, wrong_vessel_files)
        item["match_state"] = m_info["match_state"]
        item["match_label"] = m_info["match_label"]

    completed_items = sum(1 for item in checklist if item["status"] == "COMPLETED")
    total_items = len(checklist)

    vessel["checklist"] = checklist
    vessel["files"] = files
    vessel["wrong_vessel_files"] = wrong_vessel_files
    vessel["unclassified_files"] = unclassified_files
    vessel["completed_checklist_count"] = completed_items
    vessel["total_checklist_count"] = total_items
    vessel["progress_percent"] = int(round((completed_items / total_items) * 100)) if total_items > 0 else 0
    vessel["wrong_vessel_count"] = len(wrong_vessel_files)
    vessel["unclassified_count"] = len(unclassified_files)

    return vessel


def get_all_vessels() -> List[Dict[str, Any]]:
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM vessels ORDER BY id ASC")
    ids = [r["id"] for r in cursor.fetchall()]
    conn.close()
    return [get_vessel_detail(vid) for vid in ids]
