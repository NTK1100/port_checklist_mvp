import os
import io
import json
import shutil
from datetime import datetime
from typing import List, Optional, Dict, Any

import fitz  # PyMuPDF
import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fastapi import FastAPI, UploadFile, File, HTTPException, Depends, Response, BackgroundTasks, status
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from auth import (
    COOKIE_NAME,
    TOKEN_TTL_SECONDS,
    hash_password,
    verify_password,
    create_session_token,
    validate_file_security,
    sanitize_filename,
    get_current_user,
    require_roles,
)
from database import (
    init_db,
    get_connection,
    create_vessel,
    get_vessel_detail,
    get_all_vessels,
    recalculate_vessel_status,
    get_next_sequence_number,
    get_vessel_storage_abs_path,
    get_vessel_folder_meta_internal,
    get_file_record_internal,
    get_user_by_username,
    get_user_by_id,
    list_all_users,
    log_audit_event,
    get_audit_logs,
    VESSEL_TYPES,
    DOC_TYPE_INFO,
    STORAGE_ROOT,
    DB_PATH,
)
from parser_engine import parse_and_validate_file, build_standardized_filename, sanitize_token
from generate_samples import seed_initial_vessels, generate_all_sample_files, SAMPLE_DIR, create_pdf_file
from gdrive_bridge_client import (
    is_bridge_enabled,
    get_bridge_status,
    set_bridge_config,
    ping_bridge,
    load_db_from_gdrive,
    sync_db_to_gdrive,
    bg_create_vessel_folder_and_sync,
    bg_upload_doc_and_sync,
    bg_rename_doc_and_sync,
    bg_delete_doc_and_sync,
    bg_clear_vessel_folder_and_sync,
    bg_delete_vessel_folder_and_sync,
    download_doc_from_gdrive_if_missing,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")

os.makedirs(STATIC_DIR, exist_ok=True)
os.makedirs(STORAGE_ROOT, exist_ok=True)
os.makedirs(SAMPLE_DIR, exist_ok=True)

app = FastAPI(
    title="PortCall Docs",
    version="2.2.0",
)


# ==================== PYDANTIC SCHEMAS ====================

class LoginRequest(BaseModel):
    username: str
    password: str


class VesselCreateItem(BaseModel):
    vessel_name: str
    eta: str
    vessel_type: str
    voyage_no: Optional[str] = ""


class VesselBatchCreateRequest(BaseModel):
    vessels: List[VesselCreateItem]


class ClassifyFileRequest(BaseModel):
    doc_type: str


class SimulateSampleRequest(BaseModel):
    filenames: List[str]


class UserCreateRequest(BaseModel):
    username: str
    full_name: str
    password: str
    role: str = "OPS"


class UserUpdateRequest(BaseModel):
    role: Optional[str] = None
    is_active: Optional[bool] = None
    full_name: Optional[str] = None
    password: Optional[str] = None


class GDriveBridgeConfigRequest(BaseModel):
    bridge_url: str
    secret_key: Optional[str] = "PortCall_Secret_2026"


def _mime_for_ext(ext: str) -> str:
    ext = (ext or "").lower()
    if ext == ".pdf":
        return "application/pdf"
    if ext == ".docx":
        return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    if ext == ".xlsx":
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return "application/octet-stream"


# ==================== STARTUP & STATIC ====================

@app.on_event("startup")
def on_startup() -> None:
    """
    Khi server khởi động:
    1. Nếu GDRIVE_BRIDGE_URL được cấu hình, tự động gọi LOAD_DB để khôi phục port_checklist.db từ Google Drive.
    2. Khởi tạo/migrate bảng DB và sinh bộ file mẫu.
    3. Nếu trên Google Drive chưa có bản backup DB, đẩy bản DB khởi tạo lên Google Drive.
    """
    restored_from_cloud = False
    if is_bridge_enabled():
        restored_from_cloud = load_db_from_gdrive(DB_PATH)

    init_db()
    seed_initial_vessels(reset_existing=False)
    generate_all_sample_files()

    if is_bridge_enabled() and not restored_from_cloud:
        sync_db_to_gdrive(DB_PATH)


app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@app.get("/", include_in_schema=False)
def serve_spa():
    index_path = os.path.join(STATIC_DIR, "index.html")
    if not os.path.exists(index_path):
        return JSONResponse({"message": "Frontend index.html is being prepared."})
    return FileResponse(
        index_path,
        headers={
            "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",
            "Pragma": "no-cache",
            "Expires": "0",
        },
    )


# ==================== AUTHENTICATION API ====================

@app.post("/api/auth/login")
def login(payload: LoginRequest, response: Response):
    username = payload.username.strip()
    user = get_user_by_username(username)
    if not user or not verify_password(payload.password, user["password_hash"]):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Tên đăng nhập hoặc mật khẩu không chính xác.",
        )

    if not user.get("is_active"):
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Tài khoản của bạn đã bị khóa bởi Quản trị viên.",
        )

    token = create_session_token(user["username"], user["role"])
    response.set_cookie(
        key=COOKIE_NAME,
        value=token,
        httponly=True,
        samesite="lax",
        max_age=TOKEN_TTL_SECONDS,
        path="/",
    )

    log_audit_event(
        username=user["username"],
        role=user["role"],
        action="LOGIN",
        details=f"Đăng nhập hệ thống ({user['full_name']})",
    )

    return {
        "message": "Đăng nhập thành công.",
        "user": {
            "id": user["id"],
            "username": user["username"],
            "full_name": user["full_name"],
            "role": user["role"],
            "is_active": bool(user["is_active"]),
        },
    }


@app.post("/api/auth/logout")
def logout(response: Response):
    response.delete_cookie(key=COOKIE_NAME, path="/")
    return {"message": "Đã đăng xuất."}


@app.get("/api/auth/me")
def get_me(current_user: Dict[str, Any] = Depends(get_current_user)):
    return {"user": current_user}


# ==================== USER MANAGEMENT & AUDIT LOGS (ADMIN ONLY) ====================

@app.get("/api/users")
def api_list_users(current_user: Dict[str, Any] = Depends(require_roles("ADMIN"))):
    return {"users": list_all_users()}


@app.post("/api/users")
def api_create_user(
    payload: UserCreateRequest,
    background_tasks: BackgroundTasks,
    current_user: Dict[str, Any] = Depends(require_roles("ADMIN")),
):
    uname = payload.username.strip().lower()
    if not uname or len(uname) < 3:
        raise HTTPException(status_code=400, detail="Username phải có ít nhất 3 ký tự.")
    if payload.role not in ("ADMIN", "OPS", "VIEWER"):
        raise HTTPException(status_code=400, detail="Role không hợp lệ (ADMIN, OPS, VIEWER).")
    if not payload.password or len(payload.password) < 6:
        raise HTTPException(status_code=400, detail="Mật khẩu phải có ít nhất 6 ký tự.")
    if get_user_by_username(uname):
        raise HTTPException(status_code=400, detail=f"Tài khoản '{uname}' đã tồn tại.")

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        """
        INSERT INTO users (username, full_name, password_hash, role, is_active, created_at)
        VALUES (?, ?, ?, ?, 1, ?)
        """,
        (uname, payload.full_name.strip() or uname, hash_password(payload.password), payload.role, now_str),
    )
    conn.commit()
    conn.close()

    log_audit_event(
        username=current_user["username"],
        role=current_user["role"],
        action="USER_MGMT",
        details=f"Tạo tài khoản mới '{uname}' với quyền {payload.role}",
    )
    background_tasks.add_task(sync_db_to_gdrive)

    return {"message": f"Đã tạo tài khoản '{uname}'.", "users": list_all_users()}


@app.put("/api/users/{user_id}")
def api_update_user(
    user_id: int,
    payload: UserUpdateRequest,
    background_tasks: BackgroundTasks,
    current_user: Dict[str, Any] = Depends(require_roles("ADMIN")),
):
    target = get_user_by_id(user_id)
    if not target:
        raise HTTPException(status_code=404, detail="Không tìm thấy tài khoản.")

    if target["id"] == current_user["id"]:
        if payload.is_active is False:
            raise HTTPException(status_code=400, detail="Không thể tự khóa tài khoản đang đăng nhập.")
        if payload.role and payload.role != "ADMIN":
            raise HTTPException(status_code=400, detail="Không thể tự hạ quyền Quản trị viên của chính mình.")

    new_role = payload.role if payload.role in ("ADMIN", "OPS", "VIEWER") else target["role"]
    new_active = (1 if payload.is_active else 0) if payload.is_active is not None else target["is_active"]
    new_fullname = payload.full_name.strip() if payload.full_name else target["full_name"]
    new_pw_hash = hash_password(payload.password) if (payload.password and len(payload.password) >= 6) else target["password_hash"]

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        """
        UPDATE users
        SET role = ?, is_active = ?, full_name = ?, password_hash = ?
        WHERE id = ?
        """,
        (new_role, new_active, new_fullname, new_pw_hash, user_id),
    )
    conn.commit()
    conn.close()

    changes = []
    if new_role != target["role"]:
        changes.append(f"Role: {target['role']} -> {new_role}")
    if new_active != target["is_active"]:
        changes.append("Trạng thái: " + ("Mở khóa" if new_active == 1 else "Khóa tài khoản"))
    if payload.password:
        changes.append("Đặt lại mật khẩu")

    log_audit_event(
        username=current_user["username"],
        role=current_user["role"],
        action="USER_MGMT",
        details=f"Cập nhật tài khoản '{target['username']}' ({', '.join(changes) or 'Cập nhật thông tin'})",
    )
    background_tasks.add_task(sync_db_to_gdrive)

    return {"message": f"Đã cập nhật tài khoản '{target['username']}'.", "users": list_all_users()}


@app.get("/api/audit-logs")
def api_get_audit_logs(current_user: Dict[str, Any] = Depends(require_roles("ADMIN"))):
    return {"logs": get_audit_logs(limit=250)}


# ==================== GOOGLE DRIVE BRIDGE ADMIN API ====================

@app.get("/api/gdrive/status")
def api_get_gdrive_status(current_user: Dict[str, Any] = Depends(require_roles("ADMIN", "OPS", "VIEWER"))):
    return get_bridge_status()


@app.post("/api/admin/gdrive-config")
def api_save_gdrive_config(
    payload: GDriveBridgeConfigRequest,
    background_tasks: BackgroundTasks,
    current_user: Dict[str, Any] = Depends(require_roles("ADMIN")),
):
    set_bridge_config(payload.bridge_url, payload.secret_key or "PortCall_Secret_2026")
    if not is_bridge_enabled():
        return {
            "message": "Đã chuyển về chế độ lưu trữ cục bộ (Local Storage).",
            "status": get_bridge_status(),
        }

    ping_res = ping_bridge()
    if not ping_res.get("ok"):
        raise HTTPException(
            status_code=400,
            detail=f"Không thể kết nối tới Google Apps Script Bridge: {ping_res.get('error')}",
        )

    log_audit_event(
        username=current_user["username"],
        role=current_user["role"],
        action="USER_MGMT",
        details="Kết nối và kích hoạt Google Drive Bridge thành công",
    )
    background_tasks.add_task(sync_db_to_gdrive)

    return {
        "message": "Đã kết nối Google Drive Bridge và kích hoạt đồng bộ tự động!",
        "status": get_bridge_status(),
        "ping": ping_res,
    }


@app.post("/api/admin/gdrive-sync")
def api_manual_gdrive_sync(current_user: Dict[str, Any] = Depends(require_roles("ADMIN", "OPS"))):
    if not is_bridge_enabled():
        raise HTTPException(status_code=400, detail="Chưa cấu hình GDRIVE_BRIDGE_URL.")
    res = sync_db_to_gdrive()
    if not res.get("ok"):
        raise HTTPException(status_code=400, detail=f"Đồng bộ thất bại: {res.get('error')}")
    return {
        "message": "Đã đồng bộ Database lên Google Drive thành công.",
        "status": get_bridge_status(),
    }


# ==================== INTERNAL FILE PROCESSING HELPER ====================

def _process_and_store_single_file(
    vessel: Dict[str, Any],
    temp_file_path: str,
    original_filename: str,
    uploaded_by: str,
) -> Dict[str, Any]:
    vessel_id = vessel["id"]
    vessel_name = vessel["vessel_name"]

    abs_vessel_folder = get_vessel_storage_abs_path(vessel_id)
    if not abs_vessel_folder:
        raise HTTPException(status_code=500, detail="Không thể khởi tạo vùng lưu trữ cho tàu.")

    ext = os.path.splitext(original_filename)[1].lower() or ".pdf"

    parsed = parse_and_validate_file(
        file_path=temp_file_path,
        original_filename=original_filename,
        current_vessel_name=vessel_name,
    )
    doc_type = parsed["doc_type"]

    conn = get_connection()
    cursor = conn.cursor()

    seq_num = get_next_sequence_number(vessel_id, doc_type, conn=conn)
    stored_filename = build_standardized_filename(doc_type, vessel_name, seq_num, ext)
    final_abs_path = os.path.join(abs_vessel_folder, stored_filename)

    while os.path.exists(final_abs_path) and os.path.abspath(final_abs_path) != os.path.abspath(temp_file_path):
        seq_num += 1
        stored_filename = build_standardized_filename(doc_type, vessel_name, seq_num, ext)
        final_abs_path = os.path.join(abs_vessel_folder, stored_filename)

    if os.path.abspath(temp_file_path) != os.path.abspath(final_abs_path):
        shutil.move(temp_file_path, final_abs_path)

    file_size = os.path.getsize(final_abs_path) if os.path.exists(final_abs_path) else 0
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    initial_sync_status = "PENDING_SYNC" if is_bridge_enabled() else "LOCAL"

    cursor.execute(
        """
        INSERT INTO uploaded_files (
            vessel_id, original_filename, stored_filename, file_path,
            file_ext, file_size, doc_type, detected_method, detected_keyword,
            detected_vessel_name, is_wrong_vessel, error_code, error_message,
            extracted_count, extracted_codes_json, extracted_text_preview,
            uploaded_by, uploaded_at, sync_status
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            vessel_id,
            original_filename,
            stored_filename,
            stored_filename,
            ext,
            file_size,
            doc_type,
            parsed["detected_method"],
            parsed["detected_keyword"],
            parsed["detected_vessel_name"],
            parsed["is_wrong_vessel"],
            parsed["error_code"],
            parsed["error_message"],
            parsed["extracted_count"],
            json.dumps(parsed["extracted_codes"], ensure_ascii=False),
            parsed["extracted_text_preview"],
            uploaded_by,
            now_str,
            initial_sync_status,
        ),
    )
    file_id = cursor.lastrowid
    conn.commit()
    conn.close()

    return {
        "id": file_id,
        "original_filename": original_filename,
        "stored_filename": stored_filename,
        "abs_path": final_abs_path,
        "ext": ext,
        "doc_type": doc_type,
        "is_wrong_vessel": parsed["is_wrong_vessel"],
        "error_message": parsed["error_message"],
        "extracted_count": parsed["extracted_count"],
        "extracted_codes": parsed["extracted_codes"],
        "uploaded_by": uploaded_by,
        "uploaded_at": now_str,
    }


# ==================== VESSELS & FILES API (PROTECTED BY RBAC) ====================

@app.get("/api/config")
def get_system_config(current_user: Dict[str, Any] = Depends(require_roles("ADMIN", "OPS", "VIEWER"))):
    sample_descriptions = {
        "scan_001.pdf": "Cargo Manifest · Khai báo 5 B/L (BL-101..105)",
        "file_gui_sep.xlsx": "VGM Declaration (Excel) · 3/5 B/L (Thiếu 2 bộ)",
        "vgm_bo_sung.pdf": "VGM Declaration (Bổ sung) · 2/5 B/L -> _(2).pdf",
        "doc_bl_full.pdf": "Bill of Lading · Đủ 5/5 B/L",
        "crew_scan.pdf": "Crew List (PDF)",
        "crew_list_word.docx": "Crew List (Word .docx)",
        "port_clearance_99.pdf": "Port Clearance · Giấy phép rời cảng",
        "file_nham_tau.pdf": "VGM ghi sai tên tàu (MV EVER GIVEN)",
    }
    samples = []
    if os.path.exists(SAMPLE_DIR):
        for fname in sorted(os.listdir(SAMPLE_DIR)):
            fpath = os.path.join(SAMPLE_DIR, fname)
            if os.path.isfile(fpath):
                samples.append(
                    {
                        "filename": fname,
                        "size": os.path.getsize(fpath),
                        "ext": os.path.splitext(fname)[1].lower(),
                        "description": sample_descriptions.get(fname, "Chứng từ mẫu"),
                    }
                )

    return {
        "vessel_types": VESSEL_TYPES,
        "doc_types": DOC_TYPE_INFO,
        "sample_files": samples,
        "gdrive_status": get_bridge_status(),
    }


@app.get("/api/vessels")
def list_vessels(current_user: Dict[str, Any] = Depends(require_roles("ADMIN", "OPS", "VIEWER"))):
    vessels = get_all_vessels()
    stats = {
        "total": len(vessels),
        "red": sum(1 for v in vessels if v["status_color"] == "RED"),
        "orange": sum(1 for v in vessels if v["status_color"] == "ORANGE"),
        "green": sum(1 for v in vessels if v["status_color"] == "GREEN"),
        "total_files": sum(len(v.get("files", [])) for v in vessels),
    }
    return {"vessels": vessels, "stats": stats, "gdrive_status": get_bridge_status()}


@app.get("/api/vessels/{vessel_id}")
def get_single_vessel(
    vessel_id: int,
    current_user: Dict[str, Any] = Depends(require_roles("ADMIN", "OPS", "VIEWER")),
):
    vessel = get_vessel_detail(vessel_id)
    if not vessel:
        raise HTTPException(status_code=404, detail="Không tìm thấy tàu.")
    return vessel


@app.get("/api/files/{file_id}/download")
def download_file_by_id(
    file_id: int,
    current_user: Dict[str, Any] = Depends(require_roles("ADMIN", "OPS", "VIEWER")),
):
    """
    Tải hoặc xem file an toàn thông qua file_id.
    Nếu container Render vừa khởi động lại làm mất file cục bộ, tự động kéo file từ Google Drive về!
    """
    rec = get_file_record_internal(file_id)
    if not rec or not rec.get("abs_file_path"):
        raise HTTPException(status_code=404, detail="Không tìm thấy chứng từ.")

    abs_file = rec["abs_file_path"]
    if not abs_file.startswith(os.path.abspath(STORAGE_ROOT)):
        raise HTTPException(status_code=403, detail="Truy cập file bị từ chối.")

    if not os.path.exists(abs_file) and rec.get("gdrive_file_id"):
        download_doc_from_gdrive_if_missing(rec["gdrive_file_id"], abs_file)

    if not os.path.exists(abs_file):
        raise HTTPException(status_code=404, detail="File chứng từ không tồn tại trên bộ nhớ đệm.")

    media_type = _mime_for_ext(rec.get("file_ext") or "")
    return FileResponse(
        path=abs_file,
        filename=rec["stored_filename"],
        media_type=media_type,
    )


# ==================== EXPORT INSPECTION REPORT (PDF & DOCX) ====================

def _status_text_vn(color: str) -> str:
    if color == "GREEN":
        return "Đủ hồ sơ (Hoàn tất 100%)"
    if color == "ORANGE":
        return "Thiếu chứng từ (Đang bổ sung)"
    return "Cần xử lý (Chưa có chứng từ / Cảnh báo sai tàu)"


def _item_status_vn(st: str) -> str:
    if st == "COMPLETED":
        return "Đã đủ"
    if st == "INCOMPLETE":
        return "Đang thiếu"
    return "Chưa nộp"


@app.get("/api/vessels/{vessel_id}/export/pdf")
def export_vessel_checklist_pdf(
    vessel_id: int,
    current_user: Dict[str, Any] = Depends(require_roles("ADMIN", "OPS", "VIEWER")),
):
    vessel = get_vessel_detail(vessel_id)
    if not vessel:
        raise HTTPException(status_code=404, detail="Không tìm thấy tàu.")

    doc = fitz.open()
    page = doc.new_page(width=595, height=842)

    win_font = "C:/Windows/Fonts/arial.ttf"
    use_unicode = os.path.exists(win_font)
    fname = "arial_uni" if use_unicode else "helv"
    ffile = win_font if use_unicode else None

    def draw_txt(x: float, y: float, text: str, size: float = 10, color=(0.06, 0.09, 0.16)):
        page.insert_text(
            fitz.Point(x, y),
            str(text),
            fontsize=size,
            fontname=fname,
            fontfile=ffile,
            color=color,
        )

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    page.draw_rect(fitz.Rect(36, 32, 559, 92), color=(0.06, 0.09, 0.16), fill=(0.97, 0.98, 0.99), width=1)
    draw_txt(48, 56, "PHIẾU KIỂM TRA HỒ SƠ TÀU CẬP CẢNG", size=14, color=(0.06, 0.09, 0.16))
    draw_txt(48, 76, "PORT-CALL DOCUMENT VERIFICATION REPORT · PORTCALL DOCS", size=8.5, color=(0.39, 0.45, 0.55))

    y = 115
    draw_txt(40, y, f"Tên tàu: {vessel['vessel_name']}", size=11)
    draw_txt(310, y, f"Loại tàu: {vessel['vessel_type_label']}", size=10.5)
    y += 20
    draw_txt(40, y, f"Giờ dự kiến đến (ETA): {vessel['eta']}", size=10)
    draw_txt(310, y, f"Tổng số B/L theo Manifest: {vessel.get('expected_bl_count', 0)} B/L", size=10)
    y += 20

    st_color = (0.06, 0.65, 0.42) if vessel["status_color"] == "GREEN" else (
        (0.85, 0.45, 0.02) if vessel["status_color"] == "ORANGE" else (0.86, 0.15, 0.15)
    )
    draw_txt(40, y, f"Trạng thái tổng thể: {_status_text_vn(vessel['status_color'])}", size=10.5, color=st_color)
    draw_txt(310, y, f"Tiến độ: {vessel['completed_checklist_count']}/{vessel['total_checklist_count']} hạng mục", size=10)
    y += 20
    draw_txt(
        40,
        y,
        f"Người kiểm tra: {current_user['full_name']} ({current_user['username']} · {current_user['role']})",
        size=9.5,
        color=(0.28, 0.33, 0.41),
    )
    draw_txt(380, y, f"Ngày xuất: {now_str}", size=9.5, color=(0.28, 0.33, 0.41))

    y += 24
    page.draw_rect(fitz.Rect(36, y, 559, y + 24), color=(0.82, 0.85, 0.90), fill=(0.94, 0.96, 0.98), width=0.8)
    draw_txt(42, y + 16, "MÃ CHUẨN", size=8.5, color=(0.2, 0.25, 0.33))
    draw_txt(125, y + 16, "LOẠI CHỨNG TỪ", size=8.5, color=(0.2, 0.25, 0.33))
    draw_txt(255, y + 16, "TRẠNG THÁI", size=8.5, color=(0.2, 0.25, 0.33))
    draw_txt(325, y + 16, "ĐỐI CHIẾU", size=8.5, color=(0.2, 0.25, 0.33))
    draw_txt(400, y + 16, "CHỨNG TỪ ĐÃ NỘP", size=8.5, color=(0.2, 0.25, 0.33))
    y += 24

    for item in vessel.get("checklist", []):
        files = item.get("files", [])
        row_h = max(28, 18 * max(1, len(files)) + 8)
        page.draw_rect(fitz.Rect(36, y, 559, y + row_h), color=(0.88, 0.91, 0.94), fill=None, width=0.6)

        item_st = _item_status_vn(item["status"])
        c_st = (0.06, 0.65, 0.42) if item["status"] == "COMPLETED" else (
            (0.85, 0.45, 0.02) if item["status"] == "INCOMPLETE" else (0.86, 0.15, 0.15)
        )

        if item["requires_count_match"] == 1:
            match_str = f"{item['actual_count']}/{item['expected_count']} B/L"
        elif item["doc_type"] == "MANIFEST":
            match_str = f"{item['expected_count']} B/L" if item["status"] == "COMPLETED" else "0 B/L"
        else:
            match_str = "1/1" if item["status"] == "COMPLETED" else "0/1"

        draw_txt(42, y + 17, item["doc_type"], size=9)
        draw_txt(125, y + 17, item["doc_name"][:24], size=9)
        draw_txt(255, y + 17, item_st, size=9, color=c_st)
        draw_txt(325, y + 17, match_str, size=9)

        if not files:
            draw_txt(400, y + 17, "— Chưa nộp", size=8.5, color=(0.55, 0.6, 0.68))
        else:
            fy = y + 16
            for f in files:
                draw_txt(400, fy, f"{f['stored_filename'][:28]} ({f.get('uploaded_by', 'ops')})", size=8)
                fy += 15

        y += row_h

    wrong_files = vessel.get("wrong_vessel_files", [])
    unclass_files = vessel.get("unclassified_files", [])
    if wrong_files or unclass_files:
        y += 18
        draw_txt(40, y, "CẢNH BÁO KIỂM TRA CHÉO:", size=10, color=(0.86, 0.15, 0.15))
        y += 16
        for wf in wrong_files:
            draw_txt(46, y, f"• [SAI TÀU] {wf['original_filename']}: {wf['error_message']}", size=9, color=(0.86, 0.15, 0.15))
            y += 15
        for uf in unclass_files:
            draw_txt(46, y, f"• [CHƯA PHÂN LOẠI] {uf['original_filename']}: Chờ chọn loại chứng từ", size=9, color=(0.85, 0.45, 0.02))
            y += 15

    sig_y = min(740, max(y + 45, 650))
    draw_txt(65, sig_y, "ĐẠI DIỆN ĐẠI LÝ HÃNG TÀU", size=9.5)
    draw_txt(65, sig_y + 14, "(Ký, ghi rõ họ tên)", size=8.5, color=(0.45, 0.5, 0.58))

    draw_txt(360, sig_y, "CHUYÊN VIÊN KIỂM TRA HỒ SƠ", size=9.5)
    draw_txt(360, sig_y + 14, "(Ký, ghi rõ họ tên)", size=8.5, color=(0.45, 0.5, 0.58))
    draw_txt(360, sig_y + 62, current_user["full_name"], size=9.5)

    pdf_bytes = doc.tobytes()
    doc.close()

    safe_vname = sanitize_token(vessel["vessel_name"])
    filename = f"Phieu_Kiem_Tra_{safe_vname}.pdf"
    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.get("/api/vessels/{vessel_id}/export/docx")
def export_vessel_checklist_docx(
    vessel_id: int,
    current_user: Dict[str, Any] = Depends(require_roles("ADMIN", "OPS", "VIEWER")),
):
    vessel = get_vessel_detail(vessel_id)
    if not vessel:
        raise HTTPException(status_code=404, detail="Không tìm thấy tàu.")

    doc = docx.Document()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = title_p.add_run("PHIẾU KIỂM TRA HỒ SƠ TÀU CẬP CẢNG\n")
    r_title.bold = True
    r_title.font.size = Pt(15)
    r_sub = title_p.add_run("PORT-CALL DOCUMENT VERIFICATION REPORT")
    r_sub.font.size = Pt(9.5)
    r_sub.font.color.rgb = RGBColor(100, 116, 139)

    info_p = doc.add_paragraph()
    info_p.add_run("Tên tàu: ").bold = True
    info_p.add_run(f"{vessel['vessel_name']}\t\t")
    info_p.add_run("Loại tàu: ").bold = True
    info_p.add_run(f"{vessel['vessel_type_label']}\n")

    info_p.add_run("Giờ dự kiến cập cảng (ETA): ").bold = True
    info_p.add_run(f"{vessel['eta']}\t\t")
    info_p.add_run("Tổng B/L theo Manifest: ").bold = True
    info_p.add_run(f"{vessel.get('expected_bl_count', 0)} B/L\n")

    info_p.add_run("Trạng thái tổng thể: ").bold = True
    info_p.add_run(f"{_status_text_vn(vessel['status_color'])} ({vessel['completed_checklist_count']}/{vessel['total_checklist_count']} hạng mục)\n")

    info_p.add_run("Người kiểm tra: ").bold = True
    info_p.add_run(f"{current_user['full_name']} ({current_user['username']} - {current_user['role']})\t\t")
    info_p.add_run("Thời gian xuất: ").bold = True
    info_p.add_run(now_str)

    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    headers_text = ["Mã chuẩn", "Loại chứng từ", "Trạng thái", "Đối chiếu", "Chứng từ đã nộp & Người up"]
    for idx, h_txt in enumerate(headers_text):
        hdr[idx].text = h_txt
        for run in hdr[idx].paragraphs[0].runs:
            run.bold = True

    for item in vessel.get("checklist", []):
        row_cells = table.add_row().cells
        row_cells[0].text = item["doc_type"]
        row_cells[1].text = item["doc_name"]
        row_cells[2].text = _item_status_vn(item["status"])

        if item["requires_count_match"] == 1:
            match_str = f"{item['actual_count']}/{item['expected_count']} B/L"
            if item["expected_count"] > 0 and item["actual_count"] < item["expected_count"]:
                match_str += f" (Thiếu {item['expected_count'] - item['actual_count']} bộ)"
        elif item["doc_type"] == "MANIFEST":
            match_str = f"{item['expected_count']} B/L" if item["status"] == "COMPLETED" else "0 B/L"
        else:
            match_str = "1/1" if item["status"] == "COMPLETED" else "0/1"

        row_cells[3].text = match_str

        files = item.get("files", [])
        if not files:
            row_cells[4].text = "— Chưa nộp"
        else:
            file_lines = [
                f"{f['stored_filename']} (gốc: {f['original_filename']} · {f.get('uploaded_by', 'ops')} · {f.get('uploaded_at', '')})"
                for f in files
            ]
            row_cells[4].text = "\n".join(file_lines)

    wrong_files = vessel.get("wrong_vessel_files", [])
    unclass_files = vessel.get("unclassified_files", [])
    if wrong_files or unclass_files:
        wp = doc.add_paragraph()
        w_run = wp.add_run("\nCẢNH BÁO LỖI HỒ SƠ:\n")
        w_run.bold = True
        w_run.font.color.rgb = RGBColor(220, 38, 38)
        for wf in wrong_files:
            wp.add_run(f"• [SAI TÀU] {wf['original_filename']}: {wf['error_message']}\n")
        for uf in unclass_files:
            wp.add_run(f"• [CHƯA PHÂN LOẠI] {uf['original_filename']}: Chưa nhận diện được loại chứng từ.\n")

    sig_table = doc.add_table(rows=1, cols=2)
    s_cells = sig_table.rows[0].cells
    p_left = s_cells[0].paragraphs[0]
    p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_left.add_run("\nĐẠI DIỆN ĐẠI LÝ HÃNG TÀU\n").bold = True
    p_left.add_run("(Ký, ghi rõ họ tên)\n\n\n")

    p_right = s_cells[1].paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_right.add_run("\nCHUYÊN VIÊN KIỂM TRA HỒ SƠ\n").bold = True
    p_right.add_run("(Ký, ghi rõ họ tên)\n\n\n")
    p_right.add_run(current_user["full_name"]).bold = True

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)

    safe_vname = sanitize_token(vessel["vessel_name"])
    filename = f"Phieu_Kiem_Tra_{safe_vname}.docx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# ==================== VESSEL MUTATION ENDPOINTS (WITH GDRIVE BRIDGE SYNC) ====================

@app.post("/api/vessels/batch")
def create_vessels_batch(
    payload: VesselBatchCreateRequest,
    background_tasks: BackgroundTasks,
    current_user: Dict[str, Any] = Depends(require_roles("ADMIN", "OPS")),
):
    if not payload.vessels:
        raise HTTPException(status_code=400, detail="Danh sách tàu cần tạo không được để trống.")

    created = []
    for item in payload.vessels:
        if not item.vessel_name or not item.vessel_name.strip():
            continue
        v = create_vessel(
            vessel_name=item.vessel_name,
            eta=item.eta or datetime.now().strftime("%Y-%m-%d %H:%M"),
            vessel_type=item.vessel_type or "CONTAINER_INTL",
        )
        created.append(v)
        folder_meta = get_vessel_folder_meta_internal(v["id"])
        background_tasks.add_task(bg_create_vessel_folder_and_sync, v["id"], folder_meta["folder_name"])

    if not created:
        raise HTTPException(status_code=400, detail="Vui lòng nhập tên tàu hợp lệ.")

    names_str = ", ".join([v["vessel_name"] for v in created])
    log_audit_event(
        username=current_user["username"],
        role=current_user["role"],
        action="CREATE_VESSEL",
        details=f"Tạo {len(created)} lịch tàu: {names_str}",
    )

    return {
        "message": f"Đã tạo {len(created)} lịch tàu.",
        "created_vessels": created,
    }


@app.delete("/api/vessels/{vessel_id}")
def delete_vessel(
    vessel_id: int,
    background_tasks: BackgroundTasks,
    current_user: Dict[str, Any] = Depends(require_roles("ADMIN")),
):
    vessel = get_vessel_detail(vessel_id)
    if not vessel:
        raise HTTPException(status_code=404, detail="Không tìm thấy tàu.")

    folder_meta = get_vessel_folder_meta_internal(vessel_id)
    abs_folder = get_vessel_storage_abs_path(vessel_id)
    if abs_folder and os.path.exists(abs_folder):
        shutil.rmtree(abs_folder, ignore_errors=True)

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM uploaded_files WHERE vessel_id = ?", (vessel_id,))
    cursor.execute("DELETE FROM checklist_items WHERE vessel_id = ?", (vessel_id,))
    cursor.execute("DELETE FROM vessels WHERE id = ?", (vessel_id,))
    conn.commit()
    conn.close()

    log_audit_event(
        username=current_user["username"],
        role=current_user["role"],
        action="DELETE_FILE",
        details=f"Xóa tàu {vessel['vessel_name']} và toàn bộ chứng từ liên quan",
    )
    background_tasks.add_task(
        bg_delete_vessel_folder_and_sync,
        folder_meta["gdrive_folder_id"],
        folder_meta["folder_name"],
    )

    return {"message": f"Đã xóa tàu {vessel['vessel_name']}."}


@app.delete("/api/vessels/{vessel_id}/clear-files")
def clear_vessel_files(
    vessel_id: int,
    background_tasks: BackgroundTasks,
    current_user: Dict[str, Any] = Depends(require_roles("ADMIN", "OPS")),
):
    vessel = get_vessel_detail(vessel_id)
    if not vessel:
        raise HTTPException(status_code=404, detail="Không tìm thấy tàu.")

    folder_meta = get_vessel_folder_meta_internal(vessel_id)
    abs_folder = get_vessel_storage_abs_path(vessel_id)
    if abs_folder and os.path.exists(abs_folder):
        for fname in os.listdir(abs_folder):
            fpath = os.path.join(abs_folder, fname)
            if os.path.isfile(fpath):
                try:
                    os.remove(fpath)
                except Exception:
                    pass

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM uploaded_files WHERE vessel_id = ?", (vessel_id,))
    conn.commit()
    conn.close()

    log_audit_event(
        username=current_user["username"],
        role=current_user["role"],
        action="DELETE_FILE",
        details=f"Làm trống toàn bộ hồ sơ của tàu {vessel['vessel_name']}",
    )

    updated_vessel = recalculate_vessel_status(vessel_id)
    background_tasks.add_task(
        bg_clear_vessel_folder_and_sync,
        folder_meta["gdrive_folder_id"],
        folder_meta["folder_name"],
    )

    return {
        "message": f"Đã làm trống hồ sơ tàu {vessel['vessel_name']}.",
        "vessel": updated_vessel,
    }


@app.post("/api/vessels/{vessel_id}/upload")
async def upload_files_to_vessel(
    vessel_id: int,
    background_tasks: BackgroundTasks,
    files: List[UploadFile] = File(...),
    current_user: Dict[str, Any] = Depends(require_roles("ADMIN", "OPS")),
):
    vessel = get_vessel_detail(vessel_id)
    if not vessel:
        raise HTTPException(status_code=404, detail="Không tìm thấy tàu.")

    abs_vessel_folder = get_vessel_storage_abs_path(vessel_id)
    if not abs_vessel_folder:
        raise HTTPException(status_code=500, detail="Không thể truy cập vùng lưu trữ của tàu.")

    folder_meta = get_vessel_folder_meta_internal(vessel_id)
    processed_results = []
    for upload_file in files:
        content = await upload_file.read()
        clean_orig_name = validate_file_security(upload_file.filename or "document.pdf", content)

        temp_name = f"__tmp_{datetime.now().strftime('%H%M%S%f')}_{clean_orig_name}"
        temp_path = os.path.join(abs_vessel_folder, temp_name)

        with open(temp_path, "wb") as f_out:
            f_out.write(content)

        res = _process_and_store_single_file(
            vessel=vessel,
            temp_file_path=temp_path,
            original_filename=clean_orig_name,
            uploaded_by=current_user["username"],
        )
        processed_results.append(res)

        log_audit_event(
            username=current_user["username"],
            role=current_user["role"],
            action="UPLOAD_FILE",
            details=(
                f"Tàu {vessel['vessel_name']}: Tải lên '{clean_orig_name}' -> "
                f"'{res['stored_filename']}' ({res['doc_type']})"
                + (" [CẢNH BÁO: WRONG_VESSEL]" if res["is_wrong_vessel"] else "")
            ),
        )

        background_tasks.add_task(
            bg_upload_doc_and_sync,
            res["id"],
            vessel_id,
            folder_meta["folder_name"],
            res["abs_path"],
            res["stored_filename"],
            _mime_for_ext(res["ext"]),
        )

    updated_vessel = recalculate_vessel_status(vessel_id)
    return {
        "message": f"Đã xử lý {len(processed_results)} chứng từ.",
        "processed_files": processed_results,
        "vessel": updated_vessel,
    }


@app.post("/api/vessels/{vessel_id}/simulate-sample")
def simulate_sample_upload(
    vessel_id: int,
    payload: SimulateSampleRequest,
    background_tasks: BackgroundTasks = BackgroundTasks(),
    current_user: Dict[str, Any] = Depends(require_roles("ADMIN", "OPS")),
):
    vessel = get_vessel_detail(vessel_id)
    if not vessel:
        raise HTTPException(status_code=404, detail="Không tìm thấy tàu.")

    abs_vessel_folder = get_vessel_storage_abs_path(vessel_id)
    if not abs_vessel_folder:
        raise HTTPException(status_code=500, detail="Không thể truy cập vùng lưu trữ của tàu.")

    folder_meta = get_vessel_folder_meta_internal(vessel_id)
    processed_results = []
    for fname in payload.filenames:
        if fname == "__UNCLASSIFIED_DEMO__":
            orig_name = "scan_anh_mo_009.pdf"
            temp_path = os.path.join(abs_vessel_folder, f"__tmp_{datetime.now().strftime('%H%M%S%f')}_{orig_name}")
            create_pdf_file(
                temp_path,
                title="SCANNED DOCUMENT - PAGE 01",
                subtitle="Optical Scan Image - Low Contrast Maritime Attachment",
                lines=[
                    f"Vessel Name: {vessel['vessel_name']}",
                    "Reference ID: DOC-2026-99812",
                    "===",
                    "This scanned page contains no standard classification anchor keywords.",
                ],
            )
            res = _process_and_store_single_file(
                vessel=vessel,
                temp_file_path=temp_path,
                original_filename=orig_name,
                uploaded_by=current_user["username"],
            )
            processed_results.append(res)
            log_audit_event(
                username=current_user["username"],
                role=current_user["role"],
                action="UPLOAD_FILE",
                details=f"Tàu {vessel['vessel_name']}: Nạp file mẫu '{orig_name}' -> '{res['stored_filename']}'",
            )
            if background_tasks is not None:
                background_tasks.add_task(
                    bg_upload_doc_and_sync,
                    res["id"],
                    vessel_id,
                    folder_meta["folder_name"],
                    res["abs_path"],
                    res["stored_filename"],
                    _mime_for_ext(res["ext"]),
                )
            continue

        clean_fname = sanitize_filename(fname)
        sample_src = os.path.join(SAMPLE_DIR, clean_fname)
        if not os.path.exists(sample_src):
            raise HTTPException(status_code=404, detail=f"Không tìm thấy chứng từ mẫu '{clean_fname}'.")

        with open(sample_src, "rb") as f_in:
            content = f_in.read()
        validate_file_security(clean_fname, content)

        temp_path = os.path.join(abs_vessel_folder, f"__tmp_{datetime.now().strftime('%H%M%S%f')}_{clean_fname}")
        with open(temp_path, "wb") as f_out:
            f_out.write(content)

        res = _process_and_store_single_file(
            vessel=vessel,
            temp_file_path=temp_path,
            original_filename=clean_fname,
            uploaded_by=current_user["username"],
        )
        processed_results.append(res)

        log_audit_event(
            username=current_user["username"],
            role=current_user["role"],
            action="UPLOAD_FILE",
            details=(
                f"Tàu {vessel['vessel_name']}: Nạp chứng từ '{clean_fname}' -> "
                f"'{res['stored_filename']}' ({res['doc_type']})"
                + (" [WRONG_VESSEL]" if res["is_wrong_vessel"] else "")
            ),
        )
        if background_tasks is not None:
            background_tasks.add_task(
                bg_upload_doc_and_sync,
                res["id"],
                vessel_id,
                folder_meta["folder_name"],
                res["abs_path"],
                res["stored_filename"],
                _mime_for_ext(res["ext"]),
            )

    updated_vessel = recalculate_vessel_status(vessel_id)
    return {
        "message": f"Đã nạp {len(processed_results)} chứng từ mẫu.",
        "processed_files": processed_results,
        "vessel": updated_vessel,
    }


@app.put("/api/files/{file_id}/classify")
def manual_classify_file(
    file_id: int,
    payload: ClassifyFileRequest,
    background_tasks: BackgroundTasks,
    current_user: Dict[str, Any] = Depends(require_roles("ADMIN", "OPS")),
):
    new_doc_type = payload.doc_type.strip().upper()
    if new_doc_type not in DOC_TYPE_INFO or new_doc_type == "UNCLASSIFIED":
        raise HTTPException(status_code=400, detail="Loại chứng từ chọn không hợp lệ.")

    rec = get_file_record_internal(file_id)
    if not rec:
        raise HTTPException(status_code=404, detail="Không tìm thấy chứng từ.")

    vessel_id = rec["vessel_id"]
    abs_vessel_folder = rec["abs_vessel_folder"]
    old_abs_path = rec["abs_file_path"]

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT COUNT(*) AS cnt FROM uploaded_files WHERE vessel_id = ? AND doc_type = ? AND id != ?",
        (vessel_id, new_doc_type, file_id),
    )
    seq_num = (cursor.fetchone()["cnt"] or 0) + 1
    ext = rec["file_ext"] or os.path.splitext(rec["original_filename"])[1] or ".pdf"

    new_stored_filename = build_standardized_filename(new_doc_type, rec["vessel_name"], seq_num, ext)
    new_abs_path = os.path.join(abs_vessel_folder, new_stored_filename)

    while os.path.exists(new_abs_path) and os.path.abspath(new_abs_path) != os.path.abspath(old_abs_path):
        seq_num += 1
        new_stored_filename = build_standardized_filename(new_doc_type, rec["vessel_name"], seq_num, ext)
        new_abs_path = os.path.join(abs_vessel_folder, new_stored_filename)

    if os.path.exists(old_abs_path) and os.path.abspath(old_abs_path) != os.path.abspath(new_abs_path):
        shutil.move(old_abs_path, new_abs_path)

    target_parse_path = new_abs_path if os.path.exists(new_abs_path) else old_abs_path
    parsed = parse_and_validate_file(
        file_path=target_parse_path,
        original_filename=rec["original_filename"],
        current_vessel_name=rec["vessel_name"],
        forced_doc_type=new_doc_type,
    )

    cursor.execute(
        """
        UPDATE uploaded_files
        SET stored_filename = ?,
            file_path = ?,
            doc_type = ?,
            detected_method = 'MANUAL_1CLICK',
            detected_keyword = ?,
            detected_vessel_name = ?,
            is_wrong_vessel = ?,
            error_code = ?,
            error_message = ?,
            extracted_count = ?,
            extracted_codes_json = ?
        WHERE id = ?
        """,
        (
            new_stored_filename,
            new_stored_filename,
            new_doc_type,
            f"1-Click: {new_doc_type}",
            parsed["detected_vessel_name"],
            parsed["is_wrong_vessel"],
            parsed["error_code"],
            parsed["error_message"],
            parsed["extracted_count"],
            json.dumps(parsed["extracted_codes"], ensure_ascii=False),
            file_id,
        ),
    )
    conn.commit()
    conn.close()

    log_audit_event(
        username=current_user["username"],
        role=current_user["role"],
        action="UPLOAD_FILE",
        details=f"Phân loại chứng từ '{rec['original_filename']}' -> '{new_stored_filename}' ({new_doc_type})",
    )

    updated_vessel = recalculate_vessel_status(vessel_id)
    background_tasks.add_task(
        bg_rename_doc_and_sync,
        file_id,
        vessel_id,
        rec.get("gdrive_file_id", ""),
        new_stored_filename,
        rec.get("vessel_folder_name", ""),
        new_abs_path,
        _mime_for_ext(ext),
    )

    return {
        "message": f"Đã đổi phân loại thành '{new_stored_filename}'.",
        "vessel": updated_vessel,
    }


@app.delete("/api/files/{file_id}")
def delete_uploaded_file(
    file_id: int,
    background_tasks: BackgroundTasks,
    current_user: Dict[str, Any] = Depends(require_roles("ADMIN", "OPS")),
):
    rec = get_file_record_internal(file_id)
    if not rec:
        raise HTTPException(status_code=404, detail="Không tìm thấy chứng từ cần xóa.")

    vessel_id = rec["vessel_id"]
    abs_file = rec.get("abs_file_path")
    if abs_file and os.path.exists(abs_file):
        try:
            os.remove(abs_file)
        except Exception:
            pass

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM uploaded_files WHERE id = ?", (file_id,))
    conn.commit()
    conn.close()

    log_audit_event(
        username=current_user["username"],
        role=current_user["role"],
        action="DELETE_FILE",
        details=f"Xóa chứng từ '{rec['stored_filename']}' (gốc: '{rec['original_filename']}') khỏi tàu {rec.get('vessel_name', '')}",
    )

    updated_vessel = recalculate_vessel_status(vessel_id)
    background_tasks.add_task(
        bg_delete_doc_and_sync,
        rec.get("gdrive_file_id", ""),
        rec.get("vessel_folder_name", ""),
        rec["stored_filename"],
    )

    return {
        "message": f"Đã gỡ chứng từ '{rec['stored_filename']}'.",
        "vessel": updated_vessel,
    }


@app.post("/api/demo/reset")
def reset_demo_environment(
    background_tasks: BackgroundTasks,
    current_user: Dict[str, Any] = Depends(require_roles("ADMIN", "OPS")),
):
    seed_initial_vessels(reset_existing=True)
    generate_all_sample_files()
    vessels = get_all_vessels()
    log_audit_event(
        username=current_user["username"],
        role=current_user["role"],
        action="CREATE_VESSEL",
        details="Khôi phục dữ liệu tàu mẫu mặc định (MV BINH HOA & MV HAI AN)",
    )
    background_tasks.add_task(sync_db_to_gdrive)
    return {
        "message": "Đã khôi phục dữ liệu mẫu mặc định.",
        "vessels": vessels,
    }


if __name__ == "__main__":
    import socket
    import uvicorn

    env_port = os.environ.get("PORT")
    if env_port and env_port.isdigit():
        chosen_port = int(env_port)
        host_addr = "0.0.0.0"
    else:
        host_addr = "127.0.0.1"
        preferred_ports = [8000, 8080, 8001, 8888, 5000]
        chosen_port = 8080
        for p in preferred_ports:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                try:
                    s.bind(("127.0.0.1", p))
                    chosen_port = p
                    break
                except OSError:
                    continue

    print(f"Starting PortCall Docs Enterprise MVP at: http://{host_addr}:{chosen_port}")
    uvicorn.run(app, host=host_addr, port=chosen_port)
