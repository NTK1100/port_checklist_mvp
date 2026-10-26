import os
import re
import time
import json
import hmac
import base64
import hashlib
import secrets
from typing import Dict, Any, Optional, Callable

from fastapi import Request, HTTPException, status, Depends

COOKIE_NAME = "portcall_session"
SECRET_KEY = os.environ.get("PORTCALL_SECRET_KEY", "portcall-enterprise-hmac-sha256-secret-key-2026").encode("utf-8")
TOKEN_TTL_SECONDS = 86400 * 7  # 7 ngày
MAX_UPLOAD_SIZE_BYTES = 15 * 1024 * 1024  # 15MB


def hash_password(password: str, salt: Optional[str] = None) -> str:
    """Băm mật khẩu an toàn bằng PBKDF2-HMAC-SHA256 với salt ngẫu nhiên."""
    if not salt:
        salt = secrets.token_hex(16)
    dk = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt.encode("utf-8"),
        120_000,
    )
    return f"pbkdf2_sha256${salt}${dk.hex()}"


def verify_password(password: str, stored_hash: str) -> bool:
    """Kiểm tra mật khẩu chống Timing Attack bằng hmac.compare_digest."""
    try:
        algo, salt, _ = stored_hash.split("$", 2)
        if algo != "pbkdf2_sha256":
            return False
        candidate = hash_password(password, salt)
        return hmac.compare_digest(candidate, stored_hash)
    except Exception:
        return False


def create_session_token(username: str, role: str) -> str:
    """Tạo Session Token có chữ ký HMAC-SHA256."""
    payload = {
        "sub": username,
        "role": role,
        "iat": int(time.time()),
        "exp": int(time.time()) + TOKEN_TTL_SECONDS,
    }
    payload_bytes = json.dumps(payload, separators=(",", ":")).encode("utf-8")
    payload_b64 = base64.urlsafe_b64encode(payload_bytes).decode("ascii").rstrip("=")
    sig = hmac.new(SECRET_KEY, payload_b64.encode("ascii"), hashlib.sha256).hexdigest()
    return f"{payload_b64}.{sig}"


def verify_session_token(token: str) -> Optional[Dict[str, Any]]:
    """Xác thực chữ ký HMAC-SHA256 và thời hạn của Session Token."""
    if not token or "." not in token:
        return None
    try:
        payload_b64, sig = token.rsplit(".", 1)
        expected_sig = hmac.new(SECRET_KEY, payload_b64.encode("ascii"), hashlib.sha256).hexdigest()
        if not hmac.compare_digest(sig, expected_sig):
            return None
        padding = "=" * (-len(payload_b64) % 4)
        payload_bytes = base64.urlsafe_b64decode(payload_b64 + padding)
        payload = json.loads(payload_bytes.decode("utf-8"))
        if int(payload.get("exp", 0)) < int(time.time()):
            return None
        return payload
    except Exception:
        return None


def sanitize_filename(filename: str) -> str:
    """Chống Path Traversal bằng os.path.basename và loại bỏ ký tự điều khiển."""
    base = os.path.basename(filename or "document.pdf").strip()
    base = re.sub(r"[\x00-\x1f\x7f<>:\"/\\|?*]", "_", base)
    return base or "document.pdf"


def validate_file_security(raw_filename: str, content: bytes) -> str:
    """
    Kiểm tra bảo mật file tải lên (AppSec):
    1. Sanitize tên file chống Path Traversal.
    2. Giới hạn dung lượng tối đa 15MB/file.
    3. Kiểm tra Magic Bytes thực tế (%PDF- cho .pdf, PK\\x03\\x04 cho .xlsx và .docx).
    """
    clean_name = sanitize_filename(raw_filename)
    ext = os.path.splitext(clean_name)[1].lower()

    if ext not in (".pdf", ".xlsx", ".docx"):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Định dạng file '{clean_name}' không được phép. Chỉ chấp nhận .pdf, .xlsx và .docx.",
        )

    if len(content) == 0:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"File '{clean_name}' rỗng (0 byte).",
        )

    if len(content) > MAX_UPLOAD_SIZE_BYTES:
        raise HTTPException(
            status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
            detail=f"File '{clean_name}' vượt quá dung lượng tối đa 15MB.",
        )

    # Kiểm tra Magic Bytes
    if ext == ".pdf":
        if not content.startswith(b"%PDF-"):
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Từ chối bảo mật: File '{clean_name}' không có PDF Magic Bytes (%PDF-) hợp lệ.",
            )
    elif ext in (".xlsx", ".docx"):
        if not content.startswith(b"PK\x03\x04"):
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Từ chối bảo mật: File '{clean_name}' không có OpenXML Magic Bytes (PK\\x03\\x04) hợp lệ.",
            )

    return clean_name


def get_current_user(request: Request) -> Dict[str, Any]:
    """Dependency kiểm tra xác thực người dùng từ HttpOnly Cookie."""
    from database import get_user_by_username

    token = request.cookies.get(COOKIE_NAME)
    if not token:
        auth_header = request.headers.get("Authorization", "")
        if auth_header.startswith("Bearer "):
            token = auth_header[7:].strip()

    if not token:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Vui lòng đăng nhập để tiếp tục.",
        )

    payload = verify_session_token(token)
    if not payload or not payload.get("sub"):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Phiên đăng nhập không hợp lệ hoặc đã hết hạn.",
        )

    user = get_user_by_username(payload["sub"])
    if not user:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Tài khoản không tồn tại.",
        )

    if not user.get("is_active"):
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Tài khoản đã bị khóa bởi Quản trị viên.",
        )

    return {
        "id": user["id"],
        "username": user["username"],
        "full_name": user["full_name"],
        "role": user["role"],
        "is_active": bool(user["is_active"]),
    }


def require_roles(*allowed_roles: str) -> Callable:
    """Dependency kiểm tra phân quyền RBAC (ADMIN, OPS, VIEWER)."""
    def _role_checker(current_user: Dict[str, Any] = Depends(get_current_user)) -> Dict[str, Any]:
        if current_user["role"] not in allowed_roles:
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail=f"Quyền '{current_user['role']}' không được phép thực hiện thao tác này.",
            )
        return current_user

    return _role_checker
