"""
Script tiện ích cấu hình & kiểm tra kết nối Google Drive cho PortCall Docs:
1. Chế độ mặc định (Khuyên dùng - Miễn phí 100%, không cần Google Cloud Console):
   Cấu hình GDRIVE_BRIDGE_URL (Google Apps Script Web App) và BRIDGE_SECRET_KEY vào file .env
   và kiểm tra kết nối trực tiếp tới thư mục PortCall_Checklist_Storage trên Google Drive.
"""
import os
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from gdrive_bridge_client import set_bridge_config, ping_bridge, sync_db_to_gdrive


def main() -> None:
    print("=================================================================")
    print(" CẤU HÌNH GOOGLE DRIVE BRIDGE (APPS SCRIPT WEB APP) - PORTCALL DOCS")
    print("=================================================================")
    print("Hướng dẫn: Copy mã nguồn trong file 'google_drive_bridge.gs' dán vào")
    print("https://script.google.com -> Deploy -> Web app (Execute as: Me, Anyone)")
    print("-----------------------------------------------------------------")
    url = input("1. Dán GDRIVE_BRIDGE_URL (https://script.google.com/macros/s/.../exec): ").strip()
    if not url:
        print("Bạn chưa nhập URL. Giữ nguyên chế độ Local Storage.")
        return

    secret = input("2. Nhập BRIDGE_SECRET_KEY [Mặc định: PortCall_Secret_2026]: ").strip() or "PortCall_Secret_2026"
    set_bridge_config(url, secret)
    print("\nĐang kiểm tra kết nối tới Google Drive Bridge...")
    res = ping_bridge()
    if res.get("ok"):
        print("KẾT NỐI THÀNH CÔNG!")
        print(" - Thư mục gốc trên Drive:", res.get("rootFolderName"))
        print(" - Folder ID:", res.get("rootFolderId"))
        print(" - Link mở thư mục:", res.get("rootFolderUrl"))
        print("\nĐang đồng bộ thử port_checklist.db lên Google Drive...")
        sync_db_to_gdrive()
        print("\nThông tin cấu hình để gắn lên Render.com (Environment Variables):")
        print(f'GDRIVE_BRIDGE_URL="{url}"')
        print(f'BRIDGE_SECRET_KEY="{secret}"')
    else:
        print("Kết nối thất bại:", res.get("error"))


if __name__ == "__main__":
    main()
