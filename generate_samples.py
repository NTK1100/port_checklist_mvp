import os
import sys
import shutil
import warnings

warnings.filterwarnings("ignore", category=FutureWarning)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

import fitz  # PyMuPDF
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import docx

from database import init_db, get_connection, create_vessel, STORAGE_ROOT

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SAMPLE_DIR = os.path.join(BASE_DIR, "sample_files_for_demo")


def create_pdf_file(filepath: str, title: str, subtitle: str, lines: list[str]) -> None:
    """Tạo file PDF mẫu có định dạng rõ ràng bằng PyMuPDF (fitz), hỗ trợ đầy đủ tiếng Việt Unicode."""
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)  # A4

    win_font = "C:/Windows/Fonts/arial.ttf"
    use_unicode_font = os.path.exists(win_font)
    fname = "arial_uni" if use_unicode_font else "helv"
    ffile = win_font if use_unicode_font else None

    rect = fitz.Rect(36, 36, 559, 115)
    page.draw_rect(rect, color=(0.1, 0.25, 0.5), fill=(0.93, 0.96, 1.0), width=1.5)

    page.insert_text(
        fitz.Point(50, 68),
        title,
        fontsize=15,
        fontname=fname,
        fontfile=ffile,
        color=(0.05, 0.18, 0.4),
    )
    page.insert_text(
        fitz.Point(50, 92),
        subtitle,
        fontsize=10,
        fontname=fname,
        fontfile=ffile,
        color=(0.25, 0.35, 0.5),
    )

    y = 145
    for line in lines:
        if line.startswith("==="):
            page.draw_line(fitz.Point(45, y - 6), fitz.Point(550, y - 6), color=(0.7, 0.75, 0.8), width=0.8)
            y += 8
            continue
        is_bold_header = line.startswith("[") or line.startswith("Vessel Name:") or line.startswith("Total B/Ls:")
        fsize = 11.0 if is_bold_header else 10.0
        color = (0.05, 0.1, 0.25) if is_bold_header else (0.15, 0.15, 0.15)
        page.insert_text(
            fitz.Point(48, y),
            line,
            fontsize=fsize,
            fontname=fname,
            fontfile=ffile,
            color=color,
        )
        y += 20

    page.draw_line(fitz.Point(36, 790), fitz.Point(559, 790), color=(0.8, 0.8, 0.8), width=0.5)
    page.insert_text(
        fitz.Point(48, 808),
        "PortCall Docs - Maritime Document Verification System",
        fontsize=8.5,
        fontname=fname,
        fontfile=ffile,
        color=(0.45, 0.45, 0.45),
    )

    doc.save(filepath)
    doc.close()


def create_excel_vgm_file(filepath: str) -> None:
    """
    Tạo file Excel (.xlsx) mẫu: file_gui_sep.xlsx
    Nội dung là VGM Declaration của tàu MV BINH HOA, chứa 3 mã (BL-101, BL-102, BL-103).
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "VGM_Declaration"

    header_font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    sub_font = Font(name="Calibri", size=11, bold=True, color="1E293B")
    tbl_header_font = Font(name="Calibri", size=11, bold=True, color="0F172A")
    tbl_header_fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    ws["A1"] = "VERIFIED GROSS MASS - VGM DECLARATION"
    ws["A1"].font = header_font
    ws["A1"].fill = header_fill
    ws.merge_cells("A1:F1")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    ws["A2"] = "Vessel Name: MV BINH HOA"
    ws["A2"].font = sub_font

    ws["A3"] = "Port of Loading: SINGAPORE (SGSIN)"
    ws["D3"] = "Port of Discharge: HAI PHONG PORT (VNHPH)"

    ws["A4"] = "SOLAS Chapter VI, Regulation 2 - Batch 1 (Partial VGM Submission: 3 of 5 B/Ls)"

    headers = ["No.", "B/L Number", "Container No.", "Size/Type", "Verified Gross Mass (KGS)", "Weighing Method"]
    for col_idx, text in enumerate(headers, start=1):
        cell = ws.cell(row=6, column=col_idx, value=text)
        cell.font = tbl_header_font
        cell.fill = tbl_header_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center")

    rows = [
        [1, "BL-101", "CSLU1029381", "40HC", 28450, "Method 1 (SOLAS)"],
        [2, "BL-102", "MSCU4928172", "40HC", 26120, "Method 1 (SOLAS)"],
        [3, "BL-103", "OOLU8374653", "20GP", 19800, "Method 2 (SOLAS)"],
    ]

    for r_idx, r_data in enumerate(rows, start=7):
        for c_idx, val in enumerate(r_data, start=1):
            cell = ws.cell(row_idx if False else r_idx, column=c_idx, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center" if c_idx != 6 else "left")

    for col in ["A", "B", "C", "D", "E", "F"]:
        ws.column_dimensions[col].width = 22

    wb.save(filepath)
    wb.close()


def create_docx_crew_file(filepath: str) -> None:
    """
    Tạo file Word (.docx) mẫu: crew_list_word.docx
    Nội dung là IMO FAL FORM 5 - CREW LIST của tàu MV BINH HOA (có cả đoạn văn và bảng biểu).
    """
    doc = docx.Document()
    doc.add_heading("IMO FAL FORM 5 - CREW LIST (DANH SÁCH THUYỀN VIÊN)", level=1)
    doc.add_paragraph("Vessel Name: MV BINH HOA")
    doc.add_paragraph("Nationality of Ship: VIETNAM | IMO Number: 9482716 | Call Sign: 3WAB9")

    table = doc.add_table(rows=1, cols=4)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "No."
    hdr_cells[1].text = "Full Name"
    hdr_cells[2].text = "Rank / Rating"
    hdr_cells[3].text = "Passport / Seaman Book"

    crew_data = [
        ("1", "CAPT. NGUYEN VAN MINH", "Master", "B8821901"),
        ("2", "TRAN DUC THANG", "Chief Officer", "B7719202"),
        ("3", "LE HOANG NAM", "Chief Engineer", "B6610293"),
        ("4", "PHAM MINH TUAN", "Second Officer", "B5591824"),
        ("5", "VU DINH KHOA", "Bosun", "B4482715"),
    ]
    for no, name, rank, pp in crew_data:
        row_cells = table.add_row().cells
        row_cells[0].text = no
        row_cells[1].text = name
        row_cells[2].text = rank
        row_cells[3].text = pp

    doc.add_paragraph("I hereby declare that this CREW LIST (IMO FAL FORM 5) is true, complete and correct.")
    doc.save(filepath)


def generate_all_sample_files() -> list[str]:
    """Sinh đầy đủ bộ file PDF, Excel (.xlsx) & Word (.docx) mẫu trong ./sample_files_for_demo/."""
    os.makedirs(SAMPLE_DIR, exist_ok=True)
    generated = []

    # 1. scan_001.pdf: Cargo Manifest của tàu MV BINH HOA, Total B/Ls: 5 (BL-101 đến BL-105)
    f1 = os.path.join(SAMPLE_DIR, "scan_001.pdf")
    create_pdf_file(
        f1,
        title="INWARD CARGO MANIFEST - BẢN LƯỢC KHAI HÀNG HÓA",
        subtitle="IMO FAL Form 2 | International Container Shipping Manifest",
        lines=[
            "Vessel Name: MV BINH HOA",
            "Port of Loading: SINGAPORE (SGSIN)    |    Port of Discharge: CAT LAI / HAI PHONG",
            "Total B/Ls: 5",
            "===",
            "[MANIFEST SUMMARY - LIST OF BILLS OF LADING]",
            "1. B/L No: BL-101 | Container: CSLU1029381 (40HC) | Cargo: Industrial Machinery | Gross Wt: 28,450 KGS",
            "2. B/L No: BL-102 | Container: MSCU4928172 (40HC) | Cargo: Electronic Components | Gross Wt: 26,120 KGS",
            "3. B/L No: BL-103 | Container: OOLU8374653 (20GP) | Cargo: Plastic Resins (PP)  | Gross Wt: 19,800 KGS",
            "4. B/L No: BL-104 | Container: MAEU7748294 (40HC) | Cargo: Automotive Parts     | Gross Wt: 24,600 KGS",
            "5. B/L No: BL-105 | Container: TGHU9918275 (20GP) | Cargo: Steel Coils          | Gross Wt: 21,900 KGS",
            "===",
            "Declaration: We hereby certify that this CARGO MANIFEST contains 5 Bills of Lading (BL-101 to BL-105).",
            "Master / Authorized Shipping Agent Signature: CAPT. NGUYEN VAN MINH",
        ],
    )
    generated.append(f1)

    # 2. file_gui_sep.xlsx: VGM Declaration của tàu MV BINH HOA, chứa 3 mã (BL-101, BL-102, BL-103)
    f2 = os.path.join(SAMPLE_DIR, "file_gui_sep.xlsx")
    create_excel_vgm_file(f2)
    generated.append(f2)

    # 3. vgm_bo_sung.pdf: VGM Declaration của tàu MV BINH HOA chứa 2 mã còn lại (BL-104, BL-105)
    f3 = os.path.join(SAMPLE_DIR, "vgm_bo_sung.pdf")
    create_pdf_file(
        f3,
        title="VERIFIED GROSS MASS (VGM DECLARATION) - SUPPLEMENT",
        subtitle="SOLAS Chapter VI Regulation 2 | Supplemental VGM Certificate (Batch 2)",
        lines=[
            "Vessel Name: MV BINH HOA",
            "Document Type: VERIFIED GROSS MASS (VGM DECLARATION)",
            "===",
            "[SUPPLEMENTAL VGM LIST - REMAINING 2 SHIPMENTS]",
            "4. B/L No: BL-104 | Container: MAEU7748294 | Size: 40HC | VGM: 24,600 KGS | Method 1",
            "5. B/L No: BL-105 | Container: TGHU9918275 | Size: 20GP | VGM: 21,900 KGS | Method 1",
            "===",
            "Note: Supplemental VGM submission for BL-104 and BL-105 to complete 5/5 B/Ls on MV BINH HOA.",
            "Authorized Signatory: TRAN QUOC BAO - Quality & Safety Dept.",
        ],
    )
    generated.append(f3)

    # 4. doc_bl_full.pdf: Bill of Lading của tàu MV BINH HOA chứa đủ 5 mã (BL-101 đến BL-105)
    f4 = os.path.join(SAMPLE_DIR, "doc_bl_full.pdf")
    create_pdf_file(
        f4,
        title="OCEAN BILL OF LADING - CONSOLIDATED SET",
        subtitle="Combined Transport BILL OF LADING | Original Negotiable Set",
        lines=[
            "Vessel Name: MV BINH HOA",
            "Port of Loading: SINGAPORE    |    Port of Discharge: VIETNAM PORT",
            "===",
            "[CONSOLIDATED OCEAN BILL OF LADING - 5 SETS INCLUDED]",
            "• Master Bill Ref: BL-101 | Container: CSLU1029381 | Consignee: VIET TECH CORP",
            "• Master Bill Ref: BL-102 | Container: MSCU4928172 | Consignee: HANOI ELECTRONICS JSC",
            "• Master Bill Ref: BL-103 | Container: OOLU8374653 | Consignee: MEKONG PLASTICS CO., LTD",
            "• Master Bill Ref: BL-104 | Container: MAEU7748294 | Consignee: AUTO PARTS VIETNAM LTD",
            "• Master Bill Ref: BL-105 | Container: TGHU9918275 | Consignee: LOTUS STEEL MANUFACTURING",
            "===",
            "Total Bills of Lading in this consolidated document: 5 (BL-101, BL-102, BL-103, BL-104, BL-105).",
            "Issued by: PACIFIC OCEAN SHIPPING LINE",
        ],
    )
    generated.append(f4)

    # 5. crew_scan.pdf: IMO FAL Form 5 - CREW LIST của tàu MV BINH HOA
    f5 = os.path.join(SAMPLE_DIR, "crew_scan.pdf")
    create_pdf_file(
        f5,
        title="IMO FAL FORM 5 - CREW LIST (DANH SÁCH THUYỀN VIÊN)",
        subtitle="Convention on Facilitation of International Maritime Traffic",
        lines=[
            "Vessel Name: MV BINH HOA",
            "Nationality of Ship: VIETNAM    |    IMO Number: 9482716    |    Call Sign: 3WAB9",
            "===",
            "[CREW LIST ON ARRIVAL - TOTAL 18 CREW MEMBERS]",
            "1. CAPT. NGUYEN VAN MINH      | Rank: Master           | Nationality: Vietnamese | Passport: B8821901",
            "2. TRAN DUC THANG             | Rank: Chief Officer    | Nationality: Vietnamese | Passport: B7719202",
            "3. LE HOANG NAM               | Rank: Chief Engineer   | Nationality: Vietnamese | Passport: B6610293",
            "4. PHAM MINH TUAN             | Rank: Second Officer   | Nationality: Vietnamese | Passport: B5591824",
            "5. VU DINH KHOA               | Rank: Bosun            | Nationality: Vietnamese | Passport: B4482715",
            "===",
            "I hereby declare that this CREW LIST (IMO FAL FORM 5) is true, complete and correct.",
        ],
    )
    generated.append(f5)

    # 6. crew_list_word.docx: File Word (.docx) mẫu của tàu MV BINH HOA
    f6_docx = os.path.join(SAMPLE_DIR, "crew_list_word.docx")
    create_docx_crew_file(f6_docx)
    generated.append(f6_docx)

    # 7. port_clearance_99.pdf: PORT CLEARANCE của tàu MV BINH HOA
    f7 = os.path.join(SAMPLE_DIR, "port_clearance_99.pdf")
    create_pdf_file(
        f7,
        title="PORT CLEARANCE CERTIFICATE - GIẤY PHÉP RỜI CẢNG CUỐI",
        subtitle="Maritime Port Authority Official Departure Clearance",
        lines=[
            "Vessel Name: MV BINH HOA",
            "Flag: VIETNAM    |    Gross Tonnage: 18,450 GT    |    Net Tonnage: 9,820 NT",
            "Last Port of Call: SINGAPORE    |    Next Port of Call: HAI PHONG, VIETNAM",
            "===",
            "[OFFICIAL PORT CLEARANCE STATEMENT]",
            "This is to certify that the vessel MV BINH HOA has complied with all",
            "port regulations, customs, quarantine, and safety formalities at the Port of Departure.",
            "Permission is hereby granted for the vessel to proceed to its next port of call.",
            "===",
            "Document Type: PORT CLEARANCE / GIẤY PHÉP RỜI CẢNG",
            "Issued by: Harbour Master / Port Authority",
        ],
    )
    generated.append(f7)

    # 8. file_nham_tau.pdf: VGM Declaration nhưng ghi Vessel Name: MV EVER GIVEN
    f8 = os.path.join(SAMPLE_DIR, "file_nham_tau.pdf")
    create_pdf_file(
        f8,
        title="VERIFIED GROSS MASS - VGM DECLARATION",
        subtitle="SOLAS Chapter VI, Regulation 2 | Container Weight Verification",
        lines=[
            "Vessel Name: MV EVER GIVEN",
            "Port of Loading: KAOHSIUNG    |    Port of Discharge: ROTTERDAM",
            "===",
            "[VGM CONTAINER LIST - WRONG VESSEL DEMO FILE]",
            "1. B/L No: BL-901 | Container: EMCU3829101 | Size: 40HC | VGM: 29,100 KGS",
            "2. B/L No: BL-902 | Container: EMCU3829102 | Size: 40HC | VGM: 27,850 KGS",
            "===",
            "WARNING TEST: This document belongs to Vessel Name: MV EVER GIVEN (NOT MV BINH HOA).",
        ],
    )
    generated.append(f8)

    return generated


def seed_initial_vessels(reset_existing: bool = True) -> None:
    """
    Khởi tạo DB và tạo sẵn 2 con tàu mẫu (Không dùng mã chuyến V01/V02):
    1. MV BINH HOA (Tàu Container Quốc tế)
    2. MV HAI AN (Tàu Nội địa)
    """
    init_db()
    conn = get_connection()
    cursor = conn.cursor()

    if reset_existing:
        cursor.execute("DELETE FROM uploaded_files")
        cursor.execute("DELETE FROM checklist_items")
        cursor.execute("DELETE FROM vessels")
        conn.commit()
        if os.path.exists(STORAGE_ROOT):
            for entry in os.listdir(STORAGE_ROOT):
                full_p = os.path.join(STORAGE_ROOT, entry)
                if os.path.isdir(full_p):
                    shutil.rmtree(full_p, ignore_errors=True)
    else:
        cursor.execute("SELECT COUNT(*) AS cnt FROM vessels")
        if cursor.fetchone()["cnt"] > 0:
            conn.close()
            return

    conn.close()

    today_str = "2026-10-26"
    create_vessel(
        vessel_name="MV BINH HOA",
        eta=f"{today_str} 08:30",
        vessel_type="CONTAINER_INTL",
        eta_date_override=today_str,
    )
    create_vessel(
        vessel_name="MV HAI AN",
        eta=f"{today_str} 14:00",
        vessel_type="DOMESTIC",
        eta_date_override=today_str,
    )


if __name__ == "__main__":
    print("=== BẮT ĐẦU KHỞI TẠO DỮ LIỆU MẪU & FILE DEMO (QUY TẮC MỚI) ===")
    seed_initial_vessels(reset_existing=True)
    files = generate_all_sample_files()
    # Nạp sẵn bộ chứng từ mẫu lên tàu MV BINH HOA để hiển thị ngay quy tắc đặt tên mới
    from main import simulate_sample_upload, SimulateSampleRequest, list_vessels
    ops_ctx = {"id": 2, "username": "ops_kien", "full_name": "Nguyễn Trung Kiên", "role": "OPS", "is_active": True}
    vessels = list_vessels(current_user=ops_ctx)["vessels"]
    if vessels:
        simulate_sample_upload(
            vessels[0]["id"],
            SimulateSampleRequest(
                filenames=[
                    "scan_001.pdf",
                    "file_gui_sep.xlsx",
                    "vgm_bo_sung.pdf",
                    "doc_bl_full.pdf",
                    "crew_list_word.docx",
                    "port_clearance_99.pdf",
                ]
            ),
            current_user=ops_ctx,
        )
    print("Đã tạo thành công 2 tàu mẫu (MV BINH HOA & MV HAI AN) và áp dụng quy tắc đặt tên file mới.")
    print(f"Đã sinh {len(files)} file mẫu (.pdf, .xlsx, .docx) tại ./sample_files_for_demo/:")
    for f in files:
        print(" -", os.path.basename(f))
    print("=== HOÀN TẤT ===")
