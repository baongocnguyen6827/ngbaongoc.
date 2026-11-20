# -*- coding: utf-8 -*-
"""
Bộ sinh Bảng Tính Microsoft Excel (.xlsx) đa trang siêu tốc & chuyên nghiệp
Sử dụng pandas ExcelWriter kết hợp openpyxl để xuất file trong 1-2 giây
Gồm 4 Sheet: Dữ liệu sạch T10, KPI chi nhánh, Nhật ký Audit, Dữ liệu loại trừ
"""

import os
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def generate_excel_workbook(engine, output_path):
    """
    Tạo sổ làm việc Excel (.xlsx) gồm 4 Sheet chuẩn hóa, tối ưu tốc độ cực nhanh
    """
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # 1. Chuẩn bị DataFrame cho Sheet 1: Dữ liệu sạch
    df_clean = engine.cleaned_df if engine.cleaned_df is not None else pd.DataFrame([{"Thong_Bao": "Chưa có dữ liệu làm sạch"}])

    # 2. Chuẩn bị DataFrame cho Sheet 2: KPI chi nhánh
    rep = engine.management_report or {}
    store_reps = rep.get("store_reports", [])
    if store_reps:
        kpi_rows = []
        for idx, s in enumerate(store_reps, 1):
            target_rev = s.get("target_revenue", 0)
            net_rev = s.get("net_revenue", 0)
            pct = round(net_rev / target_rev * 100, 1) if target_rev > 0 else 0
            tier = "P1 (Khẩn cấp)" if pct < 85 else ("P2 (Theo dõi)" if pct < 95 else "P4 (Vượt kế hoạch)")
            kpi_rows.append({
                "STT": idx,
                "Mã Chi Nhánh": s.get("store_id", f"CN-{idx}"),
                "Tên Chi Nhánh": s.get("store_name", f"Chi nhánh {idx}"),
                "Doanh Thu Thuần (VND)": net_rev,
                "Mục Tiêu Kế Hoạch (VND)": target_rev,
                "Tỷ Lệ Đạt (%)": pct,
                "Lợi Nhuận Gộp (VND)": s.get("gross_profit", 0),
                "Biên Lợi Nhuận (%)": s.get("gross_margin", 0),
                "Xếp Hạng Điều Hành": tier
            })
        df_kpi = pd.DataFrame(kpi_rows)
    else:
        df_kpi = pd.DataFrame([{"Thong_Bao": "Chưa có dữ liệu KPI"}])

    # 3. Chuẩn bị DataFrame cho Sheet 3: Nhật ký Audit
    audit_log = engine.cleaning_audit_log or {}
    dedup = audit_log.get("dedup", {}).get("removed_rows", 0)
    period_out = audit_log.get("period_sync", {}).get("out_of_period_rows", 0)
    cancelled = audit_log.get("cancelled_filter", {}).get("cancelled_rows", 0)
    returned = audit_log.get("returns_handling", {}).get("returned_rows", 0)
    ts = audit_log.get("timestamp", "2026-10-31")

    audit_rows = [
        {"Thời Gian": ts, "Quy Tắc / Thao Tác": "1. RULE-DEDUP", "Chi Tiết Thực Thi": "Khử trùng lặp giao dịch", "Tác Động": f"-{dedup} dòng", "Trạng Thái": "ĐẠT CHUẨN ✓"},
        {"Thời Gian": ts, "Quy Tắc / Thao Tác": "2. RULE-PERIOD-SYNC", "Chi Tiết Thực Thi": "Đồng bộ hóa kỳ giao dịch chuẩn", "Tác Động": f"-{period_out} dòng", "Trạng Thái": "ĐẠT CHUẨN ✓"},
        {"Thời Gian": ts, "Quy Tắc / Thao Tác": "3. RULE-CANCELLED", "Chi Tiết Thực Thi": "Bóc tách đơn hủy ra khỏi doanh thu", "Tác Động": f"-{cancelled} đơn", "Trạng Thái": "ĐẠT CHUẨN ✓"},
        {"Thời Gian": ts, "Quy Tắc / Thao Tác": "4. RULE-RETURNS", "Chi Tiết Thực Thi": "Khấu trừ hàng hoàn trả vào DT thuần", "Tác Động": f"{returned} đơn", "Trạng Thái": "CHUẨN HÓA ✓"},
    ]
    for edit in getattr(engine, "user_custom_edits", []):
        audit_rows.append({
            "Thời Gian": edit.get("timestamp", ts),
            "Quy Tắc / Thao Tác": f"TÙY BIẾN: {edit.get('action', 'Chỉnh sửa')}",
            "Chi Tiết Thực Thi": edit.get("detail", "Thực hiện theo yêu cầu người dùng"),
            "Tác Động": f"{edit.get('rows_affected', 0)} dòng, {edit.get('cols_affected', 0)} cột",
            "Trạng Thái": "ĐÃ ÁP DỤNG ✓"
        })
    df_audit = pd.DataFrame(audit_rows)

    # 4. Chuẩn bị DataFrame cho Sheet 4: Dữ liệu loại trừ
    if engine.raw_df is not None and engine.cleaned_df is not None:
        cleaned_ids = set(engine.cleaned_df["transaction_id"])
        df_excluded = engine.raw_df[~engine.raw_df["transaction_id"].isin(cleaned_ids)]
    else:
        df_excluded = pd.DataFrame([{"Thong_Bao": "Không có dòng loại trừ"}])

    # Ghi siêu tốc vào Excel bằng pandas ExcelWriter
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_clean.to_excel(writer, sheet_name="Du_Lieu_Sach_T10", index=False)
        df_kpi.to_excel(writer, sheet_name="Tong_Hop_KPI_Chi_Nhanh", index=False)
        df_audit.to_excel(writer, sheet_name="Nhat_Ky_Audit_Chuan_Hoa", index=False)
        df_excluded.to_excel(writer, sheet_name="Du_Lieu_Loai_Tru", index=False)

        # Định dạng Header thanh lịch màu Rose-600 cho cả 4 sheets
        header_fill = PatternFill(start_color="E11D48", end_color="E11D48", fill_type="solid")
        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        center_align = Alignment(horizontal="center", vertical="center")

        for s_name in ["Du_Lieu_Sach_T10", "Tong_Hop_KPI_Chi_Nhanh", "Nhat_Ky_Audit_Chuan_Hoa", "Du_Lieu_Loai_Tru"]:
            ws = writer.sheets.get(s_name)
            if ws:
                ws.row_dimensions[1].height = 26
                for col_idx in range(1, ws.max_column + 1):
                    cell = ws.cell(row=1, column=col_idx)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = center_align
                    col_letter = get_column_letter(col_idx)
                    val_str = str(cell.value or "")
                    ws.column_dimensions[col_letter].width = max(len(val_str) + 5, 14)

    return output_path
