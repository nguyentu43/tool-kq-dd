import streamlit as st
import pandas as pd
import os
from io import BytesIO
from openpyxl.styles import Alignment
from openpyxl import load_workbook

st.set_page_config(page_title="📊 Tool Điểm Danh", layout="wide")
st.title("📊 Tool Điểm Danh - Xuất Excel")

def process_excel(file):
    # Đọc Excel, bỏ 5 dòng đầu
    df = pd.read_excel(file, skiprows=5)
    df = df.iloc[:, :-4]

    rows = []

    # Cột ngày bắt đầu từ cột 5 (index=4)
    for col in df.columns[4:]:
        for lop, group in df.groupby("Lớp"):
            danh_sach = []

            # Vắng có phép (P)
            vang_p = group[group[col] == "P"]["Họ và tên"].tolist()
            danh_sach += [f"{ten} (P)" for ten in vang_p]

            # Vắng không phép (K)
            vang_k = group[group[col] == "K"]["Họ và tên"].tolist()
            danh_sach += [f"{ten} (K)" for ten in vang_k]

            so_vang = len(danh_sach)

            if so_vang == 0:
                ghi_chu = "V0"
            else:
                ghi_chu = f"V{so_vang:02d}: " + ", ".join(danh_sach)

            rows.append({
                "Lớp": lop,
                "Ngày": col,
                "Thống kê": ghi_chu
            })

    # Chuyển thành DataFrame
    summary = pd.DataFrame(rows)

    # Xoay bảng: mỗi ngày thành 1 cột
    pivot = summary.pivot(index="Lớp", columns="Ngày", values="Thống kê").reset_index()

    # Xuất Excel ra memory
    output = BytesIO()
    pivot.to_excel(output, index=False)
    output.seek(0)

    # Mở lại bằng openpyxl để format
    wb = load_workbook(output)
    ws = wb.active

    # Set độ rộng cột
    ws.column_dimensions['A'].width = 12   # cột Lớp
    for col in ws.iter_cols(min_col=2, max_col=ws.max_column):
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = 30

    # Căn chỉnh + xuống dòng
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=1):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=2, max_col=ws.max_column):
        for cell in row:
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

    # Freeze pane tại B2
    ws.freeze_panes = "B2"

    # Lưu lại vào BytesIO
    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    return final_output


# Giao diện Streamlit
uploaded_file = st.file_uploader("📂 Tải file Excel (.xls, .xlsx)", type=["xls", "xlsx"])

if uploaded_file:
    st.success("✅ File đã tải lên. Bấm nút để xử lý.")
    if st.button("Xử lý và Tải xuống"):
        result = process_excel(uploaded_file)
        st.download_button(
            label="📥 Tải file kết quả",
            data=result,
            file_name="ketqua.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
