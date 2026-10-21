
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Alignment
from openpyxl import load_workbook

st.set_page_config(page_title="📊 Tool Điểm Danh", layout="wide")
st.title("📊 Tool Điểm Danh - Xuất Excel")


# ======================
# Hàm xử lý tên học sinh
# ======================
def xu_ly_ten(ho_ten, kieu_ten="Đầy đủ"):
    if pd.isna(ho_ten):
        return ""

    cac_tu = str(ho_ten).strip().split()

    if not cac_tu:
        return ""

    if kieu_ten == "Rút gọn toàn bộ":
        # Nguyễn Văn An -> N.V.An
        return "".join(
            tu[0].upper() + "."
            for tu in cac_tu[:-1]
        ) + cac_tu[-1]

    elif kieu_ten == "Rút gọn họ":
        # Nguyễn Văn An -> N.Văn An
        return cac_tu[0][0].upper() + "." + " ".join(cac_tu[1:])

    # Giữ nguyên họ tên
    return " ".join(cac_tu)


# ======================
# Hàm xử lý 1 file
# ======================
def process_file(file, buoi=None, kieu_ten="Đầy đủ"):
    file.seek(0)
    df = pd.read_excel(file, skiprows=5)
    df = df.iloc[:, :-4]  # Bỏ 4 cột cuối

    # Giữ thứ tự ngày theo file gốc
    ngay_list = list(df.columns[4:])

    rows = []

    for col in ngay_list:
        for lop, group in df.groupby("Lớp", sort=False):
            danh_sach = []

            # Vắng có phép (P)
            vang_p = group[group[col] == "P"]["Họ và tên"].tolist()

            # Vắng không phép (K)
            vang_k = group[group[col] == "K"]["Họ và tên"].tolist()

            # Xử lý tên theo lựa chọn
            vang_p = [
                xu_ly_ten(ten, kieu_ten) for ten in vang_p
            ]
            vang_k = [
                xu_ly_ten(ten, kieu_ten) for ten in vang_k
            ]

            danh_sach += [f"{ten} (P)" for ten in vang_p]
            danh_sach += [f"{ten} (K)" for ten in vang_k]

            so_vang = len(danh_sach)

            if so_vang == 0:
                ghi_chu = "V0"
            else:
                ghi_chu = f"V{so_vang:02d}: " + ", ".join(danh_sach)

            if buoi:
                ghi_chu = f"{buoi} {ghi_chu}"

            rows.append({
                "Lớp": lop,
                "Ngày": col,
                "Thống kê": ghi_chu
            })

    return pd.DataFrame(rows), ngay_list


# ======================
# Xuất Excel có định dạng
# ======================
def export_excel(pivot):
    output = BytesIO()
    pivot.to_excel(output, index=False)
    output.seek(0)

    wb = load_workbook(output)
    ws = wb.active

    ws.column_dimensions["A"].width = 12

    for col in ws.iter_cols(min_col=2, max_col=ws.max_column):
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = 40

    for row in ws.iter_rows(
        min_row=2,
        max_row=ws.max_row,
        min_col=2,
        max_col=ws.max_column
    ):
        for cell in row:
            cell.alignment = Alignment(
                horizontal="left",
                vertical="top",
                wrap_text=True
            )

    ws.freeze_panes = "B2"

    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    return final_output


# ======================
# Các lựa chọn viết tên
# ======================
CAC_KIEU_TEN = [
    "Đầy đủ",
    "Rút gọn họ",
    "Rút gọn toàn bộ"
]


# ======================
# Tabs giao diện
# ======================
tab1, tab2 = st.tabs([
    "📂 Xử lý 1 file",
    "📂 Xử lý 2 file (Sáng + Chiều)"
])


# ======================
# Tab 1: Xử lý 1 file
# ======================
with tab1:
    st.subheader("Xử lý 1 file Excel")

    file1 = st.file_uploader(
        "Tải file Excel",
        type=["xls", "xlsx"],
        key="onefile"
    )

    kieu_ten_1 = st.selectbox(
        "Cách hiển thị tên học sinh",
        CAC_KIEU_TEN,
        format_func=lambda x: {
            "Đầy đủ": "Đầy đủ - Nguyễn Văn An",
            "Rút gọn họ": "Rút gọn họ - N.Văn An",
            "Rút gọn toàn bộ": "Rút gọn toàn bộ - N.V.An"
        }[x],
        key="kieu_ten_1"
    )

    if file1:
        if st.button("Xử lý và Tải xuống", key="btn1"):
            try:
                df1, thu_tu_ngay = process_file(
                    file1,
                    kieu_ten=kieu_ten_1
                )

                pivot1 = df1.pivot(
                    index="Lớp",
                    columns="Ngày",
                    values="Thống kê"
                )

                pivot1 = pivot1.reindex(columns=thu_tu_ngay)
                pivot1 = pivot1.reset_index()
                pivot1.columns.name = None

                result1 = export_excel(pivot1)

                st.download_button(
                    label="📥 Tải file kết quả",
                    data=result1,
                    file_name="ketqua.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl1"
                )

            except Exception as e:
                st.error(f"Lỗi khi xử lý file: {e}")


# ======================
# Tab 2: Xử lý 2 file
# ======================
with tab2:
    st.subheader("Xử lý 2 file Excel (Sáng + Chiều)")

    col1, col2 = st.columns(2)

    with col1:
        morning_file = st.file_uploader(
            "📂 File Buổi Sáng",
            type=["xls", "xlsx"],
            key="morning"
        )

    with col2:
        afternoon_file = st.file_uploader(
            "📂 File Buổi Chiều",
            type=["xls", "xlsx"],
            key="afternoon"
        )

    kieu_ten_2 = st.selectbox(
        "Cách hiển thị tên học sinh",
        CAC_KIEU_TEN,
        format_func=lambda x: {
            "Đầy đủ": "Đầy đủ - Nguyễn Văn An",
            "Rút gọn họ": "Rút gọn họ - N.Văn An",
            "Rút gọn toàn bộ": "Rút gọn toàn bộ - N.V.An"
        }[x],
        key="kieu_ten_2"
    )

    if st.button("Xử lý và Tải xuống", key="btn2"):
        if not morning_file or not afternoon_file:
            st.error(
                "⚠️ Bạn cần tải lên đủ 2 file (Sáng + Chiều) trước khi xử lý."
            )
        else:
            try:
                df_morning, ngay_sang = process_file(
                    morning_file,
                    buoi="Sáng",
                    kieu_ten=kieu_ten_2
                )

                df_afternoon, ngay_chieu = process_file(
                    afternoon_file,
                    buoi="Chiều",
                    kieu_ten=kieu_ten_2
                )

                # Ưu tiên thứ tự ngày của file sáng
                # Bổ sung ngày chỉ có trong file chiều
                thu_tu_ngay = ngay_sang + [
                    ngay for ngay in ngay_chieu
                    if ngay not in ngay_sang
                ]

                combined = pd.concat(
                    [df_morning, df_afternoon],
                    ignore_index=True
                )

                merged = (
                    combined.groupby(
                        ["Lớp", "Ngày"],
                        sort=False
                    )["Thống kê"]
                    .apply(lambda x: "\n".join(x))
                    .reset_index()
                )

                pivot2 = merged.pivot(
                    index="Lớp",
                    columns="Ngày",
                    values="Thống kê"
                )

                pivot2 = pivot2.reindex(columns=thu_tu_ngay)
                pivot2 = pivot2.reset_index()
                pivot2.columns.name = None

                result2 = export_excel(pivot2)

                st.download_button(
                    label="📥 Tải file kết quả (Sáng + Chiều)",
                    data=result2,
                    file_name="ketqua_sang_chieu.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl2"
                )

            except Exception as e:
                st.error(f"Lỗi khi xử lý file: {e}")
