
import re
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


st.set_page_config(
    page_title="📊 Tool Điểm Danh",
    layout="wide"
)
st.title("📊 Tool Điểm Danh - Xuất Excel")


# ==================================================
# 1. CÁC KIỂU HIỂN THỊ TÊN
# ==================================================
CAC_KIEU_TEN = [
    "Đầy đủ",
    "Rút gọn họ",
    "Rút gọn toàn bộ"
]


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

    if kieu_ten == "Rút gọn họ":
        # Nguyễn Văn An -> N.Văn An
        # Nguyễn Thị Minh Khai -> N.T.Minh Khai
        if len(cac_tu) >= 4:
            phan_rut_gon = "".join(
                tu[0].upper() + "."
                for tu in cac_tu[:-2]
            )
            return phan_rut_gon + " ".join(cac_tu[-2:])

        if len(cac_tu) == 3:
            return cac_tu[0][0].upper() + "." + " ".join(cac_tu[1:])

        if len(cac_tu) == 2:
            return cac_tu[0][0].upper() + "." + cac_tu[1]

    return " ".join(cac_tu)


# ==================================================
# 2. ĐỌC VÀ XỬ LÝ FILE ĐIỂM DANH
# ==================================================
def process_file(file, buoi=None, kieu_ten="Đầy đủ"):
    file.seek(0)

    df = pd.read_excel(file, skiprows=5)

    if df.shape[1] <= 4:
        raise ValueError(
            "File không đủ cột dữ liệu sau khi bỏ 4 cột cuối."
        )

    df = df.iloc[:, :-4].copy()

    if "Lớp" not in df.columns:
        raise ValueError("Không tìm thấy cột 'Lớp' trong file.")

    if "Họ và tên" not in df.columns:
        raise ValueError("Không tìm thấy cột 'Họ và tên' trong file.")

    # Giữ nguyên thứ tự ngày theo file gốc
    ngay_list = list(df.columns[4:])

    # Bỏ các dòng không có lớp
    df = df[df["Lớp"].notna()].copy()

    # Chuẩn hóa tên lớp để hiển thị
    df["Lớp"] = df["Lớp"].astype(str).str.strip()

    rows = []

    # Duyệt ngày trước để bảo toàn thứ tự ngày gốc
    for ngay in ngay_list:
        for lop, group in df.groupby("Lớp", sort=False):
            danh_sach = []

            # Vắng có phép
            vang_p = group.loc[
                group[ngay].astype(str).str.strip().str.upper() == "P",
                "Họ và tên"
            ].tolist()

            # Vắng không phép
            vang_k = group.loc[
                group[ngay].astype(str).str.strip().str.upper() == "K",
                "Họ và tên"
            ].tolist()

            vang_p = [
                xu_ly_ten(ten, kieu_ten)
                for ten in vang_p
            ]

            vang_k = [
                xu_ly_ten(ten, kieu_ten)
                for ten in vang_k
            ]

            danh_sach.extend(
                f"{ten} (P)" for ten in vang_p
            )
            danh_sach.extend(
                f"{ten} (K)" for ten in vang_k
            )

            so_vang = len(danh_sach)

            if so_vang == 0:
                ghi_chu = "V0"
            else:
                ghi_chu = (
                    f"V{so_vang:02d}: "
                    + ", ".join(danh_sach)
                )

            if buoi:
                ghi_chu = f"{buoi} {ghi_chu}"

            rows.append({
                "Lớp": lop,
                "Ngày": ngay,
                "Thống kê": ghi_chu
            })

    return pd.DataFrame(
        rows,
        columns=["Lớp", "Ngày", "Thống kê"]
    ), ngay_list


# ==================================================
# 3. GỘP DỮ LIỆU SÁNG + CHIỀU
# ==================================================
def gop_sang_chieu(
    df_sang,
    df_chieu,
    ngay_sang,
    ngay_chieu
):
    # Ưu tiên thứ tự ngày của file sáng,
    # bổ sung các ngày chỉ có trong file chiều.
    thu_tu_ngay = list(ngay_sang)

    for ngay in ngay_chieu:
        if ngay not in thu_tu_ngay:
            thu_tu_ngay.append(ngay)

    combined = pd.concat(
        [df_sang, df_chieu],
        ignore_index=True
    )

    # Gộp hai buổi vào cùng một ô, xuống dòng riêng.
    merged = (
        combined.groupby(
            ["Lớp", "Ngày"],
            sort=False,
            dropna=False
        )["Thống kê"]
        .agg(lambda x: "\n".join(x.astype(str)))
        .reset_index()
    )

    # Sắp xếp lại theo thứ tự ngày của file gốc
    thu_tu_lop = list(dict.fromkeys(combined["Lớp"].tolist()))

    thu_tu_ngay_map = {
        ngay: i for i, ngay in enumerate(thu_tu_ngay)
    }
    thu_tu_lop_map = {
        lop: i for i, lop in enumerate(thu_tu_lop)
    }

    merged["_thu_tu_ngay"] = merged["Ngày"].map(
        thu_tu_ngay_map
    )
    merged["_thu_tu_lop"] = merged["Lớp"].map(
        thu_tu_lop_map
    )

    merged = merged.sort_values(
        ["_thu_tu_ngay", "_thu_tu_lop"],
        kind="stable"
    ).drop(
        columns=["_thu_tu_ngay", "_thu_tu_lop"]
    ).reset_index(drop=True)

    return merged, thu_tu_ngay


# ==================================================
# 4. TẠO TÊN SHEET AN TOÀN
# ==================================================
def tao_ten_sheet(ten, da_dung):
    ten = str(ten).strip()

    # Excel không cho phép các ký tự sau trong tên sheet
    ten = re.sub(r'[\[\]:*?/\\]', "_", ten)

    ten = ten.strip("'").strip()

    if not ten:
        ten = "Chua_dat_ten"

    ten = ten[:31]

    ten_goc = ten
    dem = 1

    while ten.casefold() in da_dung:
        hau_to = f"_{dem}"
        ten = ten_goc[:31 - len(hau_to)] + hau_to
        dem += 1

    da_dung.add(ten.casefold())

    return ten


# ==================================================
# 5. ĐỊNH DẠNG MỘT SHEET EXCEL
# ==================================================
def dinh_dang_sheet(ws, df):
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False

    # Tiêu đề
    for cell in ws[1]:
        cell.font = Font(
            bold=True,
            color="FFFFFF"
        )
        cell.fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78"
        )
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

    ws.row_dimensions[1].height = 28

    # Độ rộng cột cố định, không bị rộng theo số ngày
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 65

    # Định dạng nội dung
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(
                horizontal="left",
                vertical="top",
                wrap_text=True
            )

        # Căn giữa lớp và ngày
        row[0].alignment = Alignment(
            horizontal="center",
            vertical="top"
        )
        row[1].alignment = Alignment(
            horizontal="center",
            vertical="top"
        )

        # Tự tăng chiều cao dòng theo số dòng văn bản
        noi_dung = str(row[2].value or "")
        so_dong = noi_dung.count("\n") + 1

        # Ước lượng xuống dòng trong ô thống kê
        so_dong += sum(
            max(0, len(dong) - 90) // 65
            for dong in noi_dung.split("\n")
        )

        ws.row_dimensions[row[0].row].height = min(
            max(22, so_dong * 18),
            180
        )

        # Định dạng ngày nếu dữ liệu là ngày tháng
        if hasattr(row[1].value, "strftime"):
            row[1].number_format = "dd/mm/yyyy"

    # Bộ lọc tự động
    if ws.max_row >= 2:
        ws.auto_filter.ref = ws.dimensions


# ==================================================
# 6. XUẤT EXCEL: SHEET TỔNG HỢP + SHEET TỪNG LỚP
# ==================================================
def export_excel(df):
    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        # Sheet tổng hợp
        df.to_excel(
            writer,
            sheet_name="Tổng hợp",
            index=False
        )

        # Mỗi lớp một sheet riêng
        da_dung = {"tổng hợp"}

        for lop in df["Lớp"].drop_duplicates().tolist():
            ten_sheet = tao_ten_sheet(lop, da_dung)

            df_lop = df[df["Lớp"] == lop].copy()

            df_lop.to_excel(
                writer,
                sheet_name=ten_sheet,
                index=False
            )

    output.seek(0)

    wb = load_workbook(output)

    for ws in wb.worksheets:
        df_sheet = pd.DataFrame(
            ws.iter_rows(
                min_row=2,
                values_only=True
            ),
            columns=[
                cell.value for cell in ws[1]
            ]
        )

        dinh_dang_sheet(ws, df_sheet)

    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    return final_output


# ==================================================
# 7. GIAO DIỆN STREAMLIT
# ==================================================
tab1, tab2 = st.tabs([
    "📂 Xử lý 1 file",
    "📂 Xử lý 2 file (Sáng + Chiều)"
])


# ==================================================
# TAB 1: XỬ LÝ MỘT FILE
# ==================================================
with tab1:
    st.subheader("Xử lý 1 file Excel")

    file1 = st.file_uploader(
        "📂 Tải file Excel",
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

    if file1 and st.button(
        "Xử lý và Tải xuống",
        key="btn1"
    ):
        try:
            df1, thu_tu_ngay = process_file(
                file1,
                kieu_ten=kieu_ten_1
            )

            st.session_state["ketqua_1"] = export_excel(df1)

            st.success(
                f"Đã xử lý {len(df1)} dòng thống kê."
            )

        except Exception as e:
            st.error(f"Lỗi khi xử lý file: {e}")

    if st.session_state.get("ketqua_1") is not None:
        st.download_button(
            label="📥 Tải Excel (Tổng hợp + từng lớp)",
            data=st.session_state["ketqua_1"],
            file_name="ketqua_diem_danh.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl1"
        )


# ==================================================
# TAB 2: XỬ LÝ HAI FILE SÁNG + CHIỀU
# ==================================================
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

    if st.button(
        "Xử lý và Tải xuống",
        key="btn2"
    ):
        if not morning_file or not afternoon_file:
            st.error(
                "⚠️ Bạn cần tải lên đủ 2 file Sáng và Chiều."
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

                df_gop, thu_tu_ngay = gop_sang_chieu(
                    df_morning,
                    df_afternoon,
                    ngay_sang,
                    ngay_chieu
                )

                st.session_state["ketqua_2"] = export_excel(
                    df_gop
                )

                st.success(
                    f"Đã gộp dữ liệu sáng và chiều: "
                    f"{len(df_gop)} dòng thống kê."
                )

            except Exception as e:
                st.error(f"Lỗi khi xử lý file: {e}")

    if st.session_state.get("ketqua_2") is not None:
        st.download_button(
            label="📥 Tải Excel (Tổng hợp + từng lớp)",
            data=st.session_state["ketqua_2"],
            file_name="ketqua_diem_danh_sang_chieu.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl2"
        )
