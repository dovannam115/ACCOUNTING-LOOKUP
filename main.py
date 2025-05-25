import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Excel Lookup Tool", layout="centered")

# CSS cho giao diện
page_bg_img = '''
<style>
body {
    background-image: url("https://novajob.vn/uploads/recruitment/ntd/hoang_gia_luat_logo-1.png");
    background-size: contain;
    background-position: center top;
    background-repeat: no-repeat;
    background-attachment: fixed;
    background-color: #f5f5f5;
}
</style>
'''
st.markdown(page_bg_img, unsafe_allow_html=True)

st.title("🔍 Excel Lookup Tool")

option = st.radio("📌 Chọn chức năng", ["🔁 Lookup Bán ra & NXT", "📄 Lookup theo mapping"])

# --- Chức năng 1: Lookup Bán ra & NXT ---
if option == "🔁 Lookup Bán ra & NXT":
    ban_ra_file = st.file_uploader("📤 Upload file Bán ra", type=["xlsx"], key="ban_ra")
    nxt_t4_file = st.file_uploader("📤 Upload file NXT", type=["xlsx"], key="nxt_t4")

    if ban_ra_file and nxt_t4_file:
        if st.button("🚀 Chạy Lookup"):
            try:
                # Đọc toàn bộ sheet nhưng kiểm tra số cột
                ban_ra_df = pd.read_excel(ban_ra_file, sheet_name="Smart_KTSC_OK")
                if ban_ra_df.shape[1] < 26:
                    st.error(f"Sheet 'Smart_KTSC_OK' chỉ có {ban_ra_df.shape[1]} cột, cần ít nhất 26 cột.")
                    st.stop()

                nxt_t4_df = pd.read_excel(nxt_t4_file, sheet_name="F8_D", skiprows=22)
                if nxt_t4_df.shape[1] < 15:
                    st.error(f"Sheet 'F8_D' chỉ có {nxt_t4_df.shape[1]} cột, cần ít nhất 15 cột.")
                    st.stop()

                # Đặt tên cột
                nxt_t4_df.columns.values[[2, 4, 14]] = ['target_col', 'match_col', 'compare_col']
                q_col = ban_ra_df.columns[16]
                z_col = ban_ra_df.columns[25]

                # Vector hóa lookup
                merged = ban_ra_df.merge(nxt_t4_df, left_on=q_col, right_on='match_col', how='left')
                merged = merged[merged['compare_col'] <= merged[z_col]].copy()
                if merged.empty:
                    st.warning("Không tìm thấy kết quả khớp nào.")
                    ban_ra_df['lookup_result'] = "Không tìm thấy"
                else:
                    merged['diff'] = merged[z_col] - merged['compare_col']
                    result_df = merged.loc[merged.groupby(q_col)['diff'].idxmin(), ['target_col']].reset_index()
                    ban_ra_df['lookup_result'] = ban_ra_df[q_col].map(
                        result_df.set_index(q_col)['target_col']
                    ).fillna("Không tìm thấy")

                # Ghi kết quả
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    ban_ra_df.to_excel(writer, index=False, sheet_name="Smart_KTSC_OK")
                output.seek(0)

                st.success("✅ DONE")
                st.download_button(
                    label="📥 Tải file kết quả",
                    data=output,
                    file_name="BAN_RA_lookup_result.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except ValueError as ve:
                st.error(f"Lỗi: Sheet không tồn tại hoặc không đọc được. Chi tiết: {str(ve)}")
            except Exception as e:
                st.error(f"Lỗi: {str(e)}")

# --- Chức năng 2: Lookup theo mapping ---
elif option == "📄 Lookup Mua vào & NXT":
    data_file = st.file_uploader("📤 Upload file Data", type=["xlsx"], key="data")
    mapping_file = st.file_uploader("📤 Upload file Mapping", type=["xlsx"], key="mapping")

    error_threshold = st.number_input(
        "🔧 Nhập phần trăm sai số cho phép (vd: 3% là 0.03)",
        min_value=0.0, max_value=1.0, value=0.03, step=0.01
    )

    if data_file and mapping_file:
        if st.button("🚀 Chạy Lookup Mapping"):
            try:
                # Đọc toàn bộ sheet nhưng kiểm tra số cột
                data_df = pd.read_excel(data_file)
                mapping_df = pd.read_excel(mapping_file)
                if data_df.shape[1] < 5:
                    st.error(f"File Data chỉ có {data_df.shape[1]} cột, cần ít nhất 5 cột.")
                    st.stop()
                if mapping_df.shape[1] < 7:
                    st.error(f"File Mapping chỉ có {mapping_df.shape[1]} cột, cần ít nhất 7 cột.")
                    st.stop()

                # Đặt tên cột
                data_df.columns.values[[16, 25]] = ['TENDM', 'DGVND']
                mapping_df.columns.values[[2, 4, 14]] = ['target_col', 'match_col', 'compare_col']

                # Hàm làm sạch text
                def clean_text(val):
                    if isinstance(val, str):
                        return val.strip().replace("\xa0", "").replace("\n", "").replace("\r", "")
                    return val

                # Vector hóa lookup
                data_df['TENDM'] = data_df['TENDM'].apply(clean_text)
                mapping_df['match_col'] = mapping_df['match_col'].apply(clean_text)
                mapping_df['compare_col'] = pd.to_numeric(mapping_df['compare_col'], errors='coerce')

                merged = data_df.merge(mapping_df, left_on='TENDM', right_on='match_col', how='left')
                merged['error'] = abs(merged['compare_col'] - merged['DGVND']) / merged['DGVND']
                filtered = merged[
                    (merged['error'] <= error_threshold) &
                    (merged['DGVND'] != 0) &
                    (merged['DGVND'].notnull())
                ]
                result_df = filtered.groupby('TENDM').first()['target_col'].reset_index()
                data_df['lookup_result'] = data_df['TENDM'].map(
                    result_df.set_index('TENDM')['target_col']
                ).fillna("Không tìm thấy")

                # Ghi kết quả
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    data_df.to_excel(writer, index=False, sheet_name="Data_Result")
                output.seek(0)

                st.success("✅ Lookup thành công!")
                st.download_button(
                    label="📥 Tải file kết quả",
                    data=output,
                    file_name="data_lookup_result.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except ValueError as ve:
                st.error(f"Lỗi: Sheet không tồn tại hoặc không đọc được. Chi tiết: {str(ve)}")
            except Exception as e:
                st.error(f"Lỗi: {str(e)}")
