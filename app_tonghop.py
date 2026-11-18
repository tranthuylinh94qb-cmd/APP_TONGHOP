import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import os

# ==========================================
# 1. CẤU HÌNH GIAO DIỆN CHUNG
# ==========================================
st.set_page_config(page_title="Lương - PGD Quảng Ninh", layout="wide", page_icon="🏦")
st.markdown('<h2 style="color: #0056b3;">🏦 PHẦN MỀM QUẢN LÝ LƯƠNG & TỔNG HỢP</h2>', unsafe_allow_html=True)
st.markdown('<h5 style="color: #d9534f; margin-bottom: 30px;">PGD NHCSXH Quảng Ninh - Chi nhánh NHCSXH Tỉnh Quảng Trị</h5>', unsafe_allow_html=True)

TEMPLATE_FILE = "Template_Luong_NHCSXH.xlsx"

# ==========================================
# 2. HÀM NHỒI DỮ LIỆU VÀO PHÔI GỐC
# ==========================================
@st.cache_data
def xuat_file_tu_template(df_du_lieu, ten_sheet, dong_bat_dau, cot_mapping):
    # Kiểm tra xem file phôi có tồn tại trên hệ thống không
    if not os.path.exists(TEMPLATE_FILE):
        return None
        
    wb = openpyxl.load_workbook(TEMPLATE_FILE)
    
    # Kiểm tra xem tên Sheet bạn yêu cầu có nằm trong file phôi không
    if ten_sheet not in wb.sheetnames:
        return "SHEET_NOT_FOUND" 
        
    ws = wb[ten_sheet]
    
    # Đổ dữ liệu
    for index, row in df_du_lieu.iterrows():
        dong_hien_tai = dong_bat_dau + index
        for ten_cot_df, chi_so_cot_excel in cot_mapping.items():
            if ten_cot_df in row:
                ws.cell(row=dong_hien_tai, column=chi_so_cot_excel).value = row[ten_cot_df]
                
    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# ==========================================
# 3. KHỞI TẠO THAM SỐ & DỮ LIỆU
# ==========================================
if 'luong_toi_thieu' not in st.session_state:
    st.session_state.luong_toi_thieu = 3250000
if 'tien_an_ca' not in st.session_state:
    st.session_state.tien_an_ca = 730000

data_nhan_su = pd.DataFrame({
    "STT": [1, 2, 3, 4, 5],
    "Họ và tên": ["Nguyễn Tuấn Ngọc", "Trương Minh Tú", "Nguyễn Xuân Hào", "Hoàng Thị Lệ Hương", "Trần Mạnh Linh"],
    "Chức vụ": ["Giám đốc", "Phó Giám đốc", "Phó Giám đốc", "Trưởng kế toán", "Tổ trưởng"],
    "Hệ số lương": [4.25, 3.8, 3.65, 3.4, 3.1]
})

# ==========================================
# 4. MENU ĐIỀU HƯỚNG
# ==========================================
menu = st.sidebar.radio(
    "CHỌN CHỨC NĂNG NGHIỆP VỤ:",
    ["1. Trạng thái & Tham số", "2. Thanh toán Lương Kỳ I", "3. Quyết toán Lương Kỳ II", "4. Báo cáo 01A/TUL", "5. Báo cáo 01B/TUL"]
)
st.sidebar.divider()

# ==========================================
# 5. LOGIC TỪNG PHÂN HỆ NGHIỆP VỤ
# ==========================================

if menu == "1. Trạng thái & Tham số":
    st.subheader("⚙️ THÔNG TIN TRẠNG THÁI HỆ THỐNG")
    if os.path.exists(TEMPLATE_FILE):
        st.success(f"✅ Đã kết nối thành công với phôi chuẩn: `{TEMPLATE_FILE}`. Hệ thống sẵn sàng xuất dữ liệu!")
    else:
        st.error(f"❌ Cảnh báo: Chưa tìm thấy file phôi `{TEMPLATE_FILE}`. Hãy kiểm tra lại kho GitHub!")
        
    st.write("---")
    st.session_state.luong_toi_thieu = st.number_input("Mức lương tối thiểu vùng (VNĐ):", value=st.session_state.luong_toi_thieu, step=50000)
    st.session_state.tien_an_ca = st.number_input("Mức tiền Ăn ca (VNĐ):", value=st.session_state.tien_an_ca, step=10000)

elif menu == "2. Thanh toán Lương Kỳ I":
    st.subheader("💰 TẠM ỨNG LƯƠNG KỲ I")
    df_ky1 = data_nhan_su.copy()
    df_ky1["Thực lĩnh Kỳ I"] = (df_ky1["Hệ số lương"] * st.session_state.luong_toi_thieu * 0.8).astype(int)
    st.dataframe(df_ky1.style.format({"Hệ số lương": "{:.2f}", "Thực lĩnh Kỳ I": "{:,.0f}"}), use_container_width=True, hide_index=True)
    
    mapping_ky1 = {"STT": 1, "Họ và tên": 2, "Chức vụ": 3, "Hệ số lương": 4, "Thực lĩnh Kỳ I": 6}
    
    # Hàm xuất file truyền đúng tên sheet
    file_ky1 = xuat_file_tu_template(df_ky1, ten_sheet="Lương kỳ I", dong_bat_dau=8, cot_mapping=mapping_ky1)
    
    # Cơ chế bẫy lỗi Sheet
    if file_ky1 == "SHEET_NOT_FOUND":
        st.warning("⚠️ Không tìm thấy Sheet có tên 'Lương kỳ I' trong file template gốc. Vui lòng mở file Excel trên máy kiểm tra lại tên Sheet.")
    elif file_ky1:
        st.download_button("📥 Xuất Bảng Lương Kỳ I", data=file_ky1, file_name="Bang_Luong_Ky_I.xlsx", type="primary")

elif menu == "3. Quyết toán Lương Kỳ II":
    st.subheader("💵 QUYẾT TOÁN LƯƠNG KỲ II")
    df_ky2 = data_nhan_su.copy()
    df_ky2["Thực lĩnh Kỳ II"] = (df_ky2["Hệ số lương"] * st.session_state.luong_toi_thieu * 1.5 + st.session_state.tien_an_ca).astype(int)
    st.dataframe(df_ky2.style.format({"Hệ số lương": "{:.2f}", "Thực lĩnh Kỳ II": "{:,.0f}"}), use_container_width=True, hide_index=True)
    
    mapping_ky2 = {"STT": 1, "Họ và tên": 2, "Chức vụ": 3, "Thực lĩnh Kỳ II": 9}
    
    file_ky2 = xuat_file_tu_template(df_ky2, ten_sheet="Lương kỳ II", dong_bat_dau=8, cot_mapping=mapping_ky2)
    
    if file_ky2 == "SHEET_NOT_FOUND":
        st.warning("⚠️ Không tìm thấy Sheet có tên 'Lương kỳ II' trong file template gốc.")
    elif file_ky2:
        st.download_button("📥 Xuất Bảng Lương Kỳ II", data=file_ky2, file_name="Bang_Luong_Ky_II.xlsx", type="primary")

elif menu == "4. Báo cáo 01A/TUL":
    st.subheader("📑 BÁO CÁO MẪU 01A/TUL")
    df_01a = data_nhan_su.copy()
    df_01a["Tổng quỹ lương"] = (df_01a["Hệ số lương"] * st.session_state.luong_toi_thieu * 1.5).astype(int)
    st.dataframe(df_01a.style.format({"Hệ số lương": "{:.2f}", "Tổng quỹ lương": "{:,.0f}"}), use_container_width=True, hide_index=True)
    
    mapping_01a = {"STT": 1, "Họ và tên": 2, "Chức vụ": 3, "Tổng quỹ lương": 4}
    
    file_01a = xuat_file_tu_template(df_01a, ten_sheet="Mẫu 01A_TUL", dong_bat_dau=8, cot_mapping=mapping_01a)
    
    if file_01a == "SHEET_NOT_FOUND":
        st.warning("⚠️ Không tìm thấy Sheet có tên 'Mẫu 01A_TUL' trong file template gốc.")
    elif file_01a:
        st.download_button("📥 Xuất Mẫu 01A/TUL", data=file_01a, file_name="Mau_01A_TUL.xlsx", type="primary")

elif menu == "5. Báo cáo 01B/TUL":
    st.subheader("📊 BÁO CÁO MẪU 01B/TUL")
    st.info("Khu vực trích xuất dữ liệu tổng hợp chi phí (Chờ tích hợp danh sách biểu phí)")
