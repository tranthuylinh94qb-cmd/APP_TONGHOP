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

# Tên file template chuẩn trên GitHub
TEMPLATE_FILE = "Template_Luong_NHCSXH.xlsx"

# ==========================================
# 2. HÀM NHỒI DỮ LIỆU VÀO PHÔI GỐC
# ==========================================
@st.cache_data
def xuat_file_tu_template(df_du_lieu, ten_sheet, dong_bat_dau, cot_mapping):
    """
    Hàm mở file phôi, đổ số liệu từ phần mềm vào đúng tọa độ, giữ nguyên 100% định dạng form.
    """
    if not os.path.exists(TEMPLATE_FILE):
        return None # Báo lỗi nếu không có file phôi
        
    # Mở file phôi gốc
    wb = openpyxl.load_workbook(TEMPLATE_FILE)
    
    if ten_sheet in wb.sheetnames:
        ws = wb[ten_sheet]
        
        # Chạy vòng lặp để điền từng nhân sự vào bảng
        for index, row in df_du_lieu.iterrows():
            dong_hien_tai = dong_bat_dau + index
            
            # Quét các cột cần điền theo cấu hình mapping
            for ten_cot_df, chi_so_cot_excel in cot_mapping.items():
                if ten_cot_df in row:
                    ws.cell(row=dong_hien_tai, column=chi_so_cot_excel).value = row[ten_cot_df]
                    
    # Lưu vào bộ nhớ đệm để tạo link tải (không ghi đè làm hỏng file gốc trên máy chủ)
    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# ==========================================
# 3. DỮ LIỆU NHÂN SỰ & THAM SỐ (MÔ PHỎNG)
# ==========================================
if 'luong_toi_thieu' not in st.session_state:
    st.session_state.luong_toi_thieu = 3250000
if 'tien_an_ca' not in st.session_state:
    st.session_state.tien_an_ca = 730000

# Bạn có thể thay bằng pd.read_excel() để đọc danh sách từ file của bạn
data_nhan_su = pd.DataFrame({
    "STT": [1, 2, 3, 4, 5],
    "Họ và tên": ["Nguyễn Tuấn Ngọc", "Trương Minh Tú", "Nguyễn Xuân Hào", "Hoàng Thị Lệ Hương", "Trần Mạnh Linh"],
    "Chức vụ": ["Giám đốc", "Phó Giám đốc", "Phó Giám đốc", "Trưởng kế toán", "Tổ trưởng"],
    "Hệ số lương": [4.25, 3.8, 3.65, 3.4, 3.1]
})

# ==========================================
# 4. GIAO DIỆN CÁC PHÂN HỆ NGHIỆP VỤ
# ==========================================
menu = st.sidebar.radio(
    "CHỌN CHỨC NĂNG NGHIỆP VỤ:",
    ["1. Trạng thái & Tham số", "2. Thanh toán Lương Kỳ I", "3. Quyết toán Lương Kỳ II", "4. Báo cáo 01A/TUL", "5. Báo cáo 01B/TUL"]
)
st.sidebar.divider()

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
    
    # MAPPING: Cấu hình vị trí đổ số liệu vào phôi
    # "Tên cột trong df": Số thứ tự cột trên Excel (1=A, 2=B, 3=C...)
    # Bạn hãy đổi các con số dưới đây cho đúng với vị trí cột trên file phôi của bạn
    mapping_ky1 = {
        "STT": 1, 
        "Họ và tên": 2, 
        "Chức vụ": 3, 
        "Hệ số lương": 4, 
        "Thực lĩnh Kỳ I": 6  # Giả sử cột thực lĩnh là cột F
    }
    
    # dong_bat_dau=8: Nghĩa là người số 1 sẽ được điền vào dòng số 8 trên file Excel
    file_ky1 = xuat_file_tu_template(df_ky1, ten_sheet="Lương kỳ I", dong_bat_dau=8, cot_mapping=mapping_ky1)
    
    if file_ky1:
        st.download_button("📥 Xuất Bảng Lương Kỳ I (Vào Form Chuẩn)", data=file_ky1, file_name="Bang_Luong_Ky_I.xlsx", type="primary")

elif menu == "3. Quyết toán Lương Kỳ II":
    st.subheader("💵 QUYẾT TOÁN LƯƠNG KỲ II")
    df_ky2 = data_nhan_su.copy()
    df_ky2["Thực lĩnh Kỳ II"] = (df_ky2["Hệ số lương"] * st.session_state.luong_toi_thieu * 1.5 + st.session_state.tien_an_ca).astype(int)
    
    st.dataframe(df_ky2.style.format({"Hệ số lương": "{:.2f}", "Thực lĩnh Kỳ II": "{:,.0f}"}), use_container_width=True, hide_index=True)
    
    mapping_ky2 = {"STT": 1, "Họ và tên": 2, "Chức vụ": 3, "Thực lĩnh Kỳ II": 9} # Sửa lại số cho đúng cột
    file_ky2 = xuat_file_tu_template(df_ky2, ten_sheet="Lương kỳ II", dong_bat_dau=8, cot_mapping=mapping_ky2)
    
    if file_ky2:
        st.download_button("📥 Xuất Bảng Lương Kỳ II (Vào Form Chuẩn)", data=file_ky2, file_name="Bang_Luong_Ky_II.xlsx", type="primary")

elif menu == "4. Báo cáo 01A/TUL":
    st.subheader("📑 BÁO CÁO MẪU 01A/TUL")
    df_01a = data_nhan_su.copy()
    df_01a["Tổng quỹ lương"] = (df_01a["Hệ số lương"] * st.session_state.luong_toi_thieu * 1.5).astype(int)
    
    st.dataframe(df_01a.style.format({"Hệ số lương": "{:.2f}", "Tổng quỹ lương": "{:,.0f}"}), use_container_width=True, hide_index=True)
    
    mapping_01a = {"STT": 1, "Họ và tên": 2, "Chức vụ": 3, "Tổng quỹ lương": 4} # Sửa lại số
    file_01a = xuat_file_tu_template(df_01a, ten_sheet="Mẫu 01A_TUL", dong_bat_dau=8, cot_mapping=mapping_01a)
    
    if file_01a:
        st.download_button("📥 Xuất Mẫu 01A/TUL (Vào Form Chuẩn)", data=file_01a, file_name="Mau_01A_TUL.xlsx", type="primary")

elif menu == "5. Báo cáo 01B/TUL":
    st.subheader("📊 BÁO CÁO MẪU 01B/TUL")
    st.info("Khu vực trích xuất dữ liệu tổng hợp chi phí")
    # Tương tự như trên, bạn cấu hình mapping_01b và xuất file.
# Khi người dùng bấm xuất Kỳ I, code nhắm vào Sheet "Phoi_Ky1"
file_ky1 = xuat_file_tu_template(df_ky1, ten_sheet="Phoi_Ky1", dong_bat_dau=8, cot_mapping=mapping_ky1)

# Khi người dùng bấm xuất Mẫu 01A, code tự chuyển sang Sheet "Phoi_01A"
file_01a = xuat_file_tu_template(df_01a, ten_sheet="Phoi_01A", dong_bat_dau=8, cot_mapping=mapping_01a)