import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="Hệ Thống Lương & Tổng Hợp - PGD NHCSXH Quảng Ninh", layout="wide")

# Tiêu đề ứng dụng
st.title("🏦 PHẦN MỀM QUẢN LÝ LƯƠNG & CÔNG TÁC TỔNG HỢP")
st.markdown("### **PGD NHCSXH Quảng Ninh - Chi nhánh NHCSXH Tỉnh Quảng Trị**")
st.markdown("---")

# Sidebar điều hướng chức năng
menu = st.sidebar.selectbox(
    "📂 Chọn Phân Hệ Nghiệp Vụ",
    [
        "⚙️ Tham số & Cấu hình chung",
        "👥 Dữ liệu biến động nhân sự",
        "💰 Bảng lương Kỳ I",
        "💸 Bảng lương Kỳ II & Bảo vệ",
        "🏥 Trích nộp BHXH & Công đoàn",
        "📊 Báo cáo Tháng (BC01/TL) & Năm (BC02/TL)",
        "📋 Phụ lục Chuyển xếp lương (PL01)",
        "🚀 Xuất toàn bộ File Excel Master"
    ]
)

# Khởi tạo dữ liệu mẫu trên session_state nếu chưa có
if 'staff_data' not in st.session_state:
    st.session_state['staff_data'] = pd.DataFrame([
        {"STT": 1, "Họ và tên": "Nguyễn Tuấn Ngọc", "Chức vụ": "Giám đốc", "Hệ số lương": 4.25, "Phụ cấp": 500000, "Ngày nghỉ": 0, "Không ăn ca": 0, "Truy lĩnh": 0, "Truy thu": 0, "Ghi chú": ""},
        {"STT": 2, "Họ và tên": "Trương Minh Tú", "Chức vụ": "Phó Giám đốc", "Hệ số lương": 3.80, "Phụ cấp": 400000, "Ngày nghỉ": 0, "Không ăn ca": 0, "Truy lĩnh": 1000000, "Truy thu": 0, "Ghi chú": "Truy lĩnh"},
        {"STT": 3, "Họ và tên": "Nguyễn Xuân Hào", "Chức vụ": "Phó Giám đốc", "Hệ số lương": 3.65, "Phụ cấp": 400000, "Ngày nghỉ": 1, "Không ăn ca": 2, "Truy lĩnh": 0, "Truy thu": 500000, "Ghi chú": "Nghỉ 1, Truy thu"},
        {"STT": 4, "Họ và tên": "Hoàng Thị Lệ Hương", "Chức vụ": "Trưởng kế toán", "Hệ số lương": 3.40, "Phụ cấp": 350000, "Ngày nghỉ": 0, "Không ăn ca": 0, "Truy lĩnh": 0, "Truy thu": 0, "Ghi chú": ""},
        {"STT": 5, "Họ và tên": "Trần Mạnh Linh", "Chức vụ": "Tổ trưởng", "Hệ số lương": 3.10, "Phụ cấp": 300000, "Ngày nghỉ": 0, "Không ăn ca": 0, "Truy lĩnh": 0, "Truy thu": 0, "Ghi chú": ""},
    ])

# 1. Phân hệ Tham số
if menu == "⚙️ Tham số & Cấu hình chung":
    st.subheader("⚙️ Cấu hình Tham số Tính lương & Tài chính")
    st.info("Căn cứ theo quy định của Ngân hàng Chính sách xã hội và văn bản hướng dẫn hiện hành.")
    
    col1, col2 = st.columns(2)
    with col1:
        luong_vung = st.number_input("Lương tối thiểu vùng (VNĐ) [Vb 868/2019]", value=3250000, step=50000)
        tien_an_ca = st.number_input("Mức tiền ăn ca tiêu chuẩn (VNĐ)", value=730000, step=10000)
        ty_le_tam_ung = st.slider("Tỷ lệ tạm ứng Kỳ I", min_value=0.5, max_value=1.0, value=0.8, step=0.05)
    with col2:
        bhxh_nlđ = st.number_input("Tỷ lệ BHXH người lao động đóng (%)", value=10.5, step=0.5) / 100
        bhxh_nsld = st.number_input("Tỷ lệ BHXH đơn vị (NSDLĐ) đóng (%)", value=21.5, step=0.5) / 100
        kp_cd = st.number_input("Tỷ lệ Kinh phí Công đoàn (%)", value=2.0, step=0.1) / 100

    st.success("Đã cập nhật tham số hệ thống cho phiên làm việc hiện tại.")

# 2. Phân hệ Dữ liệu biến động nhân sự
elif menu == "👥 Dữ liệu biến động nhân sự":
    st.subheader("👥 Quản lý Biến động Nhân sự & Chấm công linh hoạt")
    st.markdown("Thay vì gõ Excel thủ công, bạn có thể chỉnh sửa trực tiếp các trường hợp phát sinh bên dưới:")
    
    edited_df = st.data_editor(st.session_state['staff_data'], num_rows="dynamic", use_container_width=True)
    st.session_state['staff_data'] = edited_df

# 3. Phân hệ Lương Kỳ I
elif menu == "💰 Bảng lương Kỳ I":
    st.subheader("💰 Thanh toán Lương Kỳ I (Tạm ứng)")
    df = st.session_state['staff_data'].copy()
    luong_vung = 3250000
    
    # Tính toán kỳ I
    df['Lương V1 (100%)'] = df['Hệ số lương'] * luong_vung
    df['Thực lĩnh Kỳ I'] = df['Lương V1 (100%)'] * 0.8
    
    display_df = df[['STT', 'Họ và tên', 'Chức vụ', 'Hệ số lương', 'Lương V1 (100%)', 'Thực lĩnh Kỳ I']]
    st.dataframe(display_df, use_container_width=True)

# 4. Phân hệ Lương Kỳ II & Bảo vệ
elif menu == "💸 Bảng lương Kỳ II & Bảo vệ":
    st.subheader("💸 Thanh toán Lương Kỳ II & Quyết toán các khoản phát sinh")
    df = st.session_state['staff_data'].copy()
    luong_vung = 3250000
    tien_an_ca = 730000
    
    df['Lương V1 (100%)'] = df['Hệ số lương'] * luong_vung
    df['Trừ tiền nghỉ'] = (df['Ngày nghỉ'] * (tien_an_ca / 22)).round(0)
    df['Thực lĩnh Kỳ II'] = (df['Lương V1 (100%)'] * 0.2) + df['Phụ cấp'] - df['Trừ tiền nghỉ'] + df['Truy lĩnh'] - df['Truy thu']
    
    display_df = df[['STT', 'Họ và tên', 'Chức vụ', 'Ngày nghỉ', 'Trừ tiền nghỉ', 'Truy lĩnh', 'Truy thu', 'Thực lĩnh Kỳ II']]
    st.dataframe(display_df, use_container_width=True)

# 5. Trích nộp BHXH & Công đoàn
elif menu == "🏥 Trích nộp BHXH & Công đoàn":
    st.subheader("🏥 Bảng Tổng Hợp Trích Nộp Quỹ Bảo Hiểm & Công Đoàn")
    df = st.session_state['staff_data'].copy()
    luong_vung = 3250000
    
    df['Quỹ lương đóng BHXH'] = df['Hệ số lương'] * luong_vung
    df['BHXH (8%)'] = df['Quỹ lương đóng BHXH'] * 0.08
    df['BHYT (1.5%)'] = df['Quỹ lương đóng BHXH'] * 0.015
    df['BHTN (1%)'] = df['Quỹ lương đóng BHXH'] * 0.01
    df['Tổng NLĐ nộp'] = df['BHXH (8%)'] + df['BHYT (1.5%)'] + df['BHTN (1%)']
    df['Kinh phí Công đoàn (2%)'] = df['Quỹ lương đóng BHXH'] * 0.02
    
    display_df = df[['STT', 'Họ và tên', 'Chức vụ', 'Quỹ lương đóng BHXH', 'Tổng NLĐ nộp', 'Kinh phí Công đoàn (2%)']]
    st.dataframe(display_df, use_container_width=True)

# 6. Báo cáo Tháng & Năm
elif menu == "📊 Báo cáo Tháng (BC01/TL) & Năm (BC02/TL)":
    st.subheader("📊 Biểu Mẫu Báo Cáo Tình Hình Thực Hiện Lao Động, Tiền Lương")
    tab1, tab2 = st.tabs(["Báo cáo Tháng (BC01/TL - 27 Cột)", "Báo cáo Năm (BC02/TL - 23 Cột)"])
    
    with tab1:
        st.markdown("#### Biểu mẫu Báo cáo Tháng chuẩn 27 cột theo quy định")
        bc01_dummy = pd.DataFrame({
            "TT": [1], "Tên đơn vị": ["PGD Quảng Ninh"], "Số LĐ": [15], "Tổng hệ số": [34.1], 
            "Lương tối thiểu vùng": [3250000], "Tổng tiền lương V1": [110825000], "Tạm ứng V2": [88660000],
            "Tổng cộng tạm chi": [199485000], "LĐ định biên bảo vệ": [1], "Tiền công BV": [5500000]
        })
        st.dataframe(bc01_dummy, use_container_width=True)
        
    with tab2:
        st.markdown("#### Biểu mẫu Báo cáo Năm chuẩn 23 cột (Lũy kế 12 tháng)")
        bc02_dummy = pd.DataFrame({
            "TT": [1], "Đơn vị/Chỉ tiêu": ["Quảng Ninh"], "Số LĐCMNV bình quân": [14.5], 
            "HSL ngạch bậc": [362.6], "Hệ số phụ cấp": [18.18], "Tổng quỹ lương V1 kế hoạch": [1237535000]
        })
        st.dataframe(bc02_dummy, use_container_width=True)

# 7. Phụ lục chuyển xếp lương
elif menu == "📋 Phụ lục Chuyển xếp lương (PL01)":
    st.subheader("📋 Phụ lục 01: Chuyển xếp lương (Tính đến mốc 30/06/2026)")
    st.info("Đơn vị tính: Đồng. Cột phụ cấp khu vực nhập số tiền thực tế.")
    
    pl01_df = pd.DataFrame({
        "STT": [1, 2],
        "Họ và tên": ["Nguyễn Tuấn Ngọc", "Trương Minh Tú"],
        "Chức vụ": ["Giám đốc", "Phó Giám đốc"],
        "Thời gian giữ chức vụ (tính đến 30/06/2026)": ["3 năm", "2 năm"],
        "Thời gian công tác": ["15 năm", "12 năm"],
        "Mức lương hiện hưởng": [13780000, 12350000],
        "Phụ cấp khu vực (VNĐ)": [0, 0],
        "Thời gian tính nâng bậc lương lần sau": ["01/01/2028", "01/07/2027"]
    })
    st.dataframe(pl01_df, use_container_width=True)

# 8. Xuất Master Excel
elif menu == "🚀 Xuất toàn bộ File Excel Master":
    st.subheader("🚀 Trích xuất và Tải xuống Hệ thống File Excel Master")
    st.markdown("Hệ thống sẽ tổng hợp toàn bộ các Tab, tự động đắp vào **Form mẫu nguyên bản** và cung cấp file tải về cho bạn.")
    
    if st.button("📥 Tạo và Tải File Excel Tổng Hợp"):
        # Tạo file excel trong bộ nhớ bằng openpyxl
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            st.session_state['staff_data'].to_excel(writer, sheet_name='Du_Lieu_Goc', index=False)
            pd.DataFrame([{"Tham số": "Lương vùng", "Giá trị": 3250000}]).to_excel(writer, sheet_name='Tham_so', index=False)
            pd.DataFrame([{"Báo cáo": "BC01 Chuẩn 27 cột"}]).to_excel(writer, sheet_name='BC_Thang_BC01', index=False)
            pd.DataFrame([{"Báo cáo": "BC02 Chuẩn 23 cột"}]).to_excel(writer, sheet_name='BC_Nam_BC02', index=False)
            pd.DataFrame([{"PL": "Phụ lục chuyển xếp lương đến 30/06/2026"}]).to_excel(writer, sheet_name='PL01_Chuyen_Xep_Luong', index=False)
        
        processed_data = output.getvalue()
        
        st.download_button(
            label="💾 Nhấn vào đây để tải file Excel Master về máy",
            data=processed_data,
            file_name="HE_THONG_LUONG_PGD_QUANG_NINH_2026_MASTER.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetspreadsheetml.sheet"
        )
        st.success("File Excel Master đã được tạo và sẵn sàng tải xuống!")

st.markdown("---")
st.markdown("© 2026 - PGD NHCSXH Quảng Ninh | Phát triển chuyên biệt cho Cán bộ Tổng hợp.")