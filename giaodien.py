import streamlit as st
import os

# --- LƯU Ý: KHÔNG ĐƯỢC TỰ Ý THAY ĐỔI GIAO DIỆN NÀY NẾU KHÔNG CÓ YÊU CẦU ---

def khoi_tao_session_state():
    """Khởi tạo trạng thái. Để trống Header để hiện chữ chìm, giữ nguyên chữ ký Footer."""
    if "show_guide" not in st.session_state:
        st.session_state["show_guide"] = False

    defaults = {
        'header_so': '', 
        'header_truong': '', 
        'header_to': '', 
        'header_kythi': '', 
        'header_namhoc': '', 
        'header_mon': '', 
        'header_thoigian': '',
        'footer_gt1': 'Giám thị 1',
        'footer_gt2': 'Giám thị 2'
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val

def hien_thi_sidebar(supabase=None):
    khoi_tao_session_state()
    config = {}

    with st.sidebar:
        # --- CÔNG TẮC HƯỚNG DẪN ---
        st.toggle("📖 Bật/Tắt Cẩm nang Hướng dẫn", key="show_guide")
        st.divider()

        # --- 1. CẤU HÌNH ---
        st.header("1. CẤU HÌNH")
        
        loai_mon_input = st.radio(
            "Chọn Môn Thi:",
            options=["Môn KHTN/KHXH (Toán, Hóa...)", "Môn Tiếng Anh", "Môn Đánh Giá Năng Lực"],
            index=0
        )
        
        if "KHTN" in loai_mon_input:
            config['loai_mon'] = 'MON_KHAC'
        elif "Tiếng Anh" in loai_mon_input:
            config['loai_mon'] = 'ENG'
        else:
            config['loai_mon'] = 'DGNL'

        c1, c2 = st.columns(2)
        with c1:
            config['so_luong_de'] = st.number_input("Số đề:", min_value=1, max_value=99, value=4)
        
        che_do_ma = st.radio("Cách sinh Mã đề:", ["Bắt đầu từ...", "Ngẫu nhiên"])
        if che_do_ma == "Bắt đầu từ...":
            config['ma_de_start'] = st.number_input("Mã bắt đầu:", value=101)
            config['kieu_ma_de'] = 'SEQUENTIAL'
        else:
            kieu_ngau_nhien = st.selectbox("Độ dài mã:", ["3 chữ số (VD: 142)", "4 chữ số (VD: 7924)"])
            config['kieu_ma_de'] = 'RANDOM_4' if "4 chữ số" in kieu_ngau_nhien else 'RANDOM_3'

        st.divider()

        # --- 2. TRÌNH BÀY ĐỀ THI ---
        st.header("2. TRÌNH BÀY ĐỀ THI")
        
        config['co_header'] = st.checkbox("Gắn Tiêu Đề (Sở/Trường...)", value=True, key="luu_tuy_chon_header")
        if config['co_header']:
            st.caption("✨ *Hệ thống tự động thêm chữ 'Sở GD và ĐT', 'Trường', 'Tổ', 'Năm học'. Bạn chỉ cần nhập tên ngắn gọn!*")
            
            so_in = st.text_input("Tên Sở (Chỉ nhập tên Tỉnh/TP):", key="header_so", placeholder="VD: HÀ NỘI hoặc TP HỒ CHÍ MINH")
            truong_in = st.text_input("Tên Trường (Chỉ nhập tên):", key="header_truong", placeholder="VD: CHUYÊN KHTN")
            to_in = st.text_input("Tổ Chuyên Môn (Chỉ nhập tên):", key="header_to", placeholder="VD: TOÁN - TIN")
            kythi_in = st.text_input("Kỳ Thi:", key="header_kythi", placeholder="VD: KIỂM TRA GIỮA HỌC KỲ I")
            namhoc_in = st.text_input("Năm học (Chỉ cần nhập số):", key="header_namhoc", placeholder="VD: 2025 - 2026")
            mon_in = st.text_input("Môn Thi (Chỉ nhập tên môn):", key="header_mon", placeholder="VD: TOÁN 12")
            thoigian_in = st.text_input("Thời gian (Chỉ nhập số phút):", key="header_thoigian", placeholder="VD: 90")
            
            if st.button("💾 Lưu làm mặc định cho tài khoản này", use_container_width=True):
                if supabase and "email" in st.session_state:
                    du_lieu_luu = {
                        'header_so': so_in, 'header_truong': truong_in, 'header_to': to_in,
                        'header_kythi': kythi_in, 'header_namhoc': namhoc_in, 
                        'header_mon': mon_in, 'header_thoigian': thoigian_in
                    }
                    try:
                        supabase.table("users_data").update({"cau_hinh_mac_dinh": du_lieu_luu}).eq("email", st.session_state["email"]).execute()
                        st.success("✅ Đã lưu cấu hình lên Đám mây!")
                    except Exception as e:
                        st.error(f"Lỗi lưu Đám mây: {e}")

            # --- BỘ LỌC CHỐNG LẶP TỪ THÔNG MINH ---
            val_so = f"SỞ GD VÀ ĐT {so_in}" if so_in else ""
            if val_so: val_so = val_so.replace("SỞ GD VÀ ĐT SỞ GD VÀ ĐT", "SỞ GD VÀ ĐT").replace("SỞ GD VÀ ĐT SỞ GD", "SỞ GD VÀ ĐT").replace("SỞ GD VÀ ĐT SỞ", "SỞ GD VÀ ĐT")
                
            val_truong = f"TRƯỜNG THPT {truong_in}" if truong_in else ""
            if val_truong: val_truong = val_truong.replace("TRƯỜNG THPT TRƯỜNG THPT", "TRƯỜNG THPT").replace("TRƯỜNG THPT TRƯỜNG", "TRƯỜNG THPT")
                
            val_to = f"TỔ {to_in}" if to_in else ""
            if val_to: val_to = val_to.replace("TỔ TỔ", "TỔ")
                
            val_namhoc = f"NĂM HỌC: {namhoc_in}" if namhoc_in else ""
            if val_namhoc: val_namhoc = val_namhoc.replace("NĂM HỌC: NĂM HỌC:", "NĂM HỌC:").replace("NĂM HỌC: NĂM HỌC", "NĂM HỌC:")
                
            val_mon = f"MÔN {mon_in}" if mon_in else ""
            if val_mon: val_mon = val_mon.replace("MÔN MÔN", "MÔN")

            # --- ĐỐI CHIẾU CHÌA KHÓA CHÍNH XÁC 100% VỚI XUAT_FILE_WORD.PY ---
            config['header_data'] = {
                'so': val_so,
                'truong': val_truong,
                'to_chuyen_mon': val_to,    # Chìa khóa chính xác
                'ky_thi': kythi_in,         # Chìa khóa chính xác
                'nam_hoc': val_namhoc,      # Chìa khóa chính xác
                'mon': val_mon, 
                'thoi_gian': thoigian_in    # Chìa khóa chính xác
            }

        # --- FOOTER ---
        config['co_footer'] = st.checkbox("Gắn Tiêu Đề Kết Thúc (Chữ ký)", value=True, key="luu_tuy_chon_footer")
        if config['co_footer']:
            c_f1, c_f2 = st.columns(2)
            gt1_in = c_f1.text_input("Chức danh 1:", key='footer_gt1', placeholder="VD: Giám thị 1")
            gt2_in = c_f2.text_input("Chức danh 2:", key='footer_gt2', placeholder="VD: Giám thị 2")
            config['footer_data'] = {
                'gt1': gt1_in,
                'gt2': gt2_in
            }

        st.divider()
        
        # --- CHÈN ẢNH PHIẾU ---
        st.header("3. CHÈN PHIẾU LÀM BÀI")
        with st.expander("Tải lên mẫu Phiếu"):
            config['img_phieu_to'] = st.file_uploader("Phiếu Tô (Trang 1):", type=['png', 'jpg', 'jpeg'], key="img1")
            config['img_tu_luan'] = st.file_uploader("Giấy Tự Luận (Trang 2):", type=['png', 'jpg', 'jpeg'], key="img2")
            config['file_quy_uoc'] = st.file_uploader("File Quy ước môn (.docx):", type=['docx'], key="doc_quyuoc")

        st.divider()

        # --- CÀI ĐẶT TRỘN ---
        st.header("4. CÀI ĐẶT TRỘN")
        config['tron_nhom'] = st.checkbox("Trộn nhóm (Cluster)", value=False, help="Hoán vị vị trí các nhóm câu hỏi")
        
        c_t1, c_t2 = st.columns(2)
        config['tron_mcq'] = c_t1.checkbox("Đảo A,B,C,D", value=True)
        
        disable_ds = (config['loai_mon'] != 'MON_KHAC') 
        config['tron_ds'] = c_t2.checkbox("Đảo Đ/S", value=(not disable_ds), disabled=disable_ds)

        # --- ĐIỂM SỐ ---
        if config['loai_mon'] == 'MON_KHAC':
            st.divider()
            st.header("5. THANG ĐIỂM")
            with st.expander("Nhập tổng điểm từng phần"):
                d1, d2 = st.columns(2)
                config['diem_p1'] = d1.number_input("P.I:", min_value=0.0, value=4.0, step=0.1, format="%.2f")
                config['diem_p2'] = d2.number_input("P.II:", min_value=0.0, value=4.0, step=0.1, format="%.2f")
                config['diem_p3'] = d1.number_input("P.III:", min_value=0.0, value=2.0, step=0.1, format="%.2f")
                config['diem_p4'] = d2.number_input("P.IV:", min_value=0.0, value=0.0, step=0.1, format="%.2f")

        # --- XUẤT EXCEL ---
        st.divider()
        st.header("6. XUẤT ĐÁP ÁN EXCEL")
        
        if config['loai_mon'] == 'MON_KHAC':
            excel_opts = {
                "Dọc nối tiếp (Mã | Câu | Đáp án)": 1,
                "Dọc song song (Câu | Mã 101 | Mã 102)": 2,
                "Ngang nối tiếp (Dải ruy-băng)": 3,
                "Ngang song song (Mã | 1 | 2 | 3)": 4
            }
            chon_ex = st.selectbox("Chọn định dạng Excel:", list(excel_opts.keys()), index=0)
            config['excel_mode'] = excel_opts[chon_ex]
        else:
            st.info("Môn Tiếng Anh & ĐGNL được thiết lập mặc định xuất Excel kiểu: Dọc nối tiếp (Mã | Câu | Đáp án) để tránh lệch dữ liệu do xóc tự do.")
            config['excel_mode'] = 1

    return config

def hien_thi_man_hinh_chinh(config):
    st.title("📂 TẢI ĐỀ GỐC & XỬ LÝ")
    
    mon_lbl = "KHTN/KHXH"
    if config['loai_mon'] == 'ENG': mon_lbl = "Tiếng Anh"
    elif config['loai_mon'] == 'DGNL': mon_lbl = "Đánh Giá Năng Lực"
    
    st.info(f"Đang làm việc với: **{mon_lbl}** | Tạo: **{config['so_luong_de']} đề**")
    
    return {'file_de_goc': st.file_uploader("Kéo thả file đề gốc (.docx) vào đây:", type=['docx'])}