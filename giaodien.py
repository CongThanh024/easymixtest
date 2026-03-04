import streamlit as st
import os

# --- LƯU Ý: KHÔNG ĐƯỢC TỰ Ý THAY ĐỔI GIAO DIỆN NÀY NẾU KHÔNG CÓ YÊU CẦU ---

def khoi_tao_session_state():
    """Khởi tạo trạng thái nhớ TOÀN BỘ UI cho tài khoản đăng nhập."""
    if "show_guide" not in st.session_state:
        st.session_state["show_guide"] = False

    defaults = {
        'header_so': '', 'header_truong': '', 'header_to': '', 
        'header_kythi': '', 'header_namhoc': '', 'header_mon': '', 'header_thoigian': '',
        'footer_gt1': 'Giám thị 1', 'footer_gt2': 'Giám thị 2',
        
        # --- BỘ NHỚ CẤU HÌNH (THÊM MỚI) ---
        'ui_loai_mon': "Môn KHTN/KHXH (Toán, Hóa...)", 'ui_so_de': 4,
        'ui_che_do_ma': "Bắt đầu từ...", 'ui_ma_bat_dau': 101, 'ui_kieu_ngau_nhien': "3 chữ số (VD: 142)",
        'ui_co_header': True, 'ui_co_footer': True, 'ui_tron_nhom': False,
        'ui_tron_mcq': True, 'ui_tron_ds': True,
        'ui_diem_p1': 0.0, 'ui_diem_p2': 0.0, 'ui_diem_p3': 0.0, 'ui_diem_p4': 0.0,
        'ui_excel_mode': "Dọc nối tiếp (Mã | Câu | Đáp án)"
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
            key="ui_loai_mon"
        )
        
        if "KHTN" in loai_mon_input:
            config['loai_mon'] = 'MON_KHAC'
        elif "Tiếng Anh" in loai_mon_input:
            config['loai_mon'] = 'ENG'
        else:
            config['loai_mon'] = 'DGNL'

        c1, c2 = st.columns(2)
        with c1:
            config['so_luong_de'] = st.number_input("Số đề:", min_value=1, max_value=99, key="ui_so_de")
        
        che_do_ma = st.radio("Cách sinh Mã đề:", ["Bắt đầu từ...", "Ngẫu nhiên"], key="ui_che_do_ma")
        if che_do_ma == "Bắt đầu từ...":
            config['ma_de_start'] = st.number_input("Mã bắt đầu:", key="ui_ma_bat_dau")
            config['kieu_ma_de'] = 'SEQUENTIAL'
        else:
            kieu_ngau_nhien = st.selectbox("Độ dài mã:", ["3 chữ số (VD: 142)", "4 chữ số (VD: 7924)"], key="ui_kieu_ngau_nhien")
            config['kieu_ma_de'] = 'RANDOM_4' if "4 chữ số" in kieu_ngau_nhien else 'RANDOM_3'

        st.divider()

        # --- 2. TRÌNH BÀY ĐỀ THI ---
        st.header("2. TRÌNH BÀY ĐỀ THI")
        
        config['co_header'] = st.checkbox("Gắn Tiêu Đề (Sở/Trường...)", key="ui_co_header")
        if config['co_header']:
            st.caption("✨ *Hệ thống tự động thêm chữ 'Sở GD và ĐT', 'Trường', 'Tổ', 'Năm học'. Bạn chỉ cần nhập tên ngắn gọn!*")
            
            so_in = st.text_input("Tên Sở (Chỉ nhập tên Tỉnh/TP):", key="header_so", placeholder="VD: HÀ NỘI hoặc TP HỒ CHÍ MINH")
            truong_in = st.text_input("Tên Trường (Chỉ nhập tên):", key="header_truong", placeholder="VD: CHUYÊN KHTN")
            to_in = st.text_input("Tổ Chuyên Môn (Chỉ nhập tên):", key="header_to", placeholder="VD: TOÁN - TIN")
            kythi_in = st.text_input("Kỳ Thi:", key="header_kythi", placeholder="VD: KIỂM TRA GIỮA HỌC KỲ I")
            namhoc_in = st.text_input("Năm học (Chỉ cần nhập số):", key="header_namhoc", placeholder="VD: 2025 - 2026")
            mon_in = st.text_input("Môn Thi (Chỉ nhập tên môn):", key="header_mon", placeholder="VD: TOÁN 12")
            thoigian_in = st.text_input("Thời gian (Chỉ nhập số phút):", key="header_thoigian", placeholder="VD: 90")
            
            # --- BỘ LỌC CHỐNG LẶP TỪ THÔNG MINH ---
            val_so = f"SỞ GD VÀ ĐT {so_in}" if so_in else ""
            if val_so: val_so = val_so.replace("SỞ GD VÀ ĐT SỞ GD VÀ ĐT", "SỞ GD VÀ ĐT").replace("SỞ GD VÀ ĐT SỞ GD", "SỞ GD VÀ ĐT").replace("SỞ GD VÀ ĐT SỞ", "SỞ GD VÀ ĐT")
                
            val_truong = f"TRƯỜNG THPT {truong_in}" if truong_in else ""
            if val_truong: 
                # [ĐÃ SỬA] Bắt thêm lỗi lặp chữ "THPT THPT"
                val_truong = val_truong.replace("TRƯỜNG THPT TRƯỜNG THPT", "TRƯỜNG THPT").replace("TRƯỜNG THPT TRƯỜNG", "TRƯỜNG THPT").replace("THPT THPT", "THPT")
                
            val_to = f"TỔ {to_in}" if to_in else ""
            if val_to: val_to = val_to.replace("TỔ TỔ", "TỔ")
                
            val_namhoc = f"NĂM HỌC: {namhoc_in}" if namhoc_in else ""
            if val_namhoc: val_namhoc = val_namhoc.replace("NĂM HỌC: NĂM HỌC:", "NĂM HỌC:").replace("NĂM HỌC: NĂM HỌC", "NĂM HỌC:")
                
            # [ĐÃ SỬA] KHÔNG bọc chữ MÔN nữa vì xuat_file_word.py đã có sẵn chữ "Môn: " rồi!
            val_mon = mon_in if mon_in else ""
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
        config['co_footer'] = st.checkbox("Gắn Tiêu Đề Kết Thúc (Chữ ký)", key="ui_co_footer")
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
        config['tron_nhom'] = st.checkbox("Trộn nhóm (Cluster)", help="Hoán vị vị trí các nhóm câu hỏi", key="ui_tron_nhom")
        
        c_t1, c_t2 = st.columns(2)
        config['tron_mcq'] = c_t1.checkbox("Đảo A,B,C,D", key="ui_tron_mcq")
        
        disable_ds = (config['loai_mon'] != 'MON_KHAC') 
        config['tron_ds'] = c_t2.checkbox("Đảo Đ/S", disabled=disable_ds, key="ui_tron_ds")

        # --- ĐIỂM SỐ ---
        if config['loai_mon'] == 'MON_KHAC':
            st.divider()
            st.header("5. THANG ĐIỂM")
            with st.expander("Nhập tổng điểm từng phần"):
                d1, d2 = st.columns(2)
                config['diem_p1'] = d1.number_input("P.I: TN", min_value=0.0, step=0.1, format="%.2f", key="ui_diem_p1")
                config['diem_p2'] = d2.number_input("P.II: Đ/S", min_value=0.0, step=0.1, format="%.2f", key="ui_diem_p2")
                config['diem_p3'] = d1.number_input("P.III: TLN", min_value=0.0, step=0.1, format="%.2f", key="ui_diem_p3")
                config['diem_p4'] = d2.number_input("P.IV: TL", min_value=0.0, step=0.1, format="%.2f", key="ui_diem_p4")

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
            chon_ex = st.selectbox("Chọn định dạng Excel:", list(excel_opts.keys()), key="ui_excel_mode")
            config['excel_mode'] = excel_opts[chon_ex]
        else:
            st.info("Môn Tiếng Anh & ĐGNL được thiết lập mặc định xuất Excel kiểu: Dọc nối tiếp (Mã | Câu | Đáp án) để tránh lệch dữ liệu do xóc tự do.")
            config['excel_mode'] = 1
            
        # --- 7. ĐỒNG BỘ ĐÁM MÂY TỔNG ---
        st.divider()
        st.header("7. ĐỒNG BỘ ĐÁM MÂY")
        st.caption("Lưu lại toàn bộ các tùy chọn từ Mục 1 đến Mục 6 để lần sau đăng nhập không phải chọn lại.")
        
        if st.button("💾 LƯU MỌI CÀI ĐẶT LÀM MẶC ĐỊNH", use_container_width=True, type="primary"):
            if supabase and "email" in st.session_state:
                keys_to_save = [
                    'header_so', 'header_truong', 'header_to', 'header_kythi', 'header_namhoc', 'header_mon', 'header_thoigian',
                    'footer_gt1', 'footer_gt2', 'ui_loai_mon', 'ui_so_de', 'ui_che_do_ma', 'ui_ma_bat_dau', 'ui_kieu_ngau_nhien',
                    'ui_co_header', 'ui_co_footer', 'ui_tron_nhom', 'ui_tron_mcq', 'ui_tron_ds',
                    'ui_diem_p1', 'ui_diem_p2', 'ui_diem_p3', 'ui_diem_p4', 'ui_excel_mode'
                ]
                du_lieu_luu = {k: st.session_state[k] for k in keys_to_save if k in st.session_state}
                try:
                    supabase.table("users_data").update({"cau_hinh_mac_dinh": du_lieu_luu}).eq("email", st.session_state["email"]).execute()
                    st.success("✅ Đã lưu TOÀN BỘ cấu hình lên Đám mây!")
                except Exception as e:
                    st.error(f"Lỗi lưu Đám mây: {e}")
    return config

def hien_thi_man_hinh_chinh(config):
    st.title("📂 TẢI ĐỀ GỐC & XỬ LÝ")
    
    mon_lbl = "KHTN/KHXH"
    if config['loai_mon'] == 'ENG': mon_lbl = "Tiếng Anh"
    elif config['loai_mon'] == 'DGNL': mon_lbl = "Đánh Giá Năng Lực"
    
    st.info(f"Đang làm việc với: **{mon_lbl}** | Tạo: **{config['so_luong_de']} đề**")
    
    return {'file_de_goc': st.file_uploader("Kéo thả file đề gốc (.docx) vào đây:", type=['docx'])}