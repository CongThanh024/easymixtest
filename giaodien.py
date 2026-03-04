import streamlit as st
import os

# --- LƯU Ý: KHÔNG ĐƯỢC TỰ Ý THAY ĐỔI GIAO DIỆN NÀY NẾU KHÔNG CÓ YÊU CẦU ---

def khoi_tao_session_state():
    """Khởi tạo trạng thái để lưu vết tương tác và tránh lỗi Nút Hướng dẫn"""
    if "show_guide" not in st.session_state:
        st.session_state["show_guide"] = False

    defaults = {
        'h_so': '', 'h_truong': '', 'h_to': '', 
        'h_kythi': '', 'h_namhoc': '', 'h_mon': '', 'h_thoigian': ''
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val

def hien_thi_sidebar(supabase=None):
    khoi_tao_session_state()
    config = {}

    with st.sidebar:
        # --- NÚT HƯỚNG DẪN ---
        if st.button("📖 Xem Hướng dẫn trộn đề", use_container_width=True):
            st.session_state["show_guide"] = not st.session_state["show_guide"]
            
        st.divider()

        # --- 1. CẤU HÌNH ---
        st.header("1. CẤU HÌNH")
        
        loai_mon_input = st.radio("Chọn Môn Thi:", ["Môn KHTN/KHXH (Toán, Hóa...)", "Môn Tiếng Anh", "Môn Đánh Giá Năng Lực"], index=0, key="luu_chon_mon")
        config['loai_mon'] = 'MON_KHAC' if "KHTN" in loai_mon_input else 'ENG' if "Tiếng Anh" in loai_mon_input else 'DGNL'

        c1, c2 = st.columns(2)
        with c1:
            config['so_luong_de'] = st.number_input("Số đề:", min_value=1, max_value=99, value=4, key="luu_so_de")
        
        che_do_ma = st.radio("Cách sinh Mã đề:", ["Bắt đầu từ...", "Ngẫu nhiên"], key="luu_che_do_ma")
        if che_do_ma == "Bắt đầu từ...":
            config['ma_de_start'] = st.number_input("Mã bắt đầu:", value=101, key="luu_ma_bat_dau")
            config['kieu_ma_de'] = 'SEQUENTIAL'
        else:
            kieu_ngau_nhien = st.selectbox("Độ dài mã:", ["3 chữ số (VD: 142)", "4 chữ số (VD: 7924)"], key="luu_kieu_ngau_nhien")
            config['kieu_ma_de'] = 'RANDOM_4' if "4" in kieu_ngau_nhien else 'RANDOM_3'

        st.divider()

        # --- 2. TRÌNH BÀY ĐỀ THI (ĐÃ FIX TIỀN TỐ) ---
        st.header("2. TRÌNH BÀY ĐỀ THI")
        
        config['co_header'] = st.checkbox("Gắn Tiêu Đề (Sở/Trường...)", value=True, key="luu_tuy_chon_header")
        if config['co_header']:
            st.caption("✨ *Hệ thống tự động thêm chữ 'Sở', 'Trường', 'Tổ', 'Năm học'. Bạn chỉ cần nhập tên!*")
            
            so_in = st.text_input("Tên Sở:", key="h_so", placeholder="VD: TP HỒ CHÍ MINH")
            truong_in = st.text_input("Tên Trường:", key="h_truong", placeholder="VD: TRẦN KHAI NGUYÊN")
            to_in = st.text_input("Tên Tổ:", key="h_to", placeholder="VD: TOÁN - TIN")
            kythi_in = st.text_input("Tên Kỳ thi:", key="h_kythi", placeholder="VD: KIỂM TRA GIỮA HỌC KỲ I")
            namhoc_in = st.text_input("Năm học (Chỉ cần nhập số):", key="h_namhoc", placeholder="VD: 2025 - 2026")
            mon_in = st.text_input("Tên Môn:", key="h_mon", placeholder="VD: TOÁN")
            thoigian_in = st.text_input("Thời gian (Chỉ nhập số phút):", key="h_thoigian", placeholder="VD: 90")
            
            # --- [TÍNH NĂNG MỚI] NÚT LƯU LÊN ĐÁM MÂY ---
            if st.button("💾 Lưu làm mặc định cho tài khoản này", use_container_width=True):
                if supabase and "email" in st.session_state:
                    du_lieu_luu = {
                        'h_so': so_in, 'h_truong': truong_in, 'h_to': to_in,
                        'h_kythi': kythi_in, 'h_namhoc': namhoc_in, 
                        'h_mon': mon_in, 'h_thoigian': thoigian_in
                    }
                    try:
                        # Ghi thẳng vào cột jsonb mà bạn vừa tạo
                        supabase.table("users_data").update({"cau_hinh_mac_dinh": du_lieu_luu}).eq("email", st.session_state["email"]).execute()
                        st.success("✅ Đã lưu cấu hình lên Đám mây!")
                    except Exception as e:
                        st.error(f"Lỗi lưu Đám mây: {e}")

            # Logic nối chuỗi thông minh
            val_so = f"SỞ GD VÀ ĐT {so_in}".replace("SỞ GD VÀ ĐT SỞ GD", "SỞ GD") if so_in else ""
            val_truong = f"TRƯỜNG THPT {truong_in}".replace("TRƯỜNG THPT TRƯỜNG THPT", "TRƯỜNG THPT") if truong_in else ""
            val_to = f"TỔ {to_in}".replace("TỔ TỔ", "TỔ") if to_in else ""
            val_namhoc = f"NĂM HỌC: {namhoc_in}".replace("NĂM HỌC: NĂM HỌC:", "NĂM HỌC:") if namhoc_in else ""
            val_thoigian = f"Thời gian làm bài: {thoigian_in} phút" if thoigian_in else ""
            val_mon = f"Môn: MÔN {mon_in}".replace("Môn: MÔN MÔN", "Môn: MÔN") if mon_in else ""
            
            config['header_data'] = {
                'so': val_so, 'truong': val_truong, 'to': val_to,
                'kythi': kythi_in, 'namhoc': val_namhoc,
                'mon': val_mon, 'thoigian': val_thoigian
            }

        st.divider()
        
        # --- CHÈN ẢNH PHIẾU ---
        st.header("3. CHÈN PHIẾU LÀM BÀI")
        with st.expander("Tải lên mẫu Phiếu"):
            config['img_phieu_to'] = st.file_uploader("Phiếu Tô (Trang 1):", type=['png', 'jpg', 'jpeg'], key="img1")
            config['img_tu_luan'] = st.file_uploader("Giấy Tự Luận (Trang 2):", type=['png', 'jpg', 'jpeg'], key="img2")
            
            # [MỚI] Thêm nút tải file Word Quy ước môn
            config['file_quy_uoc'] = st.file_uploader("File Quy ước môn (.docx):", type=['docx'], key="doc_quyuoc")

        st.divider()

        # --- CÀI ĐẶT TRỘN ---
        st.header("4. CÀI ĐẶT TRỘN")
        config['tron_nhom'] = st.checkbox("Trộn nhóm (Cluster)", value=False, help="Hoán vị vị trí các nhóm câu hỏi")
        
        c_t1, c_t2 = st.columns(2)
        config['tron_mcq'] = c_t1.checkbox("Đảo A,B,C,D", value=True)
        
        # DGNL/AV KHÔNG CÓ CÂU ĐÚNG SAI -> ẨN/CHÌM TÙY CHỌN NÀY
        disable_ds = (config['loai_mon'] != 'MON_KHAC') # Nếu không phải KHTN thì disable
        config['tron_ds'] = c_t2.checkbox("Đảo Đ/S", value=(not disable_ds), disabled=disable_ds)

        # --- ĐIỂM SỐ ---
        if config['loai_mon'] == 'MON_KHAC':
            st.divider()
            st.header("5. THANG ĐIỂM")
            with st.expander("Nhập tổng điểm từng phần"):
                d1, d2 = st.columns(2)
                config['diem_p1'] = d1.number_input("P.I:", min_value=0.0, value=0.0, step=0.1, format="%.2f")
                config['diem_p2'] = d2.number_input("P.II:", min_value=0.0, value=0.0, step=0.1, format="%.2f")
                config['diem_p3'] = d1.number_input("P.III:", min_value=0.0, value=0.0, step=0.1, format="%.2f")
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
    """Hiển thị màn hình chính."""
    st.title("📂 TẢI ĐỀ GỐC & XỬ LÝ")
    
    # Hiển thị tên môn rõ ràng
    mon_lbl = "KHTN/KHXH"
    if config['loai_mon'] == 'ENG': mon_lbl = "Tiếng Anh"
    elif config['loai_mon'] == 'DGNL': mon_lbl = "Đánh Giá Năng Lực"
    
    st.info(f"Đang làm việc với: **{mon_lbl}** | Tạo: **{config['so_luong_de']} đề**")
    
    return {'file_de_goc': st.file_uploader("Kéo thả file đề gốc (.docx) vào đây:", type=['docx'])}