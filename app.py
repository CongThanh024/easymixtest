import streamlit as st
import pandas as pd
import os
import shutil
import time
import io  # Thư viện xử lý RAM
from supabase import create_client, Client   # <--- BẠN DÁN VÀO ĐÂY NHÉ
from datetime import date, timedelta
import giaodien
import xuly_degoc
import xuly_degoc_av 
import thuat_toan_tron
import thuat_toan_tron_av 
import xuat_file_word      
import xuat_file_word_av   
from copy import deepcopy
from streamlit_cookies_controller import CookieController

# [SỬA TÊN TAB TRÌNH DUYỆT]
st.set_page_config(page_title="App trộn đề chuyên nghiệp", page_icon="🛠", layout="wide")

def reset_trang_thai_xu_ly():
    keys = ['proc_res', 'stats', 'errs']
    for k in keys:
        if k in st.session_state: del st.session_state[k]

def cleanup_folder(folder_path):
    if os.path.exists(folder_path):
        try:
            for f in os.listdir(folder_path):
                p = os.path.join(folder_path, f)
                if os.path.isfile(p): os.remove(p)
                elif os.path.isdir(p): shutil.rmtree(p)
        except PermissionError:
            return False, f"⚠️ Đang có file trong thư mục '{folder_path}' bị mở. Vui lòng đóng tất cả các file Word liên quan và thử lại!"
        except Exception as e:
            return False, f"Lỗi xóa thư mục: {e}"
    try:
        os.makedirs(folder_path, exist_ok=True)
        return True, ""
    except Exception as e:
        return False, f"Không thể tạo thư mục: {e}"
# ==========================================================
# HỆ THỐNG ĐĂNG NHẬP VÀ QUẢN LÝ NGƯỜI DÙNG
# ==========================================================
def check_auth(supabase: Client):
    if "logged_in" not in st.session_state: st.session_state["logged_in"] = False
    if "show_guide" not in st.session_state: st.session_state["show_guide"] = False

    # --- KHỞI TẠO BỘ QUẢN LÝ COOKIE ---
    from streamlit_cookies_controller import CookieController
    cookies = CookieController()
    auto_email = cookies.get("auto_email")

    # --- TỰ ĐỘNG ĐĂNG NHẬP (BỎ QUA FORM) NẾU CÓ COOKIE ---
    if not st.session_state["logged_in"] and auto_email:
        try:
            res = supabase.table("users_data").select("ngay_het_han, cau_hinh_mac_dinh").eq("email", auto_email).execute()
            if len(res.data) > 0:
                han_dung = date.fromisoformat(res.data[0]["ngay_het_han"])
                if date.today() <= han_dung:
                    st.session_state["logged_in"] = True
                    st.session_state["email"] = auto_email
                    st.session_state["han_dung"] = han_dung
                    
                    # KÉO CẤU HÌNH ĐÃ LƯU VỀ APP
                    cau_hinh = res.data[0].get("cau_hinh_mac_dinh")
                    if cau_hinh:
                        for k, v in cau_hinh.items():
                            st.session_state[k] = v
        except Exception:
            pass 

    # --- CHIA MẶT TIỀN 3 CỘT (CỘT GIỮA RỘNG) ---
    col_trai, col_giua, col_phai = st.columns([0.6, 2.5, 1.2])
    
    with col_trai:
        st.write("") 

    # CỘT GIỮA: LOGO & BẢNG HƯỚNG DẪN MỞ RỘNG
    with col_giua:
        try:
            st.image("logo_app.png", use_container_width=True)
        except:
            st.markdown("<h2 style='text-align: center; color: #1E88E5;'>EASY MIX TEST</h2>", unsafe_allow_html=True)
            
        # [TÍNH NĂNG MỚI] - CHỈ HIỆN TÊN TÁC GIẢ KHI ĐÃ ĐĂNG NHẬP, TRÊN 1 DÒNG
        if st.session_state["logged_in"]:
            st.markdown(
                """
                <link href="https://fonts.googleapis.com/css2?family=Dancing+Script:wght@700&display=swap" rel="stylesheet">
                <div style="text-align: center; margin-bottom: 20px;">
                    <p style="font-family: 'Dancing Script', cursive; font-size: 30px; 
                              background: -webkit-linear-gradient(45deg, #0077b6, #d90429); 
                              -webkit-background-clip: text; -webkit-text-fill-color: transparent; 
                              margin: 0; font-weight: bold; line-height: 1.2;">
                        Developed by Công Thành
                    </p>
                </div>
                """, unsafe_allow_html=True
            )
        
        if st.session_state.get("show_guide"):
            with st.container(height=650, border=True):
                st.markdown("<h4 style='text-align:center;'>📖 CẨM NANG SỬ DỤNG EASY MIX TEST</h4>", unsafe_allow_html=True)
                
                import os
                found_any = False
                
                # Quét tự động từ trang 1 đến trang 10 với đúng tên file và đuôi .png của bạn
                for i in range(1, 11):
                    file_name = f"duongdan_sudung_Page{i}.png"
                    if os.path.exists(file_name):
                        try:
                            # In sát nhau không khoảng cách
                            st.image(file_name, use_container_width=True)
                            found_any = True
                        except Exception:
                            pass
                
                if not found_any:
                    st.info("⏳ Đang chờ hệ thống cập nhật ảnh hướng dẫn (duongdan_sudung_Page1.png)...")

    # CỘT PHẢI: ĐĂNG NHẬP
    with col_phai:
        if not st.session_state["logged_in"]:
            st.info("Vui lòng đăng nhập để sử dụng")
            tab1, tab2 = st.tabs(["🔐 Đăng nhập", "📝 Đăng ký"])
            
            with tab1:
                with st.form("login_form"):
                    email = st.text_input("Email của bạn", value=auto_email if auto_email else "")
                    password = st.text_input("Mật khẩu", type="password")
                    remember = st.checkbox("Lưu đăng nhập (Lần sau tự động vào)", value=True)
                    submit_login = st.form_submit_button("Đăng nhập", use_container_width=True)
                    
                    if submit_login:
                        try:
                            user = supabase.auth.sign_in_with_password({"email": email, "password": password})
                            res = supabase.table("users_data").select("ngay_het_han, cau_hinh_mac_dinh").eq("email", email).execute()
                            if len(res.data) > 0:
                                han_dung = date.fromisoformat(res.data[0]["ngay_het_han"])
                                if date.today() <= han_dung:
                                    st.session_state["logged_in"] = True
                                    st.session_state["email"] = email
                                    st.session_state["han_dung"] = han_dung
                                    
                                    cau_hinh = res.data[0].get("cau_hinh_mac_dinh")
                                    if cau_hinh:
                                        for k, v in cau_hinh.items():
                                            st.session_state[k] = v
                                    
                                    if remember: cookies.set("auto_email", email)
                                    else: cookies.remove("auto_email")
                                        
                                    st.rerun() 
                                else:
                                    st.error(f"⚠️ Tài khoản hết hạn: {han_dung.strftime('%d/%m/%Y')}.")
                                    st.stop()
                            else:
                                st.error("Không tìm thấy gói cước.")
                        except Exception as e:
                            st.error("❌ Sai email hoặc mật khẩu!")

            with tab2:
                with st.form("register_form"):
                    reg_email = st.text_input("Nhập Email")
                    reg_password = st.text_input("Mật khẩu (>6 ký tự)", type="password")
                    reg_confirm = st.text_input("Nhập lại", type="password")
                    submit_reg = st.form_submit_button("Đăng ký", use_container_width=True)
                    if submit_reg:
                        if reg_password != reg_confirm: st.error("❌ Mật khẩu không khớp!")
                        else:
                            try:
                                new_user = supabase.auth.sign_up({"email": reg_email, "password": reg_password})
                                ngay_het_han = date.today() + timedelta(days=60)
                                supabase.table("users_data").insert({"email": reg_email, "ngay_het_han": str(ngay_het_han)}).execute()
                                st.success("🎉 Đăng ký thành công!")
                            except Exception:
                                st.error("Lỗi đăng ký (Email đã tồn tại).")
            return False
        else:
            st.success(f"👤 {st.session_state['email']}")
            st.warning(f"⏳ Hạn dùng: {st.session_state['han_dung'].strftime('%d/%m/%Y')}")
            if st.button("🚪 Đăng xuất", use_container_width=True):
                st.session_state["logged_in"] = False
                cookies.remove("auto_email")
                st.rerun()
            return True
# ==========================================================
# BẢNG ĐIỀU KHIỂN DÀNH RIÊNG CHO GIÁM ĐỐC (ADMIN)
# ==========================================================
def hien_thi_trang_admin(supabase: Client):
    st.markdown("<h2 style='text-align: center; color: #D32F2F;'>👑 PHÒNG GIÁM ĐỐC - QUẢN TRỊ HỆ THỐNG</h2>", unsafe_allow_html=True)
    st.divider()
    
    # Lấy dữ liệu từ sổ cái Supabase
    try:
        res = supabase.table("users_data").select("*").execute()
        if len(res.data) == 0:
            st.info("Chưa có khách hàng nào đăng ký.")
            return
            
        df = pd.DataFrame(res.data)
        
        # Thống kê nhanh
        st.markdown("### 📊 Thống kê khách hàng")
        st.metric("Tổng số tài khoản trên hệ thống", len(df))
        
        # Hiển thị bảng danh sách
        st.markdown("### 📋 Danh sách chi tiết")
        st.dataframe(df[['email', 'ngay_het_han', 'created_at']], use_container_width=True)
        
        st.divider()
        st.markdown("### ⚙️ Cấp quyền / Gia hạn tài khoản")
        with st.form("gia_han_form"):
            col1, col2 = st.columns(2)
            with col1:
                selected_email = st.selectbox("1. Chọn Email khách hàng cần gia hạn:", df['email'].tolist())
            with col2:
                new_date = st.date_input("2. Chọn ngày hết hạn mới:")
                
            submit_gia_han = st.form_submit_button("Cập nhật gia hạn", type="primary", use_container_width=True)
            if submit_gia_han:
                supabase.table("users_data").update({"ngay_het_han": str(new_date)}).eq("email", selected_email).execute()
                st.success(f"🎉 Đã gia hạn thành công cho khách hàng: {selected_email} đến ngày {new_date.strftime('%d/%m/%Y')}")
                time.sleep(2)
                st.rerun()
    except Exception as e:
        st.error(f"Lỗi truy xuất dữ liệu: {e}")        
def main():
    # --- KHỞI TẠO KẾT NỐI ĐÁM MÂY SUPABASE ---
    try:
        url: str = st.secrets["SUPABASE_URL"]
        key: str = st.secrets["SUPABASE_KEY"]
        supabase: Client = create_client(url, key)
    except Exception as e:
        st.error(f"⚠️ Mất kết nối Đám mây: {e}. Vui lòng kiểm tra file secrets.toml")
        return 

    # --- GÁC CỔNG BẢO VỆ ---
    is_authenticated = check_auth(supabase)
    if not is_authenticated:
        return 
    # --- PHÂN QUYỀN GIÁM ĐỐC & KHÁCH HÀNG ---
    if st.session_state.get("email") == "doancongthanh024@gmail.com":
        hien_thi_trang_admin(supabase)
        return  # <--- LỆNH CHỐT CHẶN: Ép cỗ máy dừng ở đây, KHÔNG HIỆN giao diện trộn đề ở dưới nữa!
        
    config = giaodien.hien_thi_sidebar(supabase)
    inputs = giaodien.hien_thi_man_hinh_chinh(config)

    # Dán ở đây, thẳng hàng với chữ 'inputs' ở trên
    with st.expander("📖 HƯỚNG DẪN SỬ DỤNG NHANH"):
        st.info("""
        1. **Chuẩn bị:** Đề gốc đúng định dạng Câu 1, Câu 2...
        2. **Tải lên:** Chọn file Word (.docx) từ máy.
        3. **Cấu hình:** Chọn số mã đề cần trộn.
        4. **Kết quả:** Tải file nén .zip về máy.
        """)

    # 1. LOGIC CHUẨN HÓA (ĐÃ PHÂN 2 LUỒNG SẠCH/DƠ)
    if inputs['file_de_goc']:
        current_sig = f"{inputs['file_de_goc'].name}_{inputs['file_de_goc'].size}"
        can_run = False
        
        if 'last_sig' not in st.session_state or st.session_state.last_sig != current_sig:
            can_run = True
        elif 'proc_res' in st.session_state and st.session_state.proc_res.get('mon') != config['loai_mon']:
            can_run = True

        if can_run:
            reset_trang_thai_xu_ly()
            st.session_state.last_sig = current_sig
            file_buffer = io.BytesIO(inputs['file_de_goc'].getvalue())

            with st.spinner("Đang chuẩn hóa..."):
                try:
                    if not os.path.exists("TEST_TRON"): os.makedirs("TEST_TRON")

                    stream_sach, stream_vat_ly = None, None
                    if config['loai_mon'] == 'MON_KHAC':
                        stream_sach, stream_vat_ly, stats, errs = xuly_degoc.xu_ly_va_chuan_hoa(file_buffer)
                    else:
                        proc = xuly_degoc_av.XuLyDeChuanHoaAV(config['loai_mon'])
                        stream_sach, stream_vat_ly, stats, errs = proc.xu_ly(file_buffer)

                    # [LOG] Ghi file kiểm tra (Luồng vật lý)
                    if stream_vat_ly:
                        with open("TEST_TRON/de_chuan_hoa.docx", "wb") as f:
                            f.write(stream_vat_ly.getbuffer())
                        stream_vat_ly.seek(0)

                    st.session_state.proc_res = {
                        'stats': stats, 
                        'errs': errs, 
                        'stream_sach': stream_sach,      
                        'stream_vat_ly': stream_vat_ly,  
                        'mon': config['loai_mon']
                    }
                except Exception as e:
                    st.error(f"Lỗi chuẩn hóa: {e}"); st.exception(e)

        # 2. HIỂN THỊ VÀ TRỘN
        if 'proc_res' in st.session_state and st.session_state.proc_res:
            res = st.session_state.proc_res
            stats = res['stats']
            
            if res['mon'] == config['loai_mon']:
                st.divider()
                st.subheader("📊 Báo cáo Chuẩn Hóa")
                lbl_mon = "KHTN/KHXH"
                if res['mon'] in ['ENG', 'AV']: lbl_mon = "Tiếng Anh"
                elif res['mon'] == 'DGNL': lbl_mon = "Đánh Giá Năng Lực"
                
                if res['errs'] == 0:
                    st.success(f"✅ Chuẩn hóa thành công ({lbl_mon}). File gốc không có lỗi cấu trúc.")
                else:
                    st.warning(f"⚠️ Chuẩn hóa hoàn tất ({lbl_mon}) nhưng phát hiện {res['errs']} lỗi định dạng. Vui lòng tải 'File Chuẩn Hóa' bên dưới, kéo xuống trang cuối cùng để xem chi tiết lỗi và đối chiếu.")
                
                if res['mon'] == 'MON_KHAC':
                    t = stats.get('PHAN_I',0)+stats.get('PHAN_II',0)+stats.get('PHAN_III',0)+stats.get('PHAN_IV',0)
                    df = pd.DataFrame([{"P.I":stats.get('PHAN_I'), "P.II":stats.get('PHAN_II'), "P.III":stats.get('PHAN_III'), "P.IV":stats.get('PHAN_IV'), "TỔNG":t, "LỖI": res['errs']}])
                    st.dataframe(df, hide_index=True)
                else:
                    # Tách riêng bảng báo cáo cho Tiếng Anh và ĐGNL
                    if res['mon'] in ['ENG', 'AV']:
                        df_av = pd.DataFrame([{"Groups":stats.get('GROUPS'), "Ques":stats.get('QUESTIONS'), "Err": res['errs']}])
                        st.dataframe(df_av, hide_index=True)
                    else:
                        df_dgnl = pd.DataFrame([{"Số nhóm":stats.get('GROUPS'), "Số câu":stats.get('QUESTIONS'), "Lỗi": res['errs']}])
                        st.dataframe(df_dgnl, hide_index=True)

                # NÚT TẢI VỀ NHẬN LUỒNG DƠ (Có bảng lỗi)
                st.download_button(
                    label="📥 Tải File Chuẩn Hóa (Xem kết quả & Bảng báo lỗi)",
                    data=res['stream_vat_ly'],
                    file_name="De_Chuan_Hoa.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                st.divider()
                
                if st.button("🚀 TRỘN ĐỀ NGAY", type="primary"):
                    out_dir = "TEST_TRON"
                    cleanup_folder(out_dir)
                    
                    with st.spinner("Đang xử lý..."):
                        try:
                            lst = []
                            # HÀM TRỘN NHẬN LUỒNG SẠCH BONG (Đã cắt bỏ đuôi lỗi)
                            res['stream_sach'].seek(0)
                            
                            if res['mon'] == 'MON_KHAC':
                                lst = thuat_toan_tron.tron_de(res['stream_sach'], config['so_luong_de'], config)
                            else:
                                lst = thuat_toan_tron_av.tron_de(res['stream_sach'], config['so_luong_de'], config)

                            # --- BẬT/TẮT DEBUG: LƯU FILE RAW RA Ổ CỨNG (Bỏ dấu # để bật lại) ---
                            # for item in lst:
                            #     item['file_content'].save(os.path.join(out_dir, f"raw_made_{item['exam_id']}.docx"))
                            # --------------------------------------------------------------------

                            zip_buffer = None
                            if res['mon'] == 'MON_KHAC':
                                zip_buffer = xuat_file_word.xuat_ket_qua(lst, config, out_dir, stats)
                            else:
                                zip_buffer = xuat_file_word_av.xuat_ket_qua(lst, config, out_dir)

                            st.success("🎉 Xong!")
                            if zip_buffer:
                                st.download_button("📦 TẢI KẾT QUẢ ZIP", zip_buffer, "De_dapan.zip", "application/zip")
                            else: st.error("Lỗi tạo file kết quả.")
                                
                        except Exception as e:
                            st.error(f"Lỗi: {e}"); st.exception(e)

if __name__ == "__main__":
    main()