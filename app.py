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
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False

    if not st.session_state["logged_in"]:
        # --- HIỂN THỊ LOGO CĂN GIỮA ---
        col1, col2, col3 = st.columns([1, 1.2, 1]) 
        with col2:
            try:
                st.image("logo_app.png", use_container_width=True)
            except:
                st.markdown("<h2 style='text-align: center; color: #1E88E5;'>EASY MIX TEST</h2>", unsafe_allow_html=True)
                
        st.markdown("<p style='text-align: center; font-size: 16px; margin-top: -10px;'>Vui lòng đăng nhập để sử dụng hệ thống</p>", unsafe_allow_html=True)
        # ------------------------------
        
        tab1, tab2 = st.tabs(["🔐 Đăng nhập", "📝 Đăng ký dùng thử MIỄN PHÍ"])
        
        # --- TAB ĐĂNG NHẬP ---
        with tab1:
            with st.form("login_form"):
                email = st.text_input("Email của bạn")
                password = st.text_input("Mật khẩu", type="password")
                submit_login = st.form_submit_button("Đăng nhập", use_container_width=True)
                
                if submit_login:
                    try:
                        # 1. Kiểm tra tài khoản
                        user = supabase.auth.sign_in_with_password({"email": email, "password": password})
                        
                        # 2. Kiểm tra hạn sử dụng trong sổ cái
                        res = supabase.table("users_data").select("ngay_het_han").eq("email", email).execute()
                        if len(res.data) > 0:
                            han_dung = date.fromisoformat(res.data[0]["ngay_het_han"])
                            hom_nay = date.today()
                                                        
                            if hom_nay <= han_dung:
                            st.session_state["logged_in"] = True
                            st.session_state["email"] = email
                            st.session_state["han_dung"] = han_dung
                            st.rerun() 
                            else:
                            # Đoạn này thụt vào 18 dấu cách (hoặc 1 phím Tab so với chữ 'else')
                            st.error(f"⚠️ Tài khoản hết hạn vào ngày {han_dung.strftime('%d/%m/%Y')}.")
                            st.markdown(f"""
                                <a href="https://zalo.me/0937177439" target="_blank" style="text-decoration: none;">
                                    <div style="width:100%; background-color:#0068ff; color:white; text-align:center; padding:12px; border-radius:8px; font-weight:bold;">
                                        💬 Nhấn vào đây để liên hệ Zalo gia hạn (Admin)
                                    </div>
                                </a>
                            """, unsafe_allow_html=True)
                            st.stop()
                        else:
                            st.error("Không tìm thấy thông tin gói cước. Vui lòng liên hệ Admin.")
                    except Exception as e:
                        st.error("❌ Sai email hoặc mật khẩu! Vui lòng thử lại.")

        # --- TAB ĐĂNG KÝ ---
        with tab2:
            st.info("🎁 Tặng ngay 60 ngày dùng thử VIP khi đăng ký mới!")
            with st.form("register_form"):
                reg_email = st.text_input("Nhập Email (Dùng để đăng nhập)")
                reg_password = st.text_input("Tạo Mật khẩu (Ít nhất 6 ký tự)", type="password")
                reg_confirm = st.text_input("Nhập lại Mật khẩu", type="password")
                submit_reg = st.form_submit_button("Đăng ký tài khoản", use_container_width=True)
                
                if submit_reg:
                    if reg_password != reg_confirm:
                        st.error("❌ Mật khẩu nhập lại không khớp!")
                    elif len(reg_password) < 6:
                        st.error("❌ Mật khẩu quá ngắn!")
                    else:
                        try:
                            # 1. Tạo user trên Supabase
                            new_user = supabase.auth.sign_up({"email": reg_email, "password": reg_password})
                            # 2. Tính ngày hết hạn (+60 ngày) và ghi vào sổ cái users_data
                            ngay_het_han = date.today() + timedelta(days=60)
                            supabase.table("users_data").insert({"email": reg_email, "ngay_het_han": str(ngay_het_han)}).execute()
                            
                            st.success("🎉 Đăng ký thành công! Bạn được cấp 60 ngày dùng thử. Một email xác nhận đã được gửi đến hòm thư của bạn. Vui lòng kiểm tra và bấm vào link xác nhận để kích hoạt tài khoản trước khi đăng nhập.")
                            st.info("💡 Lưu ý: Nếu không thấy email, bạn hãy kiểm tra trong mục Thư rác (Spam) nhé!")
                        except Exception as e:
                            st.error(f"Lỗi đăng ký (Có thể email này đã tồn tại): {str(e)}")
        
        return False # Trả về False để đóng chặt cửa, không cho chạy phần code trộn đề

    else:
        # Nếu đã đăng nhập thành công -> Hiển thị thông tin ở góc trái
        st.sidebar.success(f"👤 User: {st.session_state['email']}\n\n⏳ Hạn dùng: {st.session_state['han_dung'].strftime('%d/%m/%Y')}")
        if st.sidebar.button("🚪 Đăng xuất"):
            st.session_state["logged_in"] = False
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
        
    config = giaodien.hien_thi_sidebar()
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