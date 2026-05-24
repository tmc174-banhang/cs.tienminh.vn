import streamlit as st
import pandas as pd
from datetime import datetime
import os
import time

st.set_page_config(page_title="MISA SME - TOI UU SO CHUNG TU", layout="wide")

PRODUCT_FILE = "kho_hang_lien_co_so.xlsx"
HISTORY_FILE = "lich_su_lien_co_so.xlsx"

USER_CREDENTIALS = {
    "ketoan_tienminh": {"password": "tienminh2026", "role": "Phong Ke Toan"},
    "coso_1a": {"password": "1a@tienminh", "role": "Co so 1A"},
    "coso_1b": {"password": "1b@tienminh", "role": "Co so 1B"}
}

if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'user_role' not in st.session_state: st.session_state.user_role = None

def doc_so_tien_thanh_chu(so_tien):
    if so_tien == 0: return "Khong dong"
    chu_so = ["khong", "mot", "hai", "ba", "bon", "nam", "sau", "bay", "tam", "chin"]
    
    def doc_block_3_so(n, m):
        if n == 0: return ""
        tram = n // 100
        chuc = (n % 100) // 10
        don_vi = n % 10
        kq = ""
        if m or tram > 0: kq += chu_so[tram] + " tram "
        if chuc == 0:
            if m and don_vi > 0: kq += "le "
        elif chuc == 1: kq += "muoi "
        else: kq += chu_so[chuc] + " muoi "
        if don_vi > 0:
            if don_vi == 1 and chuc > 1: kq += "mot "
            elif don_vi == 5 and chuc > 0: kq += "lam "
            else: kq += chu_so[don_vi] + " "
        return kq

    tram_trieu = (so_tien // 1000000) % 1000
    tram_nghin = (so_tien // 1000) % 1000
    dong = so_tien % 1000
    
    chuoi_chu = ""
    if tram_trieu > 0: chuoi_chu += doc_block_3_so(tram_trieu, False) + "trieu "
    if tram_nghin > 0: chuoi_chu += doc_block_3_so(tram_nghin, tram_trieu > 0) + "nghin "
    if dong > 0: chuoi_chu += doc_block_3_so(dong, tram_nghin > 0 or tram_trieu > 0)
        
    chuoi_chu = chuoi_chu.strip()
    if chuoi_chu: chuoi_chu = chuoi_chu.upper() + chuoi_chu[1:]
    return chuoi_chu + " dong chan./."

if not st.session_state.logged_in:
    st.title("MISA SME TIEN MINH")
    st.subheader("HE THONG QUAN LY LIEN CO SO CO BAO MAT")
    with st.form("login_form"):
        u_input = st.text_input("Ten tai khoan (Username)")
        p_input = st.text_input("Mat khau (Password)", type="password")
        if st.form_submit_button("Dang Nhap"):
            if u_input in USER_CREDENTIALS and USER_CREDENTIALS[u_input]["password"] == p_input:
                st.session_state.logged_in = True
                st.session_state.user_role = USER_CREDENTIALS[u_input]["role"]
                st.rerun()
            else: st.error("Sai tai khoan hoac mat khau!")
else:
    co_so_user = st.session_state.user_role
    st.sidebar.title("MISA SME LIEN CO SO")
    st.sidebar.write(f"🔒 Tai khoan: **{co_so_user}**")
    if st.sidebar.button("🔓 Dang xuat"):
        st.session_state.logged_in = False
        st.session_state.user_role = None
        st.rerun()
        
    menu = st.sidebar.radio("PHAN HE CHUC NANG", ["Ban hang (Chung tu)", "In Don dat hang"])

    def get_next_so_ct():
        for _ in range(20):
            try:
                if os.path.exists(HISTORY_FILE):
                    df = pd.read_excel(HISTORY_FILE)
                    if not df.empty:
                        df_bh = df[df["So_chung_tu"].str.startswith("BH", na=False)]
                        if not df_bh.empty:
                            last_so_ct = df_bh.iloc[-1]["So_chung_tu"]
                            last_num = int(last_so_ct.replace("BH", ""))
                            return f"BH{(last_num + 1):03d}"
                return "BH001"
            except: time.sleep(0.05)
        return f"BH{int(time.time())}"

    if menu == "Ban hang (Chung tu)":
        st.title(f"Ban hang - Danh sach chung tu ({co_so_user})")
        if os.path.exists(HISTORY_FILE): df_hist = pd.read_excel(HISTORY_FILE)
        else: df_hist = pd.DataFrame()

        if co_so_user != "Phong Ke Toan":
            df_filtered = df_hist[df_hist["Co_so"] == co_so_user] if not df_hist.empty else pd.DataFrame()
            next_code_preview = get_next_so_ct()
            
            with st.expander(f"Lap chung tu ban hang moi (Ma goi y tiep theo: {next_code_preview})"):
                with st.form("form_invoice", clear_on_submit=True):
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        ma_kh = st.text_input("Ma khach hang", value="CT_TIENMINH")
                        ten_kh = st.text_input("Ten khach hang", value="gra Thu Thanh Ky Chau")
                    with c2:
                        st.text_input("Ma chung tu he thong se cap:", value=next_code_preview, disabled=True)
                        so_hd = st.text_input("So hoa don", value="0000153")
                    with c3:
                        ma_h = st.text_input("Ma hang", value="2.TT6")
                        ten_h = st.text_input("Ten hang", value="Thep tam 6lyx 200x3000x4t...")
                        sl = st.number_input("So luong xuat", min_value=0.0, value=100.0, step=0.1)
                        gia = st.number_input("Don gia", min_value=0, value=23500)
                    
                    if st.form_submit_button("Cat va Ghi so (Luu hoa don)"):
                        thanh_tien = int(gia * sl)
                        so_ct_final = get_next_so_ct()
                        new_row = {
                            "Co_so": co_so_user, "Ngay_hach_toan": datetime.now().strftime("%d/%m/%Y"),
                            "So_chung_tu": so_ct_final, "So_hoa_don": so_hd, "Ma_khach_hang": ma_kh, "Khach_hang": ten_kh,
                            "Ma_hang": ma_h, "Ten_hang": ten_h, "So_luong": sl, "Don_gia": gia, "Thanh_tien": thanh_tien
                        }
                        if os.path.exists(HISTORY_FILE):
                            df_current = pd.read_excel(HISTORY_FILE)
                            df_current = pd.concat([df_current, pd.DataFrame([new_row])], ignore_index=True)
                        else: df_current = pd.DataFrame([new_row])
                        df_current.to_excel(HISTORY_FILE, index=False)
                        st.success(f"Da Ghi so thanh cong voi Ma chung tu: {so_ct_final}")
                        st.rerun()

        if os.path.exists(HISTORY_FILE):
            df_hist = pd.read_excel(HISTORY_FILE)
            df_filtered = df_hist[df_hist["Co_so"] == co_so_user] if co_so_user != "Phong Ke Toan" else df_hist
                
            if not df_filtered.empty:
                st.subheader("Danh sach hoa don phia tren (Master)")
                df_master = df_filtered[["Co_so", "Ngay_hach_toan", "So_chung_tu", "So_hoa_don", "Khach_hang"]].drop_duplicates()
                df_sums = df_filtered.groupby("So_chung_tu")["Thanh_tien"].sum().reset_index(name="Tong_tien")
                df_master = pd.merge(df_master, df_sums, on="So_chung_tu")
                st.dataframe(df_master, use_container_width=True)
                
                st.subheader("Chi tiet hang hoa phia duoi (Detail)")
                list_ct = df_master["So_chung_tu"].tolist()
                selected_ct = st.selectbox("Chon So chung tu de xem chi tiet", list_ct)
                df_detail = df_filtered[df_filtered["So_chung_tu"] == selected_ct][["Ma_hang", "Ten_hang", "So_luong", "Don_gia", "Thanh_tien"]]
                st.dataframe(df_detail, use_container_width=True)
            else: st.info("Chua co hoa don nao phat sinh.")
        else: st.info("He thong chua co du lieu phat sinh.")

    elif menu == "In Don dat hang":
        st.title("Mau In Don Dat Hang")
        if os.path.exists(HISTORY_FILE):
            df_hist = pd.read_excel(HISTORY_FILE)
            df_filtered = df_hist[df_hist["Co_so"] == co_so_user] if co_so_user != "Phong Ke Toan" else df_hist
                
            if not df_filtered.empty:
                list_ct = df_filtered["So_chung_tu"].drop_duplicates().tolist()
                selected_ct = st.selectbox("Chon so don hang de hien thi mau in:", list_ct)
                df_select = df_filtered[df_filtered["So_chung_tu"] == selected_ct]
                
                if not df_select.empty:
                    m_info = df_select.iloc[0]
                    tong_cong = int(df_select["Thanh_tien"].sum())
                    chu_so_tien = doc_so_tien_thanh_chu(tong_cong)
                    
                    st.write("---")
                    st.write("CONG TY TNHH THUONG MAI VA DICH VU TONG HOP TIEN MINH")
                    st.write("Lo 04 Khu cong nghiep phu tro, Tinh Ha Tinh, Viet Nam.")
                    st.write("### DON DAT HANG")
                    
                    col_m1, col_m2 = st.columns(2)
                    with col_m1:
                        st.write(f"Ten khach hang: {m_info['Khach_hang']}")
                        st.write(f"Ma khach hang: {m_info['Ma_khach_hang']}")
                        st.write("Dia chi: Ky Anh - Ha Tinh")
                    with col_m2:
                        st.write(f"So chung tu: {m_info['So_chung_tu']}")
                        st.write(f"Ngay hach toan: {m_info['Ngay_hach_toan']}")
                        st.write("Loai tien: VND")
                        
                    st.write("")
                    df_print = df_select[["Ma_hang", "Ten_hang", "So_luong", "Don_gia", "Thanh_tien"]].copy()
                    df_print.columns = ["Ma hang", "Ten hang", "So luong", "Don gia", "Thanh tien"]
                    st.dataframe(df_print, use_container_width=True)
                    
                    st.write(f"**Tong tien thanh toan:** {tong_cong:,} VND")
                    st.write(f"**So tien bang chu:** *{chu_so_tien}*")
                    st.write("---")
                    st.write("Nguoi mua hang | Ke toan truong | Nguoi lap phieu")
                    st.write("(Ky, ho ten) | (Ky, ho ten) | (Ky, ho ten)")
                    st.write("")
