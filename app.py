import streamlit as st
import pandas as pd
from datetime import datetime
import os

st.set_page_config(page_title="MISA SME - LIEN CO SO", layout="wide")

PRODUCT_FILE = "kho_hang_lien_co_so.xlsx"
HISTORY_FILE = "lich_su_lien_co_so.xlsx"

def load_data():
    if os.path.exists(PRODUCT_FILE):
        st.session_state.products = pd.read_excel(PRODUCT_FILE).to_dict(orient="records")
    else:
        st.session_state.products = [
            {"Ma_hang": "2.TT6", "Ten_hang": "Thep tam 6lyx 290x3000x4t...", "Don_gia": 23500, "Co_so": "Co so 1A"},
            {"Ma_hang": "2.TT6", "Ten_hang": "Thep tam 6lyx 290x3000x4t...", "Don_gia": 23500, "Co_so": "Co so 1B"}
        ]
        pd.DataFrame(st.session_state.products).to_excel(PRODUCT_FILE, index=False)

    if os.path.exists(HISTORY_FILE):
        st.session_state.history = pd.read_excel(HISTORY_FILE).to_dict(orient="records")
    else:
        st.session_state.history = [
            {"Co_so": "Co so 1A", "Ngay_hach_toan": "24/05/2026", "So_chung_tu": "bh606", "So_hoa_don": "0000153", "Ma_khach_hang": "CT_TIENMINH", "Khach_hang": "gra Thu Thanh Ky Chau", "Ma_hang": "2.TT6", "Ten_hang": "Thep tam 6lyx 290x3000x4t...", "So_luong": 207.95, "Don_gia": 23500, "Thanh_tien": 4887000}
        ]
        pd.DataFrame(st.session_state.history).to_excel(HISTORY_FILE, index=False)

if 'products' not in st.session_state or 'history' not in st.session_state:
    load_data()

# --- MENU CHON VA LOC THEO CO SO ---
st.sidebar.title("MISA SME LIEN CO SO")
co_so_user = st.sidebar.selectbox("BAN DANG O CO SO:", ["Phong Ke Toan", "Co so 1A", "Co so 1B"])

menu = st.sidebar.radio("PHAN HE CHUC NANG", ["💼 Ban hang (Chung tu)", "🖨️ In Don dat hang"])

df_hist = pd.DataFrame(st.session_state.history)

# Tự động lọc dữ liệu theo cơ sở (Kế toán xem hết, cơ sở nào chỉ xem cơ sở đó)
if co_so_user != "Phong Ke Toan":
    df_filtered = df_hist[df_hist["Co_so"] == co_so_user]
else:
    df_filtered = df_hist

# --- PHAN HE 1: QUAN LY CHUNG TU ---
if menu == "💼 Ban hang (Chung tu)":
    st.title(f"Ban hang - Danh sach chung tu ({co_so_user})")
    
    # Quyền lập hóa đơn (Chỉ cho 2 cơ sở lập, Kế toán chỉ xem và duyệt)
    if co_so_user != "Phong Ke Toan":
        with st.expander(f"➕ Lap chung tu ban hang moi cho {co_so_user}"):
            with tk_form := st.form("form_invoice", clear_on_submit=True):
                c1, c2, c3 = st.columns(3)
                with c1:
                    ma_kh = st.text_input("Ma khach hang", value="CT_TIENMINH")
                    ten_kh = st.text_input("Ten khach hang", value="gra Thu Thanh Ky Chau")
                with c2:
                    so_ct = st.text_input("So chung tu", value=f"CT-{co_so_user.replace(' ', '')}-{len(df_filtered)+1}")
                    so_hd = st.text_input("So hoa don", value="0000153")
                with c3:
                    ma_h = st.text_input("Ma hang", value="2.TT6")
                    ten_h = st.text_input("Ten hang", value="Thep tam 6lyx 200x3000x4t...")
                    sl = st.number_input("So luong xuat", min_value=0.0, value=100.0, step=0.1)
                    gia = st.number_input("Don gia", min_value=0, value=23500)
                
                btn_sub = st.form_submit_button("💾 Cat va Ghi so (Luu hoa don)")
                if btn_sub:
                    thanh_tien = int(gia * sl)
                    new_row = {
                        "Co_so": co_so_user, "Ngay_hach_toan": datetime.now().strftime("%d/%m/%Y"),
                        "So_chung_tu": so_ct, "So_hoa_don": so_hd, "Ma_khach_hang": ma_kh, "Khach_hang": ten_kh,
                        "Ma_hang": ma_h, "Ten_hang": ten_h, "So_luong": sl, "Don_gia": gia, "Thanh_tien": thanh_tien
                    }
                    st.session_state.history.append(new_row)
                    pd.DataFrame(st.session_state.history).to_excel(HISTORY_FILE, index=False)
                    st.success("Da Ghi so chung tu len he thong!")
                    st.rerun()

    st.subheader("Danh sach hoa don phia tren (Master)")
    if df_filtered.empty:
        st.info("Chua co hoa don nao phat sinh.")
    else:
        df_master = df_filtered[["Co_so", "Ngay_hach_toan", "So_chung_tu", "So_hoa_don", "Khach_hang"]].drop_duplicates()
        df_sums = df_filtered.groupby("So_chung_tu")["Thanh_tien"].sum().reset_index(name="Tong_tien")
        df_master = pd.merge(df_master, df_sums, on="So_chung_tu")
        st.dataframe(df_master, use_container_width=True)
        
        st.subheader("Chi tiet hang hoa phia duoi (Detail)")
        list_ct = df_master["So_chung_tu"].tolist()
        selected_ct = st.selectbox("Chon So chung tu de xem chi tiet", list_ct)
        df_detail = df_filtered[df_filtered["So_chung_tu"] == selected_ct][["Ma_hang", "Ten_hang", "So_luong", "Don_gia", "Thanh_tien"]]
        st.dataframe(df_detail, use_container_width=True)

# --- PHAN HE 2: IN MAU DON DAT HANG ---
elif menu == "🖨️ In Don dat hang":
    st.title("🖨️ Mau In Don Dat Hang")
    if df_filtered.empty:
        st.info("Chua co hoa don de hien thi mau in.")
    else:
        list_ct = df_filtered["So_chung_tu"].drop_duplicates().tolist()
        selected_ct = st.selectbox("Chon so don hang de hien thi mau in:", list_ct)
        df_select = df_filtered[df_filtered["So_chung_tu"] == selected_ct]
        
        if not df_select.empty:
            m_info = df_select.iloc[0]
            tong_cong = int(df_select["Thanh_tien"].sum())
            
            st.write("---")
            st.write(f"**CONG TY TNHH THUONG MAI VA DICH VU TONG HOP TIEN MINH - {m_info['Co_so'].upper()}**")
            st.write("Lo 04 Khu cong nghiep phu tro, Tinh Ha Tinh, Viet Nam.")
            st.markdown("<h2 style='text-align: center;'>DON DAT HANG</h2>", unsafe-allow_html=True)
            
            col_m1, col_m2 = st.columns(2)
            with col_m1:
                st.write(f"**Ten khach hang:** {m_info['Khach_hang']}")
                st.write(f"**Ma khach hang:** {m_info['Ma_khach_hang']}")
                st.write("**Dia chi:** Ky Anh - Ha Tinh")
            with col_m2:
                st.write(f"**So chung tu:** {m_info['So_chung_tu']}")
                st.write(f"**Ngay hach toan:** {m_info['Ngay_hach_toan']}")
                st.write("**Loai tien:** VND")
                
            st.write("")
            df_print = df_select[["Ma_hang", "Ten_hang", "So_luong", "Don_gia", "Thanh_tien"]].copy()
            df_print.columns = ["Ma hang", "Ten hang", "So luong", "Don gia", "Thanh tien"]
            st.dataframe(df_print, use_container_width=True)
            
            st.write(f"**Tong tien thanh toan:** {tong_cong:,} VND")
            st.write("**So tien bang chu:** Bay trieu bon tram muoi ba nghin dong chan./.")
            st.write("---")
            st.write("**Nguoi mua hang** | **Ke toan truong** | **Nguoi lap phieu**")
            st.write("*(Ky, ho ten)* | *(Ky, ho ten)* | *(Ky, ho ten)*")
            st.write("")
            st.write("💡 *Nhan to hop phim **Ctrl + P** tren ban phim de tien hanh in phieu nay ra giay hoac luu file PDF.*")
