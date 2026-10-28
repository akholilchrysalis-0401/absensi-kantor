import io
import json
from datetime import date, datetime, time, timedelta
from pathlib import Path
import pandas as pd
import streamlit as st
import requests
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

st.set_page_config(page_title="Absensi Karyawan", page_icon="📋", layout="wide")

# ========== KONFIGURASI ==========
SHEET_ID = "1rsTGJS2H6uhqh84eQmlLn_uZVzopUWNxPqU_MpT9H5c"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]
BASE_DIR = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
BATAS_MASUK = time(8, 0)
BATAS_PULANG = time(17, 0)
ADMIN_PIN = str(st.secrets.get("admin_pin", "admin123"))

KARYAWAN_COL = ["nip", "nama", "divisi", "no_hp", "pin"]
ABSENSI_COL = ["tanggal", "nip", "nama", "divisi", "jam_masuk", "jam_pulang",
               "status", "lokasi", "link_foto_masuk", "link_foto_pulang", "keterangan"]

# ---------- Koneksi ----------
def get_creds():
    if "gcp_service_account" in st.secrets:
        return Credentials.from_service_account_info(dict(st.secrets["gcp_service_account"]), scopes=SCOPES)
    return Credentials.from_service_account_file(str(CREDENTIALS_FILE), scopes=SCOPES)

def init_gspread():
    return gspread.authorize(get_creds()).open_by_key(SHEET_ID)

def find_ws(sh, title):
    t = title.strip().lower()
    for ws in sh.worksheets():
        if ws.title.strip().lower() == t:
            return ws
    return None

def get_or_create_ws(sh, title, rows, cols, headers):
    ws = find_ws(sh, title)
    if ws: return ws, False
    try:
        ws = sh.add_worksheet(title=title, rows=rows, cols=cols)
        ws.append_row(headers)
        return ws, True
    except gspread.exceptions.APIError as e:
        if "already exists" in str(e):
            ws = find_ws(sh, title)
            if ws: return ws, False
        raise

def read_karyawan():
    sh = init_gspread()
    ws, created = get_or_create_ws(sh, "karyawan", 1100, len(KARYAWAN_COL), KARYAWAN_COL)
    if created: return pd.DataFrame(columns=KARYAWAN_COL)
    vals = ws.get_all_values()
    if not vals or len(vals) < 2: return pd.DataFrame(columns=KARYAWAN_COL)
    df = pd.DataFrame(vals[1:], columns=[str(v).strip() for v in vals[0]])
    for c in KARYAWAN_COL:
        if c not in df.columns: df[c] = ""
    df = df[KARYAWAN_COL]
    df = df[df.astype(str).apply(lambda c: c.str.strip()).ne("").any(axis=1)]
    df["nip"] = df["nip"].astype(str).str.strip()
    df["pin"] = df["pin"].astype(str).str.strip()
    return df.reset_index(drop=True)

def read_absensi():
    sh = init_gspread()
    ws, created = get_or_create_ws(sh, "absensi", 5000, len(ABSENSI_COL), ABSENSI_COL)
    if created: return pd.DataFrame(columns=ABSENSI_COL)
    vals = ws.get_all_values()
    if not vals or len(vals) < 2: return pd.DataFrame(columns=ABSENSI_COL)
    df = pd.DataFrame(vals[1:], columns=[str(v).strip() for v in vals[0]])
    for c in ABSENSI_COL:
        if c not in df.columns: df[c] = ""
    return df[ABSENSI_COL].reset_index(drop=True)

def append_absen(row_list):
    sh = init_gspread()
    ws, _ = get_or_create_ws(sh, "absensi", 5000, len(ABSENSI_COL), ABSENSI_COL)
    ws.append_row([str(x) for x in row_list])

def update_pulang(tanggal, nip, jam_pulang, foto_link):
    sh = init_gspread()
    ws = find_ws(sh, "absensi")
    vals = ws.get_all_values()
    for i, row in enumerate(vals[1:], start=2):
        if str(row[0]).strip() == tanggal and str(row[1]).strip() == nip:
            ws.update_cell(i, 6, jam_pulang)      # jam_pulang
            ws.update_cell(i, 10, foto_link)       # link_foto_pulang
            return

# ---------- Upload Foto ke Google Drive ----------

# https://drive.google.com/drive/u/0/folders/16jSCF9eKYW-etI3LPwc_uIqS9PEIfGBn

FOTO_FOLDER_ID = "https://drive.google.com/drive/u/0/folders/16jSCF9eKYW-etI3LPwc_uIqS9PEIfGBn"  # ganti dengan ID folder yang di-share

def upload_foto(image_bytes, filename):
    try:
        service = build("drive", "v3", credentials=get_creds(), cache_discovery=False)
        media = MediaIoBaseUpload(io.BytesIO(image_bytes), mimetype="image/jpeg", resumable=False)
        f = service.files().create(
            body={"name": filename, "parents": [FOTO_FOLDER_ID]},
            media_body=media, fields="id,webViewLink"
        ).execute()
        service.permissions().create(fileId=f["id"], body={"type": "anyone", "role": "reader"}).execute()
        return f.get("webViewLink", "")
    except Exception as e:
        return f"Gagal upload: {e}"

# ---------- Lokasi GPS ----------
def get_gps():
    try:
        from streamlit_js_eval import streamlit_js_eval
        loc = streamlit_js_eval(js_expressions="navigator.geolocation ? new Promise(r=>navigator.geolocation.getCurrentPosition(p=>r(p.coords.latitude+', '+p.coords.longitude), e=>r('Izin lokasi ditolak'))) : 'GPS tidak didukung'")
        if loc and isinstance(loc, str):
            return loc
    except Exception:
        pass
    return "Tidak terdeteksi"

# ---------- Bantuan ----------
def clean(v):
    if v is None or pd.isna(v): return ""
    return str(v).strip()

def _wib(): return datetime.utcnow() + timedelta(hours=7)
def today_str(): return _wib().date().isoformat()
def now_str(): return _wib().strftime("%H:%M:%S")

# ================= APLIKASI =================
st.title("📋 Absensi Karyawan")
st.caption("Sistem absensi online — foto selfie & lokasi GPS tersimpan otomatis")

karyawan_df = read_karyawan()

with st.sidebar:
    st.header("Menu")
    mode = st.radio("Pilih mode", ["👤 Absen Karyawan", "📊 Dashboard Admin"], label_visibility="collapsed")
    st.divider()
    st.caption(f"Jumlah karyawan terdaftar: {len(karyawan_df)}")

# ============ MODE KARYAWAN ============
if mode == "👤 Absen Karyawan":
    st.subheader("Absen Kehadiran")
    if karyawan_df.empty:
        st.warning("Belum ada data karyawan. Isi tab 'karyawan' di Google Sheet terlebih dahulu.")
        st.stop()

    if "nip_login" not in st.session_state:
        nama_list = karyawan_df["nama"].tolist()
        pilih = st.selectbox("Pilih nama karyawan", options=nama_list, key="pilih_nama")
        pin_input = st.text_input("Masukkan PIN (4 digit)", type="password", key="pin_input")
        if st.button("Masuk", type="primary"):
            row = karyawan_df[karyawan_df["nama"] == pilih]
            if not row.empty and clean(row.iloc[0]["pin"]) == clean(pin_input):
                st.session_state.nip_login = clean(row.iloc[0]["nip"])
                st.session_state.nama_login = pilih
                st.session_state.divisi_login = clean(row.iloc[0]["divisi"])
                st.rerun()
            else:
                st.error("PIN salah. Coba lagi.")
    else:
        nip = st.session_state.nip_login
        nama = st.session_state.nama_login
        divisi = st.session_state.divisi_login
        st.success(f"Login sebagai: **{nama}** ({divisi})")
        if st.button("Keluar"):
            for k in ["nip_login","nama_login","divisi_login"]:
                del st.session_state[k]
            st.rerun()

        absensi_df = read_absensi()
        hari_ini = absensi_df[(absensi_df["tanggal"] == today_str()) & (absensi_df["nip"] == nip)]

        tab1, tab2 = st.tabs(["✅ Absen Masuk / Pulang", "📝 Izin / Sakit / Cuti"])

        with tab1:
            if hari_ini.empty:
                st.info("Anda belum absen masuk hari ini.")
                foto = st.camera_input("Ambil foto selfie untuk absen MASUK", key="cam_masuk")
                if st.button("Kirim Absen Masuk", type="primary") and foto is not None:
                    with st.spinner("Mengirim absen & mengunggah foto..."):
                        jam = now_str()
                        status = "Hadir" if _wib().time() <= BATAS_MASUK else "Terlambat"
                        lokasi = get_gps()
                        link_foto = upload_foto(foto.getvalue(), f"{nip}_{today_str()}_masuk.jpg")
                        append_absen([today_str(), nip, nama, divisi, jam, "", status, lokasi, link_foto, "", ""])
                    st.success(f"✅ Absen masuk tercatat pukul {jam} — Status: {status}")
                    st.rerun()
            elif clean(hari_ini.iloc[0]["jam_pulang"]) == "":
                r = hari_ini.iloc[0]
                st.info(f"Absen masuk: {clean(r['jam_masuk'])} ({clean(r['status'])}). Silakan absen pulang.")
                foto = st.camera_input("Ambil foto selfie untuk absen PULANG", key="cam_pulang")
                if st.button("Kirim Absen Pulang", type="primary") and foto is not None:
                    with st.spinner("Mengirim absen & mengunggah foto..."):
                        jam = now_str()
                        link_foto = upload_foto(foto.getvalue(), f"{nip}_{today_str()}_pulang.jpg")
                        update_pulang(today_str(), nip, jam, link_foto)
                    st.success(f"✅ Absen pulang tercatat pukul {jam}")
                    st.rerun()
            else:
                r = hari_ini.iloc[0]
                st.success("Anda sudah selesai absen hari ini.")
                st.write(f"Masuk: {clean(r['jam_masuk'])} | Pulang: {clean(r['jam_pulang'])} | Status: {clean(r['status'])}")

        with tab2:
            jenis = st.selectbox("Jenis", ["Izin", "Sakit", "Cuti"])
            keterangan = st.text_area("Keterangan alasan")
            if st.button("Ajukan", type="primary") and keterangan.strip():
                if hari_ini.empty:
                    append_absen([today_str(), nip, nama, divisi, "", "", jenis, get_gps(), "", "", keterangan.strip()])
                    st.success(f"✅ {jenis} tercatat.")
                else:
                    st.warning("Anda sudah ada catatan absen hari ini. Hubungi admin untuk mengubah.")

# ============ MODE ADMIN ============
else:
    st.subheader("Dashboard Admin")
    pin_admin = st.text_input("PIN Admin", type="password")
    if pin_admin != ADMIN_PIN:
        if pin_admin: st.error("PIN admin salah.")
        st.stop()

    absensi_df = read_absensi()
    if absensi_df.empty:
        st.info("Belum ada data absensi.")
        st.stop()

    c1, c2, c3 = st.columns(3)
    with c1:
        tgl = st.date_input("Tanggal", value=date.today())
    with c2:
        divisi_list = sorted([x for x in karyawan_df["divisi"].dropna().unique() if clean(x)])
        div_filter = st.selectbox("Filter divisi", ["Semua"] + divisi_list)
    with c3:
        status_list = sorted([x for x in absensi_df["status"].dropna().unique() if clean(x)])
        stat_filter = st.selectbox("Filter status", ["Semua"] + status_list)

    df = absensi_df[absensi_df["tanggal"] == str(tgl)].copy()
    if div_filter != "Semua": df = df[df["divisi"] == div_filter]
    if stat_filter != "Semua": df = df[df["status"] == stat_filter]

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total catatan", len(df))
    m2.metric("Hadir", int((df["status"] == "Hadir").sum()))
    m3.metric("Terlambat", int((df["status"] == "Terlambat").sum()))
    m4.metric("Izin/Sakit/Cuti", int(df["status"].isin(["Izin","Sakit","Cuti"]).sum()))

    st.dataframe(df, use_container_width=True, hide_index=True)
    st.download_button("Download rekap (CSV)", data=df.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"rekap_absensi_{tgl}.csv", mime="text/csv")

st.divider()
st.caption("Batas masuk: 08:00 | Batas pulang: 17:00 | Data tersimpan di Google Sheets")
