import gspread
from datetime import date, time
from pathlib import Path
import pandas as pd
import streamlit as st
import json
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Transport Management System", page_icon="🚗", layout="wide")

#### google sheet --> https://docs.google.com/spreadsheets/d/1rsTGJS2H6uhqh84eQmlLn_uZVzopUWNxPqU_MpT9H5c/edit?usp=sharing

SHEET_ID = "1rsTGJS2H6uhqh84eQmlLn_uZVzopUWNxPqU_MpT9H5c"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
BASE_DIR = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "credentials.json"

TRANSPORT_COLUMNS = ["Request_ID","tgl_request","Requestor","Purpose","tgl_mulai","tgl_selesai","titik_jemput","titik_tujuan","jam_mulai","Est_waktu","penumpang","Contact Person","Driver","Car Type","Remark"]
DRIVER_COLUMNS = ["nama","hp","car_type","no_pol"]


def init_gspread():
    if "gcp_service_account" in st.secrets:
        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]), scopes=SCOPES
        )
    else:
        creds = Credentials.from_service_account_file(str(CREDENTIALS_FILE), scopes=SCOPES)
    return gspread.authorize(creds).open_by_key(SHEET_ID)

def find_ws(sh, title):
    target = title.strip().lower()
    for ws in sh.worksheets():
        if ws.title.strip().lower() == target:
            return ws
    return None

def get_or_create_ws(sh, title, rows, cols, headers):
    ws = find_ws(sh, title)
    if ws:
        return ws, False
    try:
        ws = sh.add_worksheet(title=title, rows=rows, cols=cols)
        ws.append_row(headers)
        return ws, True
    except gspread.exceptions.APIError as e:
        if "already exists" in str(e):
            ws = find_ws(sh, title)
            if ws: return ws, False
        raise

def read_transport():
    sh = init_gspread()
    ws, created = get_or_create_ws(sh, "transport_log", 100, len(TRANSPORT_COLUMNS), TRANSPORT_COLUMNS)
    if created:
        return pd.DataFrame(columns=TRANSPORT_COLUMNS)
    values = ws.get_all_values()
    if not values or len(values) < 2:
        if not values: ws.append_row(TRANSPORT_COLUMNS)
        return pd.DataFrame(columns=TRANSPORT_COLUMNS)
    headers = [str(v).strip() for v in values[0]]
    df = pd.DataFrame(values[1:], columns=headers)
    for col in TRANSPORT_COLUMNS:
        if col not in df.columns: df[col] = ""
    df = df[TRANSPORT_COLUMNS]
    mask_empty = df.drop(columns=["Request_ID"]).astype(str).apply(lambda c: c.str.strip()).eq("").all(axis=1)
    df = df[~mask_empty]
    used = set(df["Request_ID"].astype(str).str.strip())
    for idx in df.index:
        val = str(df.at[idx, "Request_ID"]).strip()
        if val == "" or val.lower() == "nan":
            n = 1
            while f"TR-{n:05d}" in used: n += 1
            df.at[idx, "Request_ID"] = f"TR-{n:05d}"
            used.add(df.at[idx, "Request_ID"])
    return df.reset_index(drop=True)

def read_drivers():
    sh = init_gspread()
    ws, created = get_or_create_ws(sh, "Driver", 50, len(DRIVER_COLUMNS), DRIVER_COLUMNS)
    if created:
        return pd.DataFrame(columns=DRIVER_COLUMNS)
    values = ws.get_all_values()
    if not values or len(values) < 2:
        if not values: ws.append_row(DRIVER_COLUMNS)
        return pd.DataFrame(columns=DRIVER_COLUMNS)
    headers = [str(v).strip() for v in values[0]]
    df = pd.DataFrame(values[1:], columns=headers)
    for col in DRIVER_COLUMNS:
        if col not in df.columns: df[col] = ""
    df = df[DRIVER_COLUMNS]
    df = df[df.astype(str).apply(lambda c: c.str.strip()).ne("").any(axis=1)]
    df["nama"] = df["nama"].astype(str).str.strip()
    return df.reset_index(drop=True)

def save_transport(df):
    sh = init_gspread()
    ws, _ = get_or_create_ws(sh, "transport_log", 100, len(TRANSPORT_COLUMNS), TRANSPORT_COLUMNS)
    ws.clear()
    ws.append_row(TRANSPORT_COLUMNS)
    out = df[TRANSPORT_COLUMNS].fillna("").astype(str).values.tolist()
    if out: ws.append_rows(out)

def next_request_id(df):
    existing = set(df["Request_ID"].astype(str))
    n = 1
    while f"TR-{n:05d}" in existing: n += 1
    return f"TR-{n:05d}"

def safe_date(value, fallback=None):
    if fallback is None: fallback = date.today()
    if value is None or pd.isna(value) or str(value).strip() == "": return fallback
    try: return pd.to_datetime(value).date()
    except Exception: return fallback

def safe_time(value, fallback=time(8, 0)):
    if value is None or pd.isna(value) or str(value).strip() == "": return fallback
    try:
        if isinstance(value, time): return value
        return pd.to_datetime(str(value)).time()
    except Exception: return fallback

def clean_text(value):
    if value is None or pd.isna(value): return ""
    return str(value).strip()

def driver_options(drivers):
    if drivers.empty: return ["Belum ditentukan"]
    return ["Belum ditentukan"] + sorted(drivers["nama"].dropna().unique().tolist())

def get_driver_vehicle_info(drivers, driver_name):
    if drivers.empty or not driver_name or driver_name == "Belum ditentukan": return ""
    row = drivers[drivers["nama"] == driver_name]
    if row.empty: return ""
    row = row.iloc[0]
    parts = [clean_text(row["car_type"]), clean_text(row["no_pol"])]
    return " / ".join([p for p in parts if p])

def make_form(defaults, drivers, key_prefix, request_id=None):
    def get(name, fallback=""):
        val = defaults.get(name, fallback)
        return "" if val is None or pd.isna(val) else val
    current_driver = clean_text(get("Driver", ""))
    options = driver_options(drivers)
    if current_driver and current_driver not in options: options.append(current_driver)
    if not current_driver: current_driver = "Belum ditentukan"
    tgl_request_val = safe_date(get("tgl_request"))
    start_date_val = safe_date(get("tgl_mulai"))
    end_date_val = safe_date(get("tgl_selesai"), fallback=start_date_val)
    if end_date_val < start_date_val: end_date_val = start_date_val
    with st.form(f"{key_prefix}_form", clear_on_submit=(key_prefix == "new")):
        st.caption(f"Nomor permintaan: {request_id or 'dibuat otomatis saat disimpan'}")
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown("**Informasi Permintaan**")
            tgl_request = st.date_input("Tanggal permintaan *", value=tgl_request_val, key=f"{key_prefix}_tgl_request")
            requestor = st.text_input("Pemohon *", value=clean_text(get("Requestor")), key=f"{key_prefix}_requestor")
            purpose = st.text_input("Keperluan *", value=clean_text(get("Purpose")), key=f"{key_prefix}_purpose")
            start_date = st.date_input("Tanggal mulai *", value=start_date_val, key=f"{key_prefix}_start_date")
            end_date = st.date_input("Tanggal selesai *", value=end_date_val, key=f"{key_prefix}_end_date", min_value=start_date_val)
        with c2:
            st.markdown("**Rencana Perjalanan**")
            pickup = st.text_input("Titik penjemputan *", value=clean_text(get("titik_jemput")), key=f"{key_prefix}_pickup")
            destination = st.text_input("Titik tujuan *", value=clean_text(get("titik_tujuan")), key=f"{key_prefix}_destination")
            start_time = st.time_input("Jam mulai *", value=safe_time(get("jam_mulai")), key=f"{key_prefix}_start_time")
            est_time = st.text_input("Estimasi waktu", value=clean_text(get("Est_waktu")), key=f"{key_prefix}_est_time")
            passengers = st.text_input("Penumpang", value=clean_text(get("penumpang")), key=f"{key_prefix}_passengers")
            contact = st.text_input("Contact person / telepon", value=clean_text(get("Contact Person")), key=f"{key_prefix}_contact")
        with c3:
            st.markdown("**Penugasan & Status**")
            car_type_default = clean_text(get("Car Type"))
            if not car_type_default and current_driver != "Belum ditentukan":
                car_type_default = get_driver_vehicle_info(drivers, current_driver)
            remark_list = ["Open","Assigned","In Progress","Closed","Cancelled"]
            rv = clean_text(get("Remark"))
            remark = st.selectbox("Status / remark", options=remark_list, index=remark_list.index(rv) if rv in remark_list else 0, key=f"{key_prefix}_remark")
            driver = st.selectbox("Pilih Driver", options=options, index=options.index(current_driver) if current_driver in options else 0, key=f"{key_prefix}_driver")
            if driver != "Belum ditentukan":
                vi = get_driver_vehicle_info(drivers, driver)
                if vi: st.info(f"🚗 Kendaraan: **{vi}**")
            car_type = st.text_input("Jenis kendaraan / nomor polisi", value=car_type_default, key=f"{key_prefix}_car_type", help="Terisi otomatis dari data driver, dapat diubah manual")
        submitted = st.form_submit_button("Simpan permintaan" if key_prefix == "new" else "Simpan perubahan", type="primary", use_container_width=True)
        if not submitted: return None
        if end_date < start_date:
            st.error("Tanggal selesai tidak boleh lebih awal dari tanggal mulai."); return None
        result = {"Request_ID": request_id or "", "tgl_request": str(tgl_request), "Requestor": requestor.strip(),
                  "Purpose": purpose.strip(), "tgl_mulai": str(start_date), "tgl_selesai": str(end_date),
                  "titik_jemput": pickup.strip(), "titik_tujuan": destination.strip(), "jam_mulai": str(start_time),
                  "Est_waktu": est_time.strip(), "penumpang": passengers.strip(), "Contact Person": contact.strip(),
                  "Driver": "" if driver == "Belum ditentukan" else driver, "Car Type": car_type.strip(), "Remark": remark}
        missing = [f for f in ["Requestor","Purpose","titik_jemput","titik_tujuan"] if not result[f]]
        if missing:
            st.error("Mohon isi kolom wajib: " + ", ".join(missing)); return None
        return result

def show_dashboard(df):
    st.subheader("Dashboard monitoring")
    if df.empty:
        st.info("Belum ada permintaan transportasi. Tambahkan permintaan melalui menu Permintaan Baru."); return
    dv = pd.to_datetime(df["tgl_mulai"], errors="coerce").dropna()
    default_date = date.today(); min_date = max_date = None
    if not dv.empty:
        min_date = dv.min().date(); max_date = dv.max().date()
        default_date = max(min_date, min(default_date, max_date))
    f1, f2, f3 = st.columns([1,1,1])
    with f1:
        show_all = st.checkbox("Tampilkan Semua Tanggal", value=False)
        selected_date = None if show_all else st.date_input("Tanggal perjalanan", value=default_date, min_value=min_date, max_value=max_date, key="dash_date")
    with f2:
        drv = sorted([x for x in df["Driver"].dropna().astype(str).unique() if x.strip()])
        sel_driver = st.selectbox("Filter driver", ["Semua driver"] + drv, key="dash_driver")
    with f3:
        sts = sorted([x for x in df["Remark"].dropna().astype(str).unique() if x.strip()])
        sel_status = st.selectbox("Filter status", ["Semua status"] + sts, key="dash_status")
    if selected_date is not None:
        starts = pd.to_datetime(df["tgl_mulai"], errors="coerce").dt.date
        ends = pd.to_datetime(df["tgl_selesai"], errors="coerce").dt.date
        day_df = df[(starts <= selected_date) & (ends >= selected_date)].copy()
    else:
        day_df = df.copy()
    if sel_driver != "Semua driver": day_df = day_df[day_df["Driver"].astype(str) == sel_driver]
    if sel_status != "Semua status": day_df = day_df[day_df["Remark"].astype(str) == sel_status]
    total = len(day_df)
    assigned = int(day_df["Driver"].fillna("").astype(str).str.strip().ne("").sum())
    closed = int(day_df["Remark"].fillna("").astype(str).str.lower().eq("closed").sum())
    m1,m2,m3,m4 = st.columns(4)
    m1.metric("Total permintaan", total); m2.metric("Sudah ada driver", assigned)
    m3.metric("Belum ada driver", total-assigned); m4.metric("Closed", closed)
    display_cols = ["Request_ID","tgl_mulai","jam_mulai","Requestor","Purpose","titik_jemput","titik_tujuan","penumpang","Driver","Car Type","Remark"]
    if day_df.empty:
        st.info("Tidak ada perjalanan yang sesuai filter.")
    else:
        ds = day_df[display_cols].copy()
        ds["tgl_mulai"] = ds["tgl_mulai"].astype(str); ds["jam_mulai"] = ds["jam_mulai"].astype(str).str.slice(0,5)
        ds = ds.rename(columns={"Request_ID":"No. Permintaan","tgl_mulai":"Tanggal Mulai","jam_mulai":"Jam","Requestor":"Pemohon","Purpose":"Keperluan","titik_jemput":"Penjemputan","titik_tujuan":"Tujuan","penumpang":"Penumpang","Car Type":"Kendaraan","Remark":"Status"})
        st.dataframe(ds, use_container_width=True, hide_index=True)
        suffix = selected_date.isoformat() if selected_date else "semua-tanggal"
        st.download_button(f"Download daftar perjalanan ({suffix})", data=day_df[display_cols].to_csv(index=False).encode("utf-8-sig"), file_name=f"transport_{suffix}.csv", mime="text/csv")
    adf = day_df[day_df["Driver"].fillna("").astype(str).str.strip().ne("")]
    if not adf.empty:
        clashes = adf.groupby(["Driver","tgl_mulai","jam_mulai"]).size(); clashes = clashes[clashes > 1]
        if not clashes.empty:
            st.warning("⚠️ Ditemukan bentrok jadwal untuk driver, tanggal & jam yang sama.")
            st.dataframe(clashes.rename("Jumlah permintaan").reset_index(), use_container_width=True)

st.title("🚗 Transport Management System")
st.caption("Data tersimpan di Google Drive — akun akholil.chrysalis@gmail.com")
try:
    data = read_transport(); drivers_df = read_drivers()
except Exception as exc:
    st.error(f"Tidak dapat terhubung ke Google Sheets: {exc}"); st.stop()
with st.sidebar:
    st.header("Menu")
    page = st.radio("Pilih halaman", ["Dashboard","Permintaan Baru","Edit Permintaan","Data Permintaan"], label_visibility="collapsed")
    st.divider(); st.caption("Sumber data: Google Sheets")
    st.caption(f"Jumlah permintaan: {len(data)}"); st.caption(f"Jumlah driver terdaftar: {len(drivers_df)}")
if page == "Dashboard":
    show_dashboard(data)
elif page == "Permintaan Baru":
    st.subheader("Form permintaan transportasi")
    rec = make_form({}, drivers_df, "new")
    if rec:
        rec["Request_ID"] = next_request_id(data)
        updated = pd.concat([data, pd.DataFrame([rec])], ignore_index=True)
        try:
            save_transport(updated); st.success(f"✅ Permintaan {rec['Request_ID']} berhasil disimpan."); st.rerun()
        except Exception as exc: st.error(f"❌ Gagal menyimpan data: {exc}")
elif page == "Edit Permintaan":
    st.subheader("Edit permintaan transportasi")
    if data.empty:
        st.info("Belum ada data yang dapat diedit.")
    else:
        choices = data["Request_ID"].astype(str).tolist()
        labels = {str(r["Request_ID"]): f"{r['Request_ID']} | {clean_text(r['Requestor'])} | {clean_text(r['Purpose'])} | {clean_text(r['tgl_mulai'])}" for _, r in data.iterrows()}
        selected_id = st.selectbox("Pilih permintaan", choices, format_func=lambda x: labels.get(x, x))
        sel = data[data["Request_ID"].astype(str) == selected_id]
        if not sel.empty:
            edited = make_form(sel.iloc[0].to_dict(), drivers_df, f"edit_{selected_id}", request_id=selected_id)
            if edited:
                updated = data.copy()
                ri = updated.index[updated["Request_ID"].astype(str) == selected_id][0]
                for k, v in edited.items(): updated.at[ri, k] = v
                try:
                    save_transport(updated); st.success(f"✅ Permintaan {selected_id} berhasil diperbarui."); st.rerun()
                except Exception as exc: st.error(f"❌ Gagal menyimpan perubahan: {exc}")
elif page == "Data Permintaan":
    st.subheader("Seluruh data permintaan")
    if data.empty:
        st.info("Belum ada permintaan transportasi.")
    else:
        search = st.text_input("Cari nomor, pemohon, keperluan, tujuan, atau driver")
        filtered = data.copy()
        if search.strip():
            mask = filtered.astype(str).apply(lambda c: c.str.contains(search.strip(), case=False, na=False)).any(axis=1)
            filtered = filtered[mask]
        st.dataframe(filtered, use_container_width=True, hide_index=True)
        st.download_button("Download seluruh data (CSV)", data=filtered.to_csv(index=False).encode("utf-8-sig"), file_name="transport_log.csv", mime="text/csv")
st.divider()
st.caption("Data tersimpan otomatis ke Google Sheets — tidak perlu menutup file Excel lagi.")