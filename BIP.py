# app.py
import streamlit as st
import pandas as pd
import io, zipfile
import plotly.express as px
import csv
from datetime import datetime, time, timedelta
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER

# ======================
# PENGATURAN STREAMLIT LAYOUT
# ======================
st.set_page_config(page_title="Absensi PT BIP", layout="centered")
st.markdown("<h1 style='text-align:center;'>📋 Aplikasi Absensi Outsoucing di PT. Japfa Comfeed Indonesia Tbk.</h1>", unsafe_allow_html=True)
st.write("")

# ======================
# UTILS
# ======================
def parse_datetime_flexible(val):
    if pd.isna(val):
        return pd.NaT
    val = str(val).strip()
    if not val:
        return pd.NaT
    if " " in val:
        date_part, time_part = val.split(" ", 1)
        time_part = time_part.replace(".", ":")
        val = f"{date_part} {time_part}"
    else:
        val = val.replace(".", "-")
    fmts = [
        "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%y %H:%M:%S", "%d/%m/%y %H:%M",
        "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M",
        "%d-%m-%Y %H:%M:%S", "%d-%m-%Y %H:%M", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d", "%Y-%m-%d"
    ]
    for fmt in fmts:
        try:
            return datetime.strptime(val, fmt)
        except:
            continue
    return pd.to_datetime(val, errors="coerce", dayfirst=True)

def clean_and_normalize(df):
    # basic column cleanup + mapping
    df.columns = [str(c).strip() for c in df.columns]
    rename_map = {
        "No.ID": "ID", "No ID": "ID", "No. ID": "ID", "NIP": "ID", "No": "ID", "NO": "ID",
        "Tgl/Waktu": "Tanggal_Waktu", "Tgl / Waktu": "Tanggal_Waktu", "Tanggal": "Tanggal_Waktu", "TANGGAL": "Tanggal_Waktu", "Waktu": "Tanggal_Waktu", 
        "WAKTU": "Tanggal_Waktu",
        "Lokasi ID": "Lokasi_ID", "Lokasi": "Lokasi_ID", "LokasiID": "Lokasi_ID",
        "Karyawan": "Nama", "KARYAWAN": "Nama", "NAMA": "Nama"
    }
    df.rename(columns={k:v for k,v in rename_map.items() if k in df.columns}, inplace=True)

    possible_cols = ["Nama", "ID", "Tanggal_Waktu", "Lokasi_ID"]
    existing = [c for c in possible_cols if c in df.columns]
    df = df[existing].copy()

    if "Tanggal_Waktu" in df.columns:
        df["Tanggal_Waktu"] = df["Tanggal_Waktu"].apply(parse_datetime_flexible)
        df["Tanggal"] = df["Tanggal_Waktu"].dt.date
        df["Waktu"] = df["Tanggal_Waktu"].dt.time
    if "Nama" in df.columns:
        df["Nama"] = df["Nama"].astype(str).str.strip()
    if "ID" in df.columns:
        df["ID"] = df["ID"].astype(str).str.strip()
    # drop rows without Nama or Tanggal_Waktu
    df = df.dropna(subset=["Nama", "Tanggal_Waktu"], how="any")
    return df

def _to_time_obj(t):
    if pd.isnull(t):
        return None
    if isinstance(t, time):
        return t
    try:
        return pd.to_datetime(t).time()
    except:
        return None

def overlaps(a_start, a_end, b_start, b_end, min_minutes=60):
    latest_start = max(a_start, b_start)
    earliest_end = min(a_end, b_end)
    delta = (earliest_end - latest_start).total_seconds() / 60.0
    return delta >= min_minutes

def encode_shifts(cek_in, cek_out, tanggal, shift_hours=8, tolerance_minutes=60):
    if pd.isnull(cek_in) and pd.isnull(cek_out):
        return None, None, None
    base_date = pd.to_datetime(tanggal).date()
    t_in = _to_time_obj(cek_in)
    t_out = _to_time_obj(cek_out)
    start = datetime.combine(base_date, t_in) if t_in else None
    end = datetime.combine(base_date, t_out) if t_out else None
    if start is None and end is not None:
        start = end - timedelta(hours=shift_hours)
    if end is None and start is not None:
        end = start + timedelta(hours=shift_hours)
    if end < start:
        end += timedelta(days=1)
    s1_start = datetime.combine(base_date, time(7,0))
    s1_end   = datetime.combine(base_date, time(15,0))
    s2_start = datetime.combine(base_date, time(15,0))
    s2_end   = datetime.combine(base_date, time(23,0))
    s3_start = datetime.combine(base_date, time(23,0))
    s3_end   = datetime.combine(base_date + timedelta(days=1), time(7,0))
    s1 = 1 if overlaps(start, end, s1_start, s1_end, min_minutes=tolerance_minutes) else 0
    s2 = 1 if overlaps(start, end, s2_start, s2_end, min_minutes=tolerance_minutes) else 0
    s3 = 1 if overlaps(start, end, s3_start, s3_end, min_minutes=tolerance_minutes) else 0
    if (s1, s2, s3) == (0,0,0) and start is not None:
        midpoints = {
            1: s1_start + (s1_end - s1_start)/2,
            2: s2_start + (s2_end - s2_start)/2,
            3: s3_start + (s3_end - s3_start)/2,
        }
        diffs = {k: abs((start - v).total_seconds()) for k,v in midpoints.items()}
        nearest = min(diffs, key=diffs.get)
        if nearest == 1: s1=1
        elif nearest == 2: s2=1
        else: s3=1
    if cek_in is None and t_out is not None and t_out < time(7,0):
        s1, s2, s3 = 0,0,1
    return s1, s2, s3

def hari_indonesia(nama_hari):
    mapping = {
        "Monday":"Senin","Tuesday":"Selasa","Wednesday":"Rabu","Thursday":"Kamis",
        "Friday":"Jumat","Saturday":"Sabtu","Sunday":"Minggu"
    }
    return mapping.get(nama_hari, nama_hari)

def safe_text(val):
    if pd.isna(val) or str(val).strip().lower() in ["nan", "none"]:
        return ""
    return str(val).strip()

def export_pdf_per_tanggal(df, tanggal):
    if df.empty:
        return None

    buffer = io.BytesIO()
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name='CenterBoldTitle',
        alignment=TA_CENTER,
        fontSize=12,
        leading=14,
        spaceAfter=12,
        fontName="Helvetica-Bold"
    ))

    tanggal_dt = pd.to_datetime(tanggal)
    hari = hari_indonesia(tanggal_dt.strftime("%A")).upper()
    tgl_str = tanggal_dt.strftime("%d %B %Y").upper()

    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        leftMargin=25, rightMargin=25, topMargin=25, bottomMargin=25
    )
    elements = []

    # === Judul ===
    judul = (
        f"<b>ABSENSI PT. BUDI INTI PERKASA</b><br/>"
        f"<b>HARI {hari} - TANGGAL {tgl_str}</b>"
    )
    elements.append(Paragraph(judul, styles["CenterBoldTitle"]))
    elements.append(Spacer(1, 10))

    # === Tabel utama ===
    data = [["NO", "NIP", "NAMA PEKERJA", "KEGIATAN",
             "SHIFT 1", "SHIFT 2", "SHIFT 3", "CEK IN", "CEK OUT"]]

    for j, row in enumerate(df.itertuples(), start=1):
        s1 = "✔" if getattr(row, "Shift1", 0) == 1 else ""
        s2 = "✔" if getattr(row, "Shift2", 0) == 1 else ""
        s3 = "✔" if getattr(row, "Shift3", 0) == 1 else ""
        nama_cap = safe_text(row.Nama).title()
        kegiatan_cap = safe_text(row.Kegiatan)

        data.append([
            j,
            safe_text(row.ID),
            nama_cap,
            kegiatan_cap,
            s1, s2, s3,
            safe_text(row.Cek_In),
            safe_text(row.Cek_Out)
        ])

    col_widths = [25, 40, 120, 95, 45, 45, 45, 55, 55]

    table = Table(data, repeatRows=1, colWidths=col_widths)
    table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightblue),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("ALIGN", (2, 1), (2, -1), "LEFT"),
        ("ALIGN", (3, 1), (3, -1), "LEFT"),
        ("ROWHEIGHT", (0, 1), (-1, -1), 12),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 12))

    # === Ringkasan total keseluruhan ===
    total_s1 = df["Shift1"].sum()
    total_s2 = df["Shift2"].sum()
    total_s3 = df["Shift3"].sum()
    total_all = int(total_s1 + total_s2 + total_s3)

    summary_data = [
        ["Shift 1", "Shift 2", "Shift 3", "Total Pekerja"],
        [int(total_s1), int(total_s2), int(total_s3), total_all]
    ]
    summary_table = Table(summary_data, colWidths=[80, 80, 80, 100])
    summary_table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
    ]))
    elements.append(summary_table)
    elements.append(Spacer(1, 36))

    doc.build(elements)
    buffer.seek(0)
    return buffer

# ======================
# UI - Uploads + Validasi 
# ======================
import csv

def read_any_csv(uploaded_file):
    try:
        # Baca sedikit sample untuk deteksi delimiter
        sample = uploaded_file.read(4096).decode("utf-8", errors="ignore")
        uploaded_file.seek(0)

        dialect = csv.Sniffer().sniff(sample, delimiters=";,|\t")
        delimiter = dialect.delimiter

        uploaded_file.seek(0)
        return pd.read_csv(uploaded_file, sep=delimiter, engine="python")

    except Exception:
        # Jika gagal deteksi, coba pakai titik koma IEC/Indonesia (;)
        try:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, sep=";", engine="python")
        except:
            # Jika tetap gagal → fallback koma (,)
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, sep=",", engine="python")

st.write("### Upload file absensi mentah (.csv / .xlsx / .xls)")

col_upload, col_info = st.columns([12, 1])

with col_upload:
    uploaded_file = st.file_uploader("", type=["csv", "xlsx", "xls"], label_visibility="collapsed")

with col_info:
    # tombol popover
    with st.popover("❗", use_container_width=True):
        st.markdown("""
**FORMAT DATA ABSENSI MENTAH**
| Kolom     | Keterangan |
|----------|------------|
| ID / NIP | Nomor Induk Karyawan |
| Nama     | Nama lengkap karyawan |
| Kegiatan | Kegiatan / Unit kerja |
| Tanggal  | Format tanggal absensi |

---

**FORMAT MASTER DATA**
| Kolom     | Contoh Isi |
|----------|-------------|
| ID       | 3 |
| Nama     | Suindro |
| Status   | PKWT / PKWTT |
| Kegiatan | Bongkaran / Silo / dll |
""")

st.markdown("""
<style>

/* Turunkan / naikkan tombol ❗ secara paksa */
button[data-testid="stPopoverButton"] {
    transform: translateY(20px) !important;   /* naik/turun: ubah angka ini */
    height: 42px !important;                /* samakan tinggi */
    width: 42px !important;                 /* kalau mau kotak */
    padding: 0 !important;
    border-radius: 6px !important;

    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
}

/* Pastikan uploader tidak runtuh */
div[data-testid="stFileUploader"] button {
    min-height: 42px !important;
}

/* Centering kolom */
div[data-testid="column"] {
    display: flex !important;
    align-items: center !important;
}

</style>
""", unsafe_allow_html=True)

st.write("")  
use_default_master = st.checkbox("Gunakan master data default (MasterData.csv)", value=True)

master_df = None
if not use_default_master:
    master_upload = st.file_uploader("Upload master data (.csv / .xlsx / .xls)", type=["csv", "xlsx", "xls"], key="master")
    if master_upload is None:
        st.warning("Unggah master data terlebih dahulu atau gunakan default.")
        st.stop()
    else:
        fname = master_upload.name.lower()
        try:
            if fname.endswith(".csv"):
                master_df = read_any_csv(master_upload)
            elif fname.endswith(".xlsx"):
                master_df = pd.read_excel(master_upload, engine="openpyxl")
            elif fname.endswith(".xls"):
                master_df = pd.read_excel(master_upload, engine="xlrd")
            else:
                st.error("Format master tidak didukung. Gunakan .csv .xlsx .xls")
                st.stop()

            # ✅ SIMPAN (overwrite) master baru agar PERSISTEN
            master_df.to_csv("MasterData.csv", index=False, sep=";")
            st.success("Master data berhasil diperbarui dan disimpan permanen!")
            
        except Exception as e:
            st.error(f"Gagal membaca file master: {e}")
            master_df = None
else:
    try:
        master_df = pd.read_csv("MasterData.csv", delimiter=";")
    except Exception:
        st.warning("masterBIP.csv tidak ditemukan di folder aplikasi.")
        master_df = None

if uploaded_file is None:
    st.info("Silakan upload file absensi mentah (.csv / .xlsx / .xls).")
    st.stop()

fname = uploaded_file.name.lower()
try:
    if fname.endswith(".csv"):
        df_raw = read_any_csv(uploaded_file)
    elif fname.endswith(".xlsx"):
        df_raw = pd.read_excel(uploaded_file, engine="openpyxl")
    elif fname.endswith(".xls"):
        df_raw = pd.read_excel(uploaded_file, engine="xlrd")
    else:
        st.error("Format file tidak didukung. Gunakan .csv / .xlsx / .xls.")
        st.stop()

except Exception as e:
    st.error(f"Gagal membaca file absensi: {e}")
    st.stop()

# ======================
# VALIDASI FORMAT FILE ABSENSI (USING clean_and_normalize)
# ======================
try:
    df_clean = clean_and_normalize(df_raw)
    if "Nama" not in df_clean.columns or "Tanggal_Waktu" not in df_clean.columns:
        st.error("❌ Format file absensi tidak sesuai. Pastikan kolom Nama dan Tanggal/Waktu tersedia.\n\nPastikan Format Data Sesuai.")
        st.stop()
    if df_clean.empty:
        st.error("❌ File absensi setelah pembersihan menghasilkan data kosong. Pastikan file benar.")
        st.stop()
except Exception as e:
    st.error(f"❌ Gagal memproses file absensi: {e}\n\nPastikan Format Data Sesuai.")
    st.stop()

# ======================
# VALIDASI MASTER DATA (CASE INSENSITIVE, TOLERAN)
# ======================
if master_df is not None:
    # normalize headers
    master_df.columns = master_df.columns.str.strip()
    master_cols_upper = [c.upper() for c in master_df.columns]

    # required: ID and NAMA
    if not any(c.upper() == "ID" for c in master_df.columns) and not any(c.upper() == "NIP" for c in master_df.columns):
        st.error("❌ Master data tidak memiliki kolom 'ID' atau 'NIP'.\n\nPastikan Format Data Sesuai.")
        st.stop()
    if not any(c.upper() == "NAMA" for c in master_df.columns):
        st.error("❌ Master data tidak memiliki kolom 'NAMA'.\n\nPastikan Format Data Sesuai.")
        st.stop()

    id_col = None
    for c in master_df.columns:
        if c.strip().upper() in ["ID", "NIP"]:
            id_col = c
            break
    name_col = None
    for c in master_df.columns:
        if c.strip().upper() == "NAMA":
            name_col = c
            break
    status_col = None
    for c in master_df.columns:
        if c.strip().upper() == "STATUS":
            status_col = c
            break
    kegiatan_col = None
    for c in master_df.columns:
        if c.strip().upper() in ["KEGIATAN", "KEGIATAN "]:
            kegiatan_col = c
            break

    # create normalized master_df with columns ID, Nama, STATUS, KEGIATAN
    master_norm = pd.DataFrame()
    master_norm["ID"] = master_df[id_col].astype(str).str.strip() if id_col else ""
    master_norm["Nama"] = master_df[name_col].astype(str).str.strip() if name_col else ""
    master_norm["Status"] = master_df[status_col].astype(str).str.strip() if status_col else ""
    if kegiatan_col:
        master_norm["Kegiatan"] = master_df[kegiatan_col].astype(str).str.strip()
    else:
        master_norm["Kegiatan"] = ""
    # remove accidental header-like rows etc
    master_df = master_norm.copy()
else:
    # no master data loaded: continue but warn
    st.warning("⚠ Tidak ada Master Data dimuat — proses akan lanjut tanpa informasi Status/Kegiatan master.")

# ======================
# PROSES UTAMA (SAMA DENGAN ALUR KAMU)
# ======================
try:
    # tentukan Cek_In / Cek_Out berdasarkan Lokasi_ID bila ada
    if "Lokasi_ID" in df_clean.columns:
        df_clean["Cek_In"]  = df_clean.apply(lambda x: x["Waktu"] if x["Lokasi_ID"] == 2 else None, axis=1)
        df_clean["Cek_Out"] = df_clean.apply(lambda x: x["Waktu"] if x["Lokasi_ID"] == 1 else None, axis=1)
    else:
        df_clean["Cek_In"] = df_clean["Waktu"]
        df_clean["Cek_Out"] = df_clean["Waktu"]

    result = df_clean.groupby(["ID","Nama","Tanggal"]).agg({
        "Cek_In": lambda x: min([t for t in x if t is not None], default=None),
        "Cek_Out": lambda x: max([t for t in x if t is not None], default=None)
    }).reset_index()

    # standar nama di result supaya matching master by name works
    result["Nama"] = result["Nama"].astype(str).str.strip().str.upper()

    # merge kegiatan if master available (try matching on Nama)
    if master_df is not None and "Nama" in master_df.columns:
        master_df["Nama_up"] = master_df["Nama"].astype(str).str.strip().str.upper()
        # if master has Kegiatan, use it
        if "Kegiatan" in master_df.columns:
            master_df["Kegiatan"] = master_df["Kegiatan"].fillna("")
            master_df["Kegiatan_up"] = master_df["Kegiatan"].astype(str).str.strip()
            # merge on Nama
            final_result = pd.merge(result, master_df[["Nama_up","Kegiatan_up"]].rename(columns={"Nama_up":"Nama","Kegiatan_up":"KEGIATAN"}), on="Nama", how="left")
            final_result.rename(columns={"KEGIATAN":"Kegiatan"}, inplace=True)
        else:
            final_result = result.copy()
            final_result["Kegiatan"] = pd.NA
    else:
        final_result = result.copy()
        final_result["Kegiatan"] = pd.NA

    # encode shifts
    final_result[["Shift1","Shift2","Shift3"]] = final_result.apply(
        lambda r: pd.Series(encode_shifts(r["Cek_In"], r["Cek_Out"], r["Tanggal"], shift_hours=8)),
        axis=1
    )

    st.success("✅ Data absensi berhasil diproses.")

    # ======================
    # BAGIAN UNDUH (PDF / BULANAN ZIP / CSV REKAP)
    # ======================
    tanggal_all = sorted(final_result["Tanggal"].dropna().unique())
    if len(tanggal_all) > 0:
        tanggal_pilih = st.selectbox("Pilih tanggal untuk unduh PDF harian:", tanggal_all)
        col1, col2 = st.columns(2)

        # ==== HARiAN ====
        df_harian = final_result[final_result["Tanggal"] == tanggal_pilih]
        if not df_harian.empty:
            pdf_buf = export_pdf_per_tanggal(df_harian, tanggal_pilih)
        else:
            pdf_buf = None

        if pdf_buf:
            with col1:
                st.download_button(
                    label=f"⬇️ Unduh Harian (PDF)",
                    data=pdf_buf,
                    file_name=f"absen_{tanggal_pilih}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
        else:
            with col1:
                st.warning("Tidak ada data kegiatan valid pada tanggal ini.")

        # ==== BULANAN ZIP ====
        mem_zip = io.BytesIO()
        with zipfile.ZipFile(mem_zip, mode="w") as zf:
            for tgl in tanggal_all:
                df_day = final_result[final_result["Tanggal"] == tgl]
                if df_day.empty:
                    continue
                pdfb = export_pdf_per_tanggal(df_day, tgl)
                if pdfb:
                    fname = f"absen_{tgl}.pdf".replace("/", "-")
                    zf.writestr(fname, pdfb.read())
        mem_zip.seek(0)

        with col2:
            st.download_button(
                label="⬇️ Unduh Bulanan (ZIP)",
                data=mem_zip.getvalue(),
                file_name="rekap_absensi_bulanan.zip",
                mime="application/zip",
                use_container_width=True
            )

        # ==== REKAP BULANAN CSV ====
        st.write("")
        st.markdown("### 📊 Unduh Rekap Bulanan (CSV)")

        # ensure datetime
        final_result = final_result.copy()
        final_result["Tanggal"] = pd.to_datetime(final_result["Tanggal"])

        # month days
        first_date = final_result["Tanggal"].min()
        month_start = first_date.replace(day=1)
        month_end = (first_date + pd.offsets.MonthEnd(0)).normalize()
        month_days = pd.date_range(month_start, month_end)
        day_cols = [d.day for d in month_days]  # list 1..28/29/30/31

        # presence and day
        tmp = final_result.copy()
        tmp["Hadir"] = tmp[["Shift1", "Shift2", "Shift3"]].sum(axis=1) > 0
        tmp["day"] = tmp["Tanggal"].dt.day

        # normalize ID types for safe merge with master by ID
        if "ID" in tmp.columns:
            tmp["ID"] = tmp["ID"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()

        # normalize master_df ID to string as well
        if master_df is not None and "ID" in master_df.columns:
            master_df["ID"] = master_df["ID"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()

        # add Status column from master if available (try by ID first)
        if master_df is not None:
            if "ID" in master_df.columns and "Status" in master_df.columns:
                # use ID-based merge
                tmp = tmp.merge(master_df[["ID", "Status"]].rename(columns={"ID":"ID"}), on="ID", how="left")
                tmp["Status"] = tmp["Status"].fillna("")
            elif "Nama" in master_df.columns and "Status" in master_df.columns:
                # fallback to name-based
                tmp["Nama_up"] = tmp["Nama"].astype(str).str.strip().str.upper()
                master_df["Nama_up"] = master_df["Nama"].astype(str).str.strip().str.upper()
                tmp = tmp.merge(master_df[["Nama_up","Status"]].rename(columns={"Nama_up":"Nama"}), on="Nama", how="left")
                tmp["Status"] = tmp["Status"].fillna("")
            else:
                tmp["Status"] = ""
        else:
            tmp["Status"] = ""

        # pivot table (index includes status so it will be carried to rekap)
        rekap = (
            tmp.pivot_table(
                index=["ID", "Nama", "Kegiatan", "Status"],
                columns="day",
                values="Hadir",
                aggfunc="max",
                fill_value=False
            )
            .reset_index()
        )

        # ensure all days present
        for d in day_cols:
            if d not in rekap.columns:
                rekap[d] = False

        # order columns
        ordered = ["ID", "Nama", "Kegiatan", "Status"] + day_cols
        rekap = rekap[[c for c in ordered if c in rekap.columns]]
        # sort rows
        rekap = rekap.sort_values(["Nama", "ID"]).reset_index(drop=True)

        for d in day_cols:
            rekap[d] = rekap[d].apply(lambda x: "✔" if bool(x) else "")
        rekap["Total"] = rekap[day_cols].apply(lambda row: sum(1 for v in row if v == "✔"), axis=1)

        # clean NaN
        rekap["Kegiatan"] = rekap["Kegiatan"].fillna("")
        rekap["Status"] = rekap["Status"].fillna("")

        # rename ID -> NIP
        rekap.rename(columns={"ID": "NIP"}, inplace=True)
        rekap_csv = rekap.copy()
        rekap_csv = rekap_csv.replace("✔", "v")
        rekap_csv["Nama"] = rekap_csv["Nama"].astype(str).str.title()

        # save to buffer with utf-8-sig and semicolon
        csv_buf = io.StringIO()
        rekap_csv.to_csv(csv_buf, index=False, sep=";", encoding="utf-8-sig")
        csv_buf.seek(0)

        st.download_button(
            label="⬇️ Unduh Rekap Bulanan (CSV)",
            data=csv_buf.getvalue(),
            file_name=f"rekap_absensi_{month_start.strftime('%Y_%m')}.csv",
            mime="text/csv",
            use_container_width=True
        )

        # ======================
        # DASHBOARD (tampil setelah proses berhasil)
        # ======================
        st.write("")
        st.markdown("### 📈 Dashboard Analisis Absensi Bulanan")

        # ensure final_result exists
        final_result["Tanggal"] = pd.to_datetime(final_result["Tanggal"])

        # 1) line: jumlah karyawan unik per tanggal
        final_result["Tanggal"] = pd.to_datetime(final_result["Tanggal"])
        bulan_nama = final_result["Tanggal"].dt.strftime("%B %Y").iloc[0]
        df_line = final_result.groupby("Tanggal")["ID"].nunique().reset_index(name="Jumlah Karyawan")
        df_line["Hari"] = df_line["Tanggal"].dt.day

        # Buat grafik
        fig_line = px.line(
            df_line,
            x="Hari",
            y="Jumlah Karyawan",
            markers=True,
            title=f"Jumlah Karyawan Bulan {bulan_nama}"
        )

        fig_line.update_layout(
            xaxis=dict(
                tickmode="linear",
                dtick=1,
                title="Tanggal"
            ),
            yaxis_title="Jumlah Karyawan"
        )

        st.plotly_chart(fig_line, use_container_width=True)

        # 2) bar kegiatan
        df_kegiatan = final_result.groupby("Kegiatan")["ID"].nunique().reset_index(name="Jumlah Karyawan").sort_values("Jumlah Karyawan", ascending=False)
        fig_bar_kegiatan = px.bar(df_kegiatan, x="Kegiatan", y="Jumlah Karyawan", text="Jumlah Karyawan", title="Jumlah Karyawan per Kegiatan")
        fig_bar_kegiatan.update_traces(textposition="outside")
        st.plotly_chart(fig_bar_kegiatan, use_container_width=True)

        # 3) status horizontal bar (from rekap)
        if "Status" in rekap.columns:
            df_status = rekap.groupby("Status")["NIP"].nunique().reset_index(name="Jumlah Karyawan").sort_values("Jumlah Karyawan", ascending=True)
        else:
            df_status = pd.DataFrame(columns=["Status", "Jumlah Karyawan"])
        fig_bar_status = px.bar(df_status, x="Jumlah Karyawan", y="Status", orientation="h", text="Jumlah Karyawan", title="Jumlah Karyawan per Status")
        fig_bar_status.update_traces(textposition="outside")
        st.plotly_chart(fig_bar_status, use_container_width=True)

        # 4) pie shift
        shift_totals = final_result[["Shift1","Shift2","Shift3"]].sum().reset_index()
        shift_totals.columns = ["Shift","Jumlah"]
        shift_totals["Shift"] = shift_totals["Shift"].str.replace("Shift", "Shift ")
        fig_pie = px.pie(shift_totals, names="Shift", values="Jumlah", title="Distribusi Jumlah Karyawan per Shift", hole=0.3)
        st.plotly_chart(fig_pie, use_container_width=True)

    else:
        st.warning("Tidak ditemukan tanggal valid dalam data absensi.")
except Exception as e:
    st.error(f"Terjadi kesalahan saat membaca/ memproses data: {e}")
