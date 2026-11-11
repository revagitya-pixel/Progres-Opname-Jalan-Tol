import os, sqlite3, uuid, shutil
from pathlib import Path
from datetime import date, datetime, timedelta
from io import BytesIO

import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont, ImageOps
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter

# ==============================================================
# PROGRES OPNAME JALAN TOL V4.1 - penyempurnaan V4
# Tetap menggunakan tabel v4_* agar kompatibel dengan database V4.
# ==============================================================

st.set_page_config(page_title="Progres Opname Jalan Tol V4.1",
                   page_icon="🏗️", layout="wide",
                   initial_sidebar_state="expanded")

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("PROGRES_DATA_DIR", str(APP_DIR / "data")))
PHOTO_DIR = Path(os.environ.get("PROGRES_PHOTO_DIR", str(DATA_DIR / "foto")))
DB_PATH = Path(os.environ.get("PROGRES_DB", str(DATA_DIR / "progres_v4.db")))
DATA_DIR.mkdir(parents=True, exist_ok=True)
PHOTO_DIR.mkdir(parents=True, exist_ok=True)

NAVY = "17365D"
BLUE = "D9EAF7"
LIGHT = "F3F6FA"
GREEN = "E2F0D9"
RED = "FCE4D6"
WHITE = "FFFFFF"

def db():
    con = sqlite3.connect(str(DB_PATH), timeout=30, check_same_thread=False)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys=ON")
    con.execute("PRAGMA journal_mode=WAL")
    return con

def table_columns(con, table):
    return {r["name"] for r in con.execute(f"PRAGMA table_info({table})").fetchall()}

def init_db():
    with db() as con:
        con.executescript("""
        CREATE TABLE IF NOT EXISTS v4_opname (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tanggal TEXT NOT NULL, proyek TEXT NOT NULL, paket TEXT,
            zona TEXT NOT NULL, area TEXT, typical TEXT,
            sta_awal TEXT NOT NULL, sta_akhir TEXT, pekerjaan TEXT NOT NULL,
            metode_volume TEXT NOT NULL, panjang REAL DEFAULT 0,
            lebar REAL DEFAULT 0, tinggi REAL DEFAULT 0,
            volume_rencana REAL NOT NULL, volume_realisasi REAL NOT NULL,
            satuan TEXT NOT NULL, target_progress REAL DEFAULT 0,
            aktual_progress REAL DEFAULT 0, keterangan TEXT, foto_path TEXT,
            status TEXT DEFAULT 'Draft', dibuat_pada TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS v4_schedule (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            proyek TEXT NOT NULL, zona TEXT NOT NULL, pekerjaan TEXT NOT NULL,
            sta TEXT NOT NULL, tanggal_mulai TEXT NOT NULL,
            tanggal_selesai TEXT NOT NULL, bobot REAL DEFAULT 1,
            volume_rencana REAL DEFAULT 0, satuan TEXT NOT NULL,
            UNIQUE(proyek,zona,pekerjaan,sta)
        );
        CREATE TABLE IF NOT EXISTS v4_approvals (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            opname_id INTEGER NOT NULL, status TEXT NOT NULL,
            catatan TEXT, pemeriksa TEXT, waktu TEXT NOT NULL,
            FOREIGN KEY(opname_id) REFERENCES v4_opname(id)
        );
        """)
        # Tambahan kolom aman untuk database V4 lama.
        cols = table_columns(con, "v4_opname")
        additions = {
            "foto_path": "TEXT", "paket": "TEXT", "area": "TEXT",
            "typical": "TEXT", "sta_akhir": "TEXT", "keterangan": "TEXT",
            "status": "TEXT DEFAULT 'Draft'", "target_progress": "REAL DEFAULT 0",
            "aktual_progress": "REAL DEFAULT 0", "panjang": "REAL DEFAULT 0",
            "lebar": "REAL DEFAULT 0", "tinggi": "REAL DEFAULT 0"
        }
        for name, definition in additions.items():
            if name not in cols:
                con.execute(f"ALTER TABLE v4_opname ADD COLUMN {name} {definition}")
        con.execute("CREATE INDEX IF NOT EXISTS idx_v4_opname_project_date ON v4_opname(proyek,tanggal)")
        con.execute("CREATE INDEX IF NOT EXISTS idx_v4_opname_status ON v4_opname(status)")

init_db()

def read_df(sql, params=()):
    with db() as con:
        return pd.read_sql_query(sql, con, params=params)

def pct(real, plan):
    return 0.0 if not plan or float(plan) <= 0 else max(0.0, float(real) / float(plan) * 100.0)

def fmt_num(v, digits=2):
    try:
        return f"{float(v or 0):,.{digits}f}".replace(",", "_").replace(".", ",").replace("_", ".")
    except Exception:
        return "0,00"

def normalize_photo(uploaded, opname_id):
    if uploaded is None:
        return None
    im = ImageOps.exif_transpose(Image.open(uploaded)).convert("RGB")
    im.thumbnail((1800, 1800))
    draw = ImageDraw.Draw(im)
    stamp = datetime.now().strftime("%d-%m-%Y %H:%M")
    label = f"PROGRES OPNAME | {stamp} | ID {opname_id}"
    try:
        font = ImageFont.truetype("DejaVuSans.ttf", max(16, im.width // 55))
    except Exception:
        font = ImageFont.load_default()
    box = draw.textbbox((0, 0), label, font=font)
    x, y = 18, max(10, im.height - (box[3]-box[1]) - 28)
    draw.rectangle((8, y-8, min(im.width-8, box[2]-box[0]+32), y+(box[3]-box[1])+12), fill=(0,0,0))
    draw.text((x, y), label, font=font, fill=(255,255,255))
    path = PHOTO_DIR / f"opname_{opname_id}_{uuid.uuid4().hex[:8]}.jpg"
    im.save(path, "JPEG", quality=88, optimize=True)
    return str(path)

def get_opname(project="Semua proyek"):
    if project and project != "Semua proyek":
        return read_df("SELECT * FROM v4_opname WHERE proyek=? ORDER BY tanggal DESC,id DESC", (project,))
    return read_df("SELECT * FROM v4_opname ORDER BY tanggal DESC,id DESC")

def get_schedule(project="Semua proyek"):
    if project and project != "Semua proyek":
        return read_df("SELECT * FROM v4_schedule WHERE proyek=? ORDER BY tanggal_mulai", (project,))
    return read_df("SELECT * FROM v4_schedule ORDER BY tanggal_mulai")

def insert_opname(d, photo):
    with db() as con:
        cur = con.execute("""
        INSERT INTO v4_opname
        (tanggal,proyek,paket,zona,area,typical,sta_awal,sta_akhir,pekerjaan,
         metode_volume,panjang,lebar,tinggi,volume_rencana,volume_realisasi,
         satuan,target_progress,aktual_progress,keterangan,foto_path,status,dibuat_pada)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (d["tanggal"],d["proyek"],d["paket"],d["zona"],d["area"],d["typical"],
              d["sta_awal"],d["sta_akhir"],d["pekerjaan"],d["metode_volume"],
              d["panjang"],d["lebar"],d["tinggi"],d["volume_rencana"],
              d["volume_realisasi"],d["satuan"],d["target_progress"],
              d["aktual_progress"],d["keterangan"],None,"Draft",
              datetime.now().isoformat(timespec="seconds")))
        oid = cur.lastrowid
        photo_path = normalize_photo(photo, oid)
        if photo_path:
            con.execute("UPDATE v4_opname SET foto_path=? WHERE id=?", (photo_path,oid))
    return oid

def save_schedule(d):
    with db() as con:
        con.execute("""
        INSERT INTO v4_schedule
        (proyek,zona,pekerjaan,sta,tanggal_mulai,tanggal_selesai,bobot,volume_rencana,satuan)
        VALUES (?,?,?,?,?,?,?,?,?)
        ON CONFLICT(proyek,zona,pekerjaan,sta) DO UPDATE SET
        tanggal_mulai=excluded.tanggal_mulai,tanggal_selesai=excluded.tanggal_selesai,
        bobot=excluded.bobot,volume_rencana=excluded.volume_rencana,satuan=excluded.satuan
        """, tuple(d.values()))

def set_status(oid, status, note, reviewer):
    with db() as con:
        con.execute("UPDATE v4_opname SET status=? WHERE id=?", (status, oid))
        con.execute("INSERT INTO v4_approvals(opname_id,status,catatan,pemeriksa,waktu) VALUES(?,?,?,?,?)",
                    (oid,status,note,reviewer,datetime.now().isoformat(timespec="seconds")))

def expected_progress(row, on_date):
    start = pd.to_datetime(row["tanggal_mulai"]).date()
    end = pd.to_datetime(row["tanggal_selesai"]).date()
    if on_date < start: return 0.0
    if on_date >= end or start == end: return 100.0
    return (on_date-start).days / max(1,(end-start).days) * 100.0

def excel_report(op, sched, approvals, project_name):
    """Workbook rapi: ringkasan, detail, jadwal, riwayat approval + foto tertanam."""
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        summary = pd.DataFrame([
            ["Nama Proyek", project_name],
            ["Tanggal Cetak", datetime.now().strftime("%d-%m-%Y %H:%M")],
            ["Jumlah Data Opname", len(op)],
            ["Jumlah Disetujui", int((op["status"]=="Disetujui").sum()) if not op.empty else 0],
            ["Jumlah Draft/Diajukan", int(op["status"].isin(["Draft","Diajukan"]).sum()) if not op.empty else 0],
            ["Total Volume Realisasi", float(op.loc[op["status"]=="Disetujui","volume_realisasi"].sum()) if not op.empty else 0],
        ], columns=["Uraian","Nilai"])
        summary.to_excel(writer, sheet_name="Ringkasan", index=False, startrow=2)
        detail_cols = ["id","tanggal","proyek","paket","zona","area","typical","sta_awal","sta_akhir",
                       "pekerjaan","metode_volume","panjang","lebar","tinggi","volume_rencana",
                       "volume_realisasi","satuan","target_progress","aktual_progress","status",
                       "keterangan","foto_path"]
        detail = op.reindex(columns=detail_cols).copy()
        detail.rename(columns={
            "id":"ID","tanggal":"Tanggal","proyek":"Proyek","paket":"Paket","zona":"Zona",
            "area":"Area Kerja","typical":"Typical","sta_awal":"STA Awal","sta_akhir":"STA Akhir",
            "pekerjaan":"Jenis Pekerjaan","metode_volume":"Metode Volume","panjang":"Panjang (m)",
            "lebar":"Lebar (m)","tinggi":"Tinggi/Tebal (m)","volume_rencana":"Volume Rencana",
            "volume_realisasi":"Volume Realisasi","satuan":"Satuan","target_progress":"Target (%)",
            "aktual_progress":"Aktual (%)","status":"Status","keterangan":"Keterangan","foto_path":"Lokasi Foto"
        }).to_excel(writer, sheet_name="Detail Opname", index=False, startrow=2)
        sched.to_excel(writer, sheet_name="Jadwal", index=False, startrow=2)
        approvals.to_excel(writer, sheet_name="Riwayat Approval", index=False, startrow=2)

    bio.seek(0)
    wb = load_workbook(bio)
    thin = Side(style="thin", color="D9E2F3")
    for ws in wb.worksheets:
        ws.freeze_panes = "A4"
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.outlinePr.summaryBelow = True
        ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=max(2,ws.max_column))
        ws.cell(1,1, f"LAPORAN PROGRES OPNAME JALAN TOL - {ws.title.upper()}")
        ws.cell(1,1).font = Font(name="Arial", size=14, bold=True, color=WHITE)
        ws.cell(1,1).fill = PatternFill("solid", fgColor=NAVY)
        ws.cell(1,1).alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 28
        header_row = 3
        for cell in ws[header_row]:
            cell.font = Font(name="Arial", bold=True, color=WHITE)
            cell.fill = PatternFill("solid", fgColor=NAVY)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(left=thin,right=thin,top=thin,bottom=thin)
        for row in ws.iter_rows(min_row=header_row+1):
            for cell in row:
                cell.font = Font(name="Arial", size=10)
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                cell.border = Border(left=thin,right=thin,top=thin,bottom=thin)
                if cell.row % 2 == 0:
                    cell.fill = PatternFill("solid", fgColor=LIGHT)
        ws.auto_filter.ref = f"A{header_row}:{get_column_letter(ws.max_column)}{ws.max_row}"
        for col in range(1, ws.max_column+1):
            letter = get_column_letter(col)
            max_len = max([len(str(ws.cell(r,col).value or "")) for r in range(header_row, min(ws.max_row,80)+1)] + [10])
            ws.column_dimensions[letter].width = min(max(max_len+2,12),32)
        ws.sheet_format.defaultRowHeight = 22
        ws.print_title_rows = f"1:{header_row}"
        ws.page_margins.left = 0.25; ws.page_margins.right = 0.25
        ws.page_margins.top = 0.5; ws.page_margins.bottom = 0.5

    # Foto dimasukkan sebagai gambar Excel, bukan hanya tautan/path.
    ws = wb["Detail Opname"]
    headers = {c.value:c.column for c in ws[3]}
    photo_col = headers.get("Lokasi Foto")
    if photo_col:
        ws.cell(3, photo_col, "Dokumentasi Foto")
        ws.column_dimensions[get_column_letter(photo_col)].width = 25
        for idx, (_, row) in enumerate(op.iterrows(), start=4):
            path = row.get("foto_path")
            ws.row_dimensions[idx].height = 95
            if path and os.path.isfile(str(path)):
                try:
                    img = XLImage(str(path))
                    img.width, img.height = 125, 85
                    ws.add_image(img, f"{get_column_letter(photo_col)}{idx}")
                except Exception:
                    ws.cell(idx,photo_col,"Foto gagal disisipkan")
            else:
                ws.cell(idx,photo_col,"Tidak ada foto")
    # Format angka/progres.
    for r in range(4, wb["Detail Opname"].max_row+1):
        for col_name in ["Volume Rencana","Volume Realisasi"]:
            c = headers.get(col_name)
            if c: wb["Detail Opname"].cell(r,c).number_format = '#,##0.00'
        for col_name in ["Target (%)","Aktual (%)"]:
            c = headers.get(col_name)
            if c: wb["Detail Opname"].cell(r,c).number_format = '0.00"%"'
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()

# ---------------------- AUTH ----------------------
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "login_name" not in st.session_state:
    st.session_state.login_name = "Administrator"

with st.sidebar:
    st.title("🏗️ Progres Opname")
    st.caption("Jalan Tol | V4.1")
    if not st.session_state.logged_in:
        st.subheader("Login")
        uname = st.text_input("Nama pengguna", value="admin")
        password = st.text_input("Password", type="password")
        if st.button("Masuk", use_container_width=True):
            expected = os.environ.get("APP_PASSWORD", "tol123")
            if password == expected:
                st.session_state.logged_in = True
                st.session_state.login_name = uname.strip() or "Administrator"
                st.rerun()
            else:
                st.error("Password tidak sesuai.")
        st.caption("Akun awal: admin / tol123. Ganti APP_PASSWORD sebelum digunakan bersama.")
        st.stop()
    st.success(f"Login: {st.session_state.login_name}")
    page = st.radio("Menu", ["Dashboard","Input Opname","Jadwal & Target","Rekap Zona",
                             "Daftar Opname","Verifikasi","Laporan & Ekspor","Backup & Data"])
    st.divider()
    if st.button("Keluar", use_container_width=True):
        st.session_state.logged_in = False
        st.rerun()

projects_df = read_df("SELECT proyek FROM v4_opname UNION SELECT proyek FROM v4_schedule")
projects = sorted(projects_df["proyek"].dropna().astype(str).unique().tolist())
with st.sidebar:
    selected_project = st.selectbox("Filter proyek", ["Semua proyek"] + projects)
op = get_opname(selected_project)
sched = get_schedule(selected_project)
approvals = read_df("""SELECT a.id,a.opname_id,a.status,a.catatan,a.pemeriksa,a.waktu,
                       o.proyek,o.zona,o.pekerjaan,o.sta_awal
                       FROM v4_approvals a LEFT JOIN v4_opname o ON o.id=a.opname_id
                       ORDER BY a.waktu DESC""")
if selected_project != "Semua proyek":
    approvals = approvals[approvals["proyek"] == selected_project]

st.title("🏗️ Aplikasi Progres Opname Jalan Tol")
st.caption("Monitoring progres pekerjaan, pencatatan volume, dokumentasi foto, jadwal, approval, dan laporan.")

# ---------------------- DASHBOARD ----------------------
if page == "Dashboard":
    st.header("Dashboard Monitoring")
    approved = op[op.status=="Disetujui"].copy() if not op.empty else op
    zones = op.zona.nunique() if not op.empty else 0
    delayed_rows = []
    for _, s in sched.iterrows():
        a = approved[(approved.proyek==s.proyek)&(approved.zona==s.zona)&
                     (approved.pekerjaan==s.pekerjaan)&(approved.sta_awal==s.sta)]
        actual = pct(a.volume_realisasi.sum() if not a.empty else 0, s.volume_rencana)
        plan = expected_progress(s,date.today())
        delayed_rows.append({"Zona":s.zona,"Pekerjaan":s.pekerjaan,"STA":s.sta,
                             "Rencana (%)":plan,"Aktual (%)":actual,
                             "Deviasi (%)":actual-plan,
                             "Status":"Tertinggal" if actual+0.01<plan else "Sesuai/lebih cepat"})
    avg = (sum(x["Aktual (%)"] for x in delayed_rows)/len(delayed_rows)) if delayed_rows else (
        approved.aktual_progress.mean() if not approved.empty else 0)
    c1,c2,c3,c4=st.columns(4)
    c1.metric("Total opname",len(op)); c2.metric("Zona terdata",zones)
    c3.metric("Progres aktual rata-rata",f"{avg:.1f}%")
    c4.metric("Item tertinggal",sum(x["Status"]=="Tertinggal" for x in delayed_rows))
    st.divider()
    left,right=st.columns(2)
    with left:
        st.subheader("Status opname")
        if not op.empty:
            counts=op.status.value_counts().rename_axis("Status").reset_index(name="Jumlah")
            st.bar_chart(counts.set_index("Status"))
        else: st.info("Belum ada data opname.")
    with right:
        st.subheader("Rencana vs aktual")
        if delayed_rows:
            chart=pd.DataFrame(delayed_rows)
            st.bar_chart(chart.set_index("Pekerjaan")[["Rencana (%)","Aktual (%)"]])
        else: st.info("Isi jadwal dan opname yang disetujui untuk menampilkan grafik.")
    st.subheader("Daftar pekerjaan yang tertinggal")
    late=pd.DataFrame(delayed_rows)
    if not late.empty:
        late=late[late.Status=="Tertinggal"].sort_values("Deviasi (%)")
        st.dataframe(late,use_container_width=True,hide_index=True) if not late.empty else st.success("Tidak ada item tertinggal.")
    else: st.caption("Deteksi keterlambatan memerlukan baseline jadwal.")

# ---------------------- INPUT ----------------------
elif page == "Input Opname":
    st.header("Input Opname Lapangan")
    with st.form("opname_form",clear_on_submit=True):
        a,b,c=st.columns(3)
        tanggal=a.date_input("Tanggal opname",date.today())
        proyek=b.text_input("Nama proyek",selected_project if selected_project!="Semua proyek" else "")
        paket=c.text_input("Paket pekerjaan")
        a,b,c=st.columns(3)
        zona=a.text_input("Zona kerja",placeholder="Contoh: Zona 1 / Main Road")
        area=b.text_input("Area kerja",placeholder="Contoh: Struktur / Drainase")
        typical=c.text_input("Typical",placeholder="Contoh: Typical PC-01")
        a,b=st.columns(2)
        sta_awal=a.text_input("STA awal",placeholder="12+500")
        sta_akhir=b.text_input("STA akhir",placeholder="12+550")
        pekerjaan=st.text_input("Jenis pekerjaan",placeholder="Pengecoran pile cap")
        metode=st.radio("Metode volume",["Input volume langsung","Hitung dari dimensi"],horizontal=True)
        panjang=lebar=tinggi=0.0
        if metode=="Hitung dari dimensi":
            a,b,c=st.columns(3)
            panjang=a.number_input("Panjang (m)",min_value=0.0,step=0.1)
            lebar=b.number_input("Lebar (m)",min_value=0.0,step=0.1)
            tinggi=c.number_input("Tinggi/tebal (m)",min_value=0.0,step=0.1)
            volume=panjang*lebar*tinggi
            satuan="m³"
            st.info(f"Volume hasil perhitungan: {volume:,.3f} m³")
        else:
            a,b=st.columns(2)
            volume=a.number_input("Volume realisasi",min_value=0.0,step=0.1)
            satuan=b.selectbox("Satuan",["m³","m²","m","kg","unit","titik","ls"])
        a,b=st.columns(2)
        rencana=a.number_input("Volume rencana/kontrak item",min_value=0.01,step=0.1)
        target=b.number_input("Target progres (%)",min_value=0.0,max_value=100.0,step=1.0)
        keterangan=st.text_area("Keterangan lapangan")
        foto=st.file_uploader("Foto dokumentasi",type=["jpg","jpeg","png"])
        submit=st.form_submit_button("Simpan Opname",use_container_width=True)
    if submit:
        if not all([proyek.strip(),zona.strip(),sta_awal.strip(),pekerjaan.strip()]):
            st.error("Proyek, zona, STA awal, dan pekerjaan wajib diisi.")
        elif metode=="Hitung dari dimensi" and min(panjang,lebar,tinggi)<=0:
            st.error("Panjang, lebar, dan tinggi/tebal harus lebih dari nol.")
        elif volume>rencana:
            st.error("Volume realisasi melebihi volume rencana. Periksa kembali data kontrak.")
        else:
            d={"tanggal":tanggal.isoformat(),"proyek":proyek.strip(),"paket":paket.strip(),
               "zona":zona.strip(),"area":area.strip(),"typical":typical.strip(),
               "sta_awal":sta_awal.strip(),"sta_akhir":sta_akhir.strip(),"pekerjaan":pekerjaan.strip(),
               "metode_volume":metode,"panjang":panjang,"lebar":lebar,"tinggi":tinggi,
               "volume_rencana":rencana,"volume_realisasi":volume,"satuan":satuan,
               "target_progress":target,"aktual_progress":pct(volume,rencana),"keterangan":keterangan.strip()}
            oid=insert_opname(d,foto)
            st.success(f"Data opname ID #{oid} tersimpan dengan status Draft.")
            st.info(f"Progres volume item: {pct(volume,rencana):.2f}%. Data masuk progres resmi setelah disetujui.")
    st.caption("Foto disimpan di folder data/foto dan diberi watermark tanggal serta ID opname.")

# ---------------------- SCHEDULE ----------------------
elif page == "Jadwal & Target":
    st.header("Jadwal dan Target Pekerjaan")
    with st.form("schedule_form"):
        a,b=st.columns(2)
        pr=a.text_input("Nama proyek",selected_project if selected_project!="Semua proyek" else "")
        zn=b.text_input("Zona")
        a,b=st.columns(2)
        job=a.text_input("Jenis pekerjaan")
        sta=b.text_input("STA / kode item")
        a,b=st.columns(2)
        start=a.date_input("Tanggal mulai",date.today())
        finish=b.date_input("Tanggal selesai",date.today()+timedelta(days=30))
        a,b,c=st.columns(3)
        weight=a.number_input("Bobot item (%)",min_value=0.0,max_value=100.0,value=1.0)
        vol=b.number_input("Volume rencana",min_value=0.01,step=0.1)
        unit=c.selectbox("Satuan",["m³","m²","m","kg","unit","titik","ls"])
        save=st.form_submit_button("Simpan / Perbarui Jadwal",use_container_width=True)
    if save:
        if not all([pr.strip(),zn.strip(),job.strip(),sta.strip()]) or finish<start:
            st.error("Lengkapi identitas item dan pastikan tanggal selesai tidak sebelum tanggal mulai.")
        else:
            save_schedule({"proyek":pr.strip(),"zona":zn.strip(),"pekerjaan":job.strip(),"sta":sta.strip(),
                           "tanggal_mulai":start.isoformat(),"tanggal_selesai":finish.isoformat(),
                           "bobot":weight,"volume_rencana":vol,"satuan":unit})
            st.success("Baseline jadwal berhasil disimpan atau diperbarui."); st.rerun()
    sched=get_schedule(selected_project)
    if not sched.empty:
        st.subheader("Baseline pekerjaan")
        st.dataframe(sched,use_container_width=True,hide_index=True)
        dates=pd.date_range(min(pd.to_datetime(sched.tanggal_mulai).min(),pd.Timestamp(date.today())),
                            max(pd.to_datetime(sched.tanggal_selesai).max(),pd.Timestamp(date.today())),freq="D")
        approved=op[op.status=="Disetujui"].copy() if not op.empty else pd.DataFrame()
        curve=[]
        total_weight=sched.bobot.sum()
        for dt in dates:
            plan=actual_w=0.0
            for _,s in sched.iterrows():
                w=float(s.bobot)
                plan+=w*expected_progress(s,dt.date())/100
                a=approved[(approved.proyek==s.proyek)&(approved.zona==s.zona)&
                           (approved.pekerjaan==s.pekerjaan)&(approved.sta_awal==s.sta)&
                           (pd.to_datetime(approved.tanggal).dt.date<=dt.date())] if not approved.empty else pd.DataFrame()
                actual=pct(a.volume_realisasi.sum() if not a.empty else 0,s.volume_rencana)
                actual_w+=w*actual/100
            curve.append({"Tanggal":dt,"Rencana kumulatif (%)":plan/total_weight*100 if total_weight else 0,
                          "Aktual kumulatif (%)":actual_w/total_weight*100 if total_weight else 0})
        st.subheader("Kurva-S Rencana dan Aktual")
        st.line_chart(pd.DataFrame(curve).set_index("Tanggal"))
        st.caption("Progres aktual hanya menghitung opname berstatus Disetujui.")
    else: st.info("Belum ada baseline jadwal.")

# ---------------------- RECAP ----------------------
elif page == "Rekap Zona":
    st.header("Rekap per Zona, Area, Pekerjaan, dan STA")
    approved=op[op.status=="Disetujui"].copy() if not op.empty else op
    if approved.empty: st.info("Belum ada opname yang disetujui.")
    else:
        summary=approved.groupby(["zona","area","pekerjaan","satuan"],dropna=False).agg(
            volume_rencana=("volume_rencana","max"),volume_realisasi=("volume_realisasi","sum"),
            jumlah_opname=("id","count")).reset_index()
        summary["Progres (%)"]=summary.apply(lambda r:pct(r.volume_realisasi,r.volume_rencana),axis=1)
        st.subheader("Ringkasan zona")
        st.dataframe(summary,use_container_width=True,hide_index=True)
        st.subheader("Rincian STA")
        st.dataframe(approved[["tanggal","zona","area","pekerjaan","sta_awal","sta_akhir",
                               "volume_realisasi","satuan","status"]],use_container_width=True,hide_index=True)

# ---------------------- LIST ----------------------
elif page == "Daftar Opname":
    st.header("Daftar Opname dan Dokumentasi")
    if op.empty: st.info("Belum ada data opname.")
    else:
        f1,f2,f3=st.columns(3)
        status_filter=f1.selectbox("Filter status",["Semua","Draft","Diajukan","Disetujui","Ditolak"])
        zone_options=["Semua"]+sorted(op.zona.dropna().astype(str).unique().tolist())
        zone_filter=f2.selectbox("Filter zona",zone_options)
        search=f3.text_input("Cari pekerjaan / STA")
        view=op.copy()
        if status_filter!="Semua": view=view[view.status==status_filter]
        if zone_filter!="Semua": view=view[view.zona==zone_filter]
        if search:
            view=view[view.apply(lambda r:search.lower() in (str(r.pekerjaan)+" "+str(r.sta_awal)+" "+str(r.id)).lower(),axis=1)]
        st.dataframe(view[["id","tanggal","proyek","zona","area","pekerjaan","sta_awal",
                           "volume_rencana","volume_realisasi","satuan","aktual_progress","status"]],
                     use_container_width=True,hide_index=True)
        if not view.empty:
            oid=st.selectbox("Pilih ID untuk detail",view.id.astype(int).tolist())
            row=view[view.id==oid].iloc[0]
            st.markdown(f"**{row.pekerjaan} | STA {row.sta_awal} – {row.sta_akhir or '-'}**")
            c1,c2,c3=st.columns(3)
            c1.metric("Volume realisasi",f"{fmt_num(row.volume_realisasi)} {row.satuan}")
            c2.metric("Progres",f"{row.aktual_progress:.2f}%")
            c3.metric("Status",row.status)
            if row.foto_path and os.path.isfile(str(row.foto_path)):
                st.image(row.foto_path,caption=f"Dokumentasi opname #{oid}",use_container_width=True)
            else: st.warning("Foto tidak tersedia atau file foto sudah dipindahkan.")
            if row.keterangan: st.info(row.keterangan)
            with st.expander("Edit keterangan / data status"):
                new_note=st.text_area("Keterangan",value=str(row.keterangan or ""),key=f"note_{oid}")
                if st.button("Simpan keterangan"):
                    with db() as con: con.execute("UPDATE v4_opname SET keterangan=? WHERE id=?",(new_note,oid))
                    st.success("Keterangan diperbarui."); st.rerun()

# ---------------------- APPROVAL ----------------------
elif page == "Verifikasi":
    st.header("Verifikasi dan Persetujuan Opname")
    st.warning("V4.1 masih memakai nama pemeriksa yang diinput manual; ini belum merupakan tanda tangan digital atau autentikasi multi-role.")
    pending=op[op.status.isin(["Draft","Diajukan"])] if not op.empty else op
    if pending.empty: st.success("Tidak ada data yang menunggu verifikasi.")
    else:
        st.dataframe(pending[["id","tanggal","proyek","zona","pekerjaan","sta_awal",
                              "volume_realisasi","satuan","status"]],use_container_width=True,hide_index=True)
        oid=st.selectbox("Pilih ID opname",pending.id.astype(int).tolist())
        row=pending[pending.id==oid].iloc[0]
        if row.foto_path and os.path.isfile(str(row.foto_path)): st.image(row.foto_path,width=360)
        reviewer=st.text_input("Nama pemeriksa")
        decision=st.radio("Keputusan",["Diajukan","Disetujui","Ditolak"],horizontal=True)
        note=st.text_area("Catatan pemeriksaan")
        if st.button("Simpan keputusan",use_container_width=True):
            if not reviewer.strip(): st.error("Nama pemeriksa wajib diisi.")
            else:
                set_status(int(oid),decision,note,reviewer.strip())
                st.success(f"Status opname #{oid} menjadi {decision}."); st.rerun()
    st.subheader("Riwayat verifikasi")
    st.dataframe(approvals,use_container_width=True,hide_index=True) if not approvals.empty else st.caption("Belum ada riwayat approval.")

# ---------------------- REPORT ----------------------
elif page == "Laporan & Ekspor":
    st.header("Laporan dan Ekspor")
    if op.empty: st.info("Belum ada data untuk diekspor.")
    else:
        st.caption("File Excel berisi ringkasan, detail opname, jadwal, riwayat approval, dan foto yang disisipkan langsung ke sel Excel.")
        st.dataframe(op[["id","tanggal","proyek","zona","pekerjaan","sta_awal",
                         "volume_realisasi","satuan","aktual_progress","status"]],
                     use_container_width=True,hide_index=True)
        csv=op.to_csv(index=False).encode("utf-8-sig")
        st.download_button("⬇️ Download CSV",csv,"laporan_opname_v41.csv","text/csv")
        xlsx=excel_report(op,sched,approvals,selected_project)
        st.download_button("⬇️ Download Excel + Foto",xlsx,
                           f"laporan_progres_opname_{datetime.now():%Y%m%d_%H%M}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
        st.info("Catatan: foto ditanam sebagai objek gambar Excel. Ukuran file akan bertambah sesuai jumlah dan resolusi foto.")

# ---------------------- BACKUP ----------------------
elif page == "Backup & Data":
    st.header("Backup dan Pengelolaan Data")
    st.write(f"**Lokasi database:** `{DB_PATH}`")
    st.write(f"**Folder foto:** `{PHOTO_DIR}`")
    c1,c2=st.columns(2)
    c1.metric("Ukuran database",f"{DB_PATH.stat().st_size/1024:.1f} KB" if DB_PATH.exists() else "0 KB")
    c2.metric("Jumlah foto",len(list(PHOTO_DIR.glob("*"))) if PHOTO_DIR.exists() else 0)
    if st.button("Buat backup database + foto",use_container_width=True):
        stamp=datetime.now().strftime("%Y%m%d_%H%M%S")
        backup=DATA_DIR/f"backup_opname_{stamp}"
        backup.mkdir(exist_ok=True)
        with db() as con:
            dest=sqlite3.connect(str(backup/"progres_v4.db"))
            con.backup(dest); dest.close()
        if PHOTO_DIR.exists(): shutil.copytree(PHOTO_DIR,backup/"foto",dirs_exist_ok=True)
        zip_path=DATA_DIR/f"backup_opname_{stamp}.zip"
        import zipfile
        with zipfile.ZipFile(zip_path,"w",zipfile.ZIP_DEFLATED) as z:
            for p in backup.rglob("*"):
                if p.is_file(): z.write(p,p.relative_to(backup))
        st.session_state["backup_file"]=str(zip_path)
        st.success("Backup berhasil dibuat.")
    if st.session_state.get("backup_file") and os.path.isfile(st.session_state["backup_file"]):
        p=Path(st.session_state["backup_file"])
        st.download_button("⬇️ Download backup ZIP",p.read_bytes(),p.name,"application/zip")
    st.warning("Jika aplikasi dijalankan di Colab, folder lokal runtime tetap dapat terhapus saat runtime di-reset. Untuk permanen, arahkan PROGRES_DATA_DIR ke folder Google Drive yang sudah di-mount.")

st.divider()
st.caption("V4.1 | Data opname resmi dihitung setelah status Disetujui. Selalu cocokkan volume, bobot, dan metode pembayaran dengan BOQ serta prosedur kontrak proyek.")
