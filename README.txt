# Aplikasi Progres Opname Jalan Tol V4.1

## Fitur
- Dashboard ringkasan progres, status opname, deviasi jadwal, dan pekerjaan tertinggal.
- Input opname dengan volume langsung atau hitung dimensi panjang x lebar x tinggi.
- Dokumentasi foto dengan watermark ID dan waktu.
- Jadwal & target pekerjaan serta Kurva-S rencana vs aktual.
- Rekap zona, area, pekerjaan, dan STA.
- Daftar opname dengan filter, pencarian, dan detail foto.
- Verifikasi Draft/Diajukan/Disetujui/Ditolak dan riwayat pemeriksaan.
- Ekspor CSV dan Excel terformat.
- Excel memiliki sheet Ringkasan, Detail Opname, Jadwal, Riwayat Approval; foto tertanam sebagai gambar di sheet Detail Opname.
- Backup database dan folder foto ke ZIP.

## Menjalankan di komputer
```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
streamlit run app_progres_opname_v4_1.py
```
Buka alamat lokal yang ditampilkan Streamlit.

## Login awal
- Username: admin (label pengguna)
- Password: tol123

Ganti password sebelum dipakai: set environment `APP_PASSWORD` dengan password yang kuat.

## Menjaga data tetap ada
Secara default data berada di folder `data/` di samping file aplikasi:
- `data/progres_v4.db`
- `data/foto/`

Jangan menghapus folder `data`. Untuk Colab, mount Google Drive lalu arahkan `PROGRES_DATA_DIR` ke folder Drive tersebut sebelum menjalankan Streamlit. Folder runtime `/content` biasa tidak permanen.

## Kompatibilitas V4
Aplikasi memakai tabel `v4_opname`, `v4_schedule`, dan `v4_approvals`. Letakkan aplikasi ini dengan database V4 lama yang sama melalui variabel `PROGRES_DB` agar data V4 sebelumnya terbaca. Foto lama tetap perlu berada pada path yang tercatat di database atau dipindahkan dan path-nya diperbarui.

## Catatan
V4.1 merupakan penyempurnaan prototipe. Approval masih memakai nama pemeriksa manual, belum tanda tangan digital/role terpisah. Rumus bobot, volume, dan progres harus divalidasi dengan BOQ, kontrak, serta SOP proyek sebelum digunakan sebagai dasar pembayaran.
