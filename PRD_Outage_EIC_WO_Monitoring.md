# Product Requirements Document (PRD)
## Outage EIC Work Order Monitoring System
**PLTU MSW (2 x 30MW) &bull; Section Electric, Instrument & Control (EIC)**

| Metrik Dokumen | Keterangan |
| :--- | :--- |
| **Versi Dokumen** | **1.5 (Production Release)** |
| **Disusun Oleh** | M. Farhan Tandia (EIC & IT Supervisor) & Tim EIC PLTU MSW |
| **Tanggal Pembaruan** | **30 Agustus 2026** |
| **Status Dokumen** | **Approved & Deployed in Production** |
| **Target Runtime** | Standalone Binary (`server.exe`) & Python 3.8+ Localhost / LAN |

---

## 📜 Daftar Isi (Table of Contents)
1. [Latar Belakang & Pernyataan Masalah](#1-latar-belakang--pernyataan-masalah)
2. [Tujuan Produk & Manfaat Bisnis](#2-tujuan-produk--manfaat-bisnis)
3. [Arsitektur Sistem & Karakteristik Deployment](#3-arsitektur-sistem--karakteristik-deployment)
4. [Daftar File yang Diperlukan untuk Menjalankan Aplikasi](#4-daftar-file-yang-diperlukan-untuk-menjalankan-aplikasi-required-files)
5. [Struktur Data & Spesifikasi Master Excel](#5-struktur-data--spesifikasi-master-excel)
6. [Kebutuhan Fungsional (Functional Requirements)](#6-kebutuhan-fungsional-functional-requirements)
7. [Kebutuhan Non-Fungsional (Non-Functional Requirements)](#7-kebutuhan-non-fungsional-non-functional-requirements)
8. [Riwayat Perubahan & Evolusi Pengembangan yang Sangat Lengkap (Complete Changelog)](#8-riwayat-perubahan--evolusi-pengembangan-lengkap-changelog)
9. [Petunjuk Pengoperasian & Alur Kerja Lapangan](#9-petunjuk-pengoperasian--alur-kerja-lapangan)

---

## 1. Latar Belakang & Pernyataan Masalah

Setiap pelaksanaan pemeliharaan periodik pembangkit (*Major / Minor Outage*), tim Section Electric, Instrument & Control (EIC) PLTU MSW mengelola ratusan item pekerjaan lapangan yang sebelumnya tersebar di berbagai file spreadsheet terpisah:

1. **`Progress_Outage_Unit_1_EIC.xlsx`** — Tracker Work Order utama section EIC berisi daftar WO beserta checklist sub-task teknis (misal: *inspection frame motor*, *regreasing*, *megger test*, *solo run*), status progress, dan PIC pelaksana.
2. **`AMP-MSW-Progress_Actuator_Unit_1_2026.xlsx`** — Tracker verifikasi motorized actuator valve per area (Boiler, Turbine, Aux) dengan tahapan *General Inspection* dan *Function Test*.
3. **`JAPA-MSW-Progress_Transmitter_&_Switch_Unit_2_2026.xls`** — Tracker verifikasi dan kalibrasi field transmitter (PTX, TTX) dan pressure switch (PSW) dengan data *As Found* vs *As Left*.
4. **`PIC_Outage_Unit_1_2026.xlsx`** — Master pembagian scope pekerjaan per kategori (Vendor Scope vs MSW Scope) dan plotting PIC.

### Permasalahan yang Dihadapi:
- **Fragmentasi Data:** Data tersebar di 4+ file spreadsheet dengan format berbeda, memicu risiko duplikasi, konflik versi data, dan inkonsistensi.
- **Ketiadaan Visibilitas Terpadu:** Manajemen dan supervisor tidak memiliki satu dashboard real-time untuk memantau kemajuan total outage tanpa merekap manual dari banyak lembar kerja.
- **Keterlambatan Pelaporan:** Penyusunan laporan progres harian, pembuatan kurva capaian (Kurva-S), dan ringkasan ke grup koordinasi membutuhkan waktu rekap manual yang menyita waktu operasional lapangan.

---

## 2. Tujuan Produk & Manfaat Bisnis

Platform **Outage EIC Work Order Monitoring System** dibangun untuk mencapai sasaran strategis:
- **Unifikasi Database:** Menyatukan seluruh data WO, Actuator, Instrument, dan PIC ke dalam database terstandarisasi (`Template_Outage_EIC_Monitoring_unit 1.xlsx` dan `unit 2.xlsx`).
- **Sinkronisasi Dua Arah Real-Time (*Two-Way Live Sync*):** Setiap centang checklist, pengisian temuan, dan perubahan status di aplikasi web langsung disinkronkan secara aman (*thread-safe*) ke spreadsheet Excel.
- **Otomatisasi Pelaporan:** Menghasilkan 4 opsi laporan resmi berstandar operasional, pencetakan PDF, ekspor spreadsheet utuh, Kurva-S visual, dan format ringkasan WhatsApp dengan 1-klik.
- **Digitalisasi Bukti Lapangan (*Paperless*):** Pengunggahan multi-foto temuan lapangan dan dokumentasi rekomendasi tindak lanjut secara terpusat.

---

## 3. Arsitektur Sistem & Karakteristik Deployment

Sistem dirancang dengan arsitektur **Zero-Dependency Portable Server** yang tangguh, ringan, dan siap dijalankan di lingkungan industri pembangkit tanpa memerlukan instalasi database server yang rumit:

- **Backend:** Pure Python HTTP Server (`server.py`) berbasis `http.server` dan `openpyxl` untuk manipulasi data Excel.
- **Frontend:** Single-Page Application (SPA) responsif menggunakan Vanilla JavaScript, Modern CSS (Glassmorphism, CSS Grid, CSS Variables), serta SVG Chart Generator untuk visualisasi Kurva-S.
- **Thread-Safety:** Mekanisme penguncian file (`threading.Lock`) untuk mencegah *race condition* saat beberapa pengguna membaca/menulis ke file Excel yang sama.
- **Mode Akses:**
  - Standalone Application (`server.exe`) tanpa instalasi Python.
  - Akses lokal via `http://localhost:8000`.
  - Akses multi-perangkat via Jaringan LAN / WiFi Pembangkit (`http://<IP_KOMPUTER_SERVER>:8000`).

---

## 4. Daftar File yang Diperlukan untuk Menjalankan Aplikasi (Required Files)

Untuk mendistribusikan dan menjalankan aplikasi di komputer operasional mana pun, file dan folder berikut **wajib tersedia dalam satu direktori kerja**:

```plaintext
d:\msw\msw_eic_om\
├── server.exe                                    # [UTAMA] Binary aplikasi executable Windows (Standalone)
├── Template_Outage_EIC_Monitoring_unit 1.xlsx   # [UTAMA] Master database spreadsheet Unit 1
├── Template_Outage_EIC_Monitoring_unit 2.xlsx   # [UTAMA] Master database spreadsheet Unit 2
├── Finding/                                     # [UTAMA] Direktori penyimpanan foto & deskripsi temuan
│   ├── UNIT 1/
│   └── UNIT 2/
├── start_app.bat                                # [OPSIONAL] Skrip peluncur otomatis browser
├── server.py                                    # [DEV] Source code Python backend & frontend
├── server.spec                                  # [DEV] Konfigurasi kompilasi PyInstaller
├── README.md                                    # [DOKUMEN] Panduan pengguna & dokumentasi GitHub
└── PRD_Outage_EIC_WO_Monitoring.md              # [DOKUMEN] Dokumen spesifikasi kebutuhan produk (PRD)
```

### Rincian Kebutuhan Minimum:
| Kategori | Spesifikasi Minimum |
| :--- | :--- |
| **Sistem Operasi** | Windows 7 / 8 / 10 / 11 (64-bit) |
| **Penyimpanan** | Ruang disk minimal 150 MB (termasuk penyimpanan foto temuan) |
| **Browser** | Google Chrome, Microsoft Edge, Firefox, atau browser modern lainnya |
| **Aplikasi Excel** | Microsoft Excel 2010+ / LibreOffice / WPS Office (Hanya jika ingin membuka manual) |
| **Port Jaringan** | Port `8000` (dapat diubah jika diperlukan) |

---

## 5. Struktur Data & Spesifikasi Master Excel

Master database menggunakan workbook `.xlsx` yang terdiri dari 8 lembar kerja (*sheets*):

1. **`WorkOrder`**: Menyimpan data utama WO: `No, No_WO, Unit, Job_Description, Area, Tanggal_Schedule, Tanggal_Actual_Start, Tanggal_Finish, Status, PIC, N_Task, Persen_Progress, Scope, Remarks, Temuan, Tindak_Lanjut, Jumlah_Foto`.
2. **`WorkOrder_Checklist`**: Menyimpan sub-task teknis per WO: `No_WO, Sub_Task_Description, Tanggal, PIC_Task, Selesai_TRUE_FALSE, Temuan, Tindak_Lanjut, Jumlah_Foto`.
3. **`ActuatorValve`**: Menyimpan data seluruh motorized actuator valve: `Equipment_ID, Area, Equipment_Description, KKS, Unit, PIC, Status, Persen_Progress, Finish_Date, General_Inspection_TRUE_FALSE, Function_Test_TRUE_FALSE, Remarks, Temuan, Tindak_Lanjut, Jumlah_Foto`.
4. **`Instrument_PressureTX`**: Data transmitter tekanan: `No, Area, Equipment, KKS, Unit, Range, Tanggal/Finish_Date, Status_WDONE_TRUE_FALSE, Remarks, Temuan, Tindak_Lanjut, Jumlah_Foto`.
5. **`Instrument_TemperatureTX`**: Data transmitter suhu: `No, Area, Equipment, KKS, Unit, Range, Tanggal/Finish_Date, Status_WDONE_TRUE_FALSE, Remarks, Temuan, Tindak_Lanjut, Jumlah_Foto`.
6. **`Instrument_PressureSwitch`**: Data switch tekanan lengkap: `No, Area, Description, KKS, Unit, Sub_Area, Set_Point, Contact_Type_NO_NC, AsFound_Set, AsFound_Reset, AsLeft_Set, AsLeft_Reset, Status_OK_NotOK, Status_WDONE_TRUE_FALSE, Dated, Finish_Date, Remarks, Temuan, Tindak_Lanjut, Jumlah_Foto`.
7. **`PIC_Scope_Master`**: Master pembagian kategori scope (A-E) dan pemetaan PIC.
8. **`Dashboard_Summary`**: Ringkasan kalkulasi otomatis via formula `COUNTIFS`.

---

## 6. Kebutuhan Fungsional (Functional Requirements)

### 6.1. Real-Time Dashboard & Manajemen KPI
- **FR-1.1:** Menghitung otomatis persentase Grand Total Outage berdasarkan perolehan sub-task yang telah diselesaikan.
- **FR-1.2:** Menampilkan metrik terpisah untuk *Work Orders*, *Actuator Valves*, *Instruments*, dan *Active Findings*.
- **FR-1.3:** Mendukung perpindahan instan antara **UNIT 1** dan **UNIT 2** melalui tombol toggle di Outage Banner.
- **FR-1.4:** Menyediakan filter cepat (*Quick Filter Chips*): *Semua Item*, *🚨 Ada Temuan / Foto*, *⏳ In Progress*, dan *☑️ Selesai*.
- **FR-1.5:** Paginasi data dinamis dengan pilihan ukuran halaman: **20 / hal** *(Default)*, **40 / hal**, dan **Semua Item**.

### 6.2. Manajemen Work Order & Sub-Task
- **FR-2.1:** Format kartu akordion (*Collapsible Card*) dengan status awal tertutup (*collapsed by default*) saat pertama dimuat.
- **FR-2.2:** Checklist sub-task interaktif dengan tanggal pelaksanaan otomatis dan badge tipe pekerjaan (*Electrical*, *Instrument*, *Mechanical*, *Testing*).
- **FR-2.3:** Fitur **Batch Action Sub-Task**: tombol `✓ Selesai Semua` (100% Finish) dan `↺ Reset` (0% Sched-OK).
- **FR-2.4:** Penambahan sub-task baru secara *Manual*, *Pilih dari Master Actuator*, atau *Pilih dari Master Instrument*.
- **FR-2.5:** Auto-Fill tanggal selesai (`DD/MM/YYYY`) saat seluruh sub-task selesai 100% dan pengosongan otomatis jika status belum tuntas.

### 6.3. Sinkronisasi Dua Arah Lintas Komponen (*Cross-Component Sync*)
- **FR-3.1:** Mencocokkan sub-task WO dengan komponen Actuator / Instrument berdasarkan kesamaan Tag KKS dan padanan nama deskripsi (misal: *DRAUGHT* &harr; *DRAFT*, *ID FAN* &harr; *INDUCED DRAUGHT FAN*).
- **FR-3.2:** Mencentang sub-task actuator/instrument di kartu WO secara instan menyinkronkan status komponen terkait di lembar `ActuatorValve` / `Instrument_*` menjadi `FINISH` (100%) dan mengisi tanggal selesai.
- **FR-3.3:** Mengubah status pada tab Actuator atau Instrument secara otomatis menyinkronkan checklist sub-task pada kartu Work Order yang bersangkutan.

### 6.4. Pusat Laporan Resmi & Ekspor Multi-Format
- **FR-4.1:** Menu terpadu **`📑 Report`** di header yang menyajikan 4 opsi laporan resmi:
  1. *Laporan 1: Progress Harian & Rekap Temuan* (dengan filter rentang tanggal).
  2. *Laporan 2: Work Order & Sub-Task Lengkap*.
  3. *Laporan 3: Actuator Valves Matrix*.
  4. *Laporan 4: Instruments Verification (PTX, TTX, PSW)*.
- **FR-4.2:** Ekspor file Excel utuh (`.xlsx`) termutakhir secara langsung via tombol **"📥 Unduh Excel"**.
- **FR-4.3:** Format cetak siap pakai (*Print-Ready Layout*) yang teroptimasi untuk ekspor PDF.

### 6.5. Kurva-S & Analisis Tren Progres
- **FR-5.1:** Visualisasi grafik Kurva-S interaktif berbasis vektor SVG yang memetakan kumulatif capaian aktual harian terhadap target rencana outage.
- **FR-5.2:** Tabel rincian harian (*daily breakdown*) pertambahan task selesai per tanggal untuk setiap kategori pekerjaan.

### 6.6. Generator Laporan WhatsApp Ringkas
- **FR-6.1:** Otomatisasi penyusunan pesan teks berformat WhatsApp yang memuat rekapitulasi persentase progres, daftar pekerjaan selesai hari ini, dan daftar temuan terbuka.
- **FR-6.2:** Tombol 1-klik **"📋 Salin ke Clipboard"** untuk kenyamanan pelaporan cepat ke grup koordinasi.

### 6.7. Dokumentasi Temuan (*Findings*) & Galeri Foto
- **FR-7.1:** Modal pencatatan anomali lapangan (`Temuan`) dan rekomendasi perbaikan (`Tindak Lanjut`).
- **FR-7.2:** Drag-and-drop & file picker multi-foto lapangan yang otomatis tersimpan di folder `Finding/UNIT X/<Nama Item>/`.
- **FR-7.3:** Lightbox pratinjau foto resolusi penuh dan sinkronisasi jumlah foto ke database Excel.

---

## 7. Kebutuhan Non-Fungsional (Non-Functional Requirements)

- **Performance & Latency:** Waktu respons baca/tulis ke file Excel < 1 detik.
- **UI/UX Excellence:** Desain modern responsif dengan dukungan tema ganda (*Dark Mode* & *Light Mode*) menggunakan palet warna industrial HSL dan tipografi *Inter* & *JetBrains Mono*.
- **Data Integrity:** Proteksi penulisan dengan file lock untuk mencegah korupsi file spreadsheet.
- **Ease of Deployment:** Cukup klik dua kali `server.exe` tanpa perlu dependensi Python runtime di komputer klien.

---

## 8. Riwayat Perubahan & Evolusi Pengembangan Lengkap (Changelog)

Berikut adalah catatan riwayat evolusi sistem yang sangat lengkap dari tahap inisiasi awal hingga rilis produksi:

### 🔹 Versi 1.5 (30 Agustus 2026) — *Current Production Release*
- **Paginasi Dinamis:** Mengubah opsi jumlah item per halaman menjadi **20 / hal** *(Default)*, **40 / hal**, dan **Semua**.
- **Accordion State Initialization:** Menghapus logika auto-expand pada kartu pertama sehingga saat pertama dimuat atau di-refresh, seluruh kartu tertutup rapi (*collapsed by default*).
- **Interactive Header Logo:** Menambahkan aksi klik pada badge logo `⚡ PLTU MSW EIC` untuk me-refresh halaman seketika.
- **Dokumentasi Komprehensif:** Pembaruan dokumen `README.md` (GitHub-ready dengan screenshot) dan `PRD_Outage_EIC_WO_Monitoring.md`.

### 🔹 Versi 1.4 (30 Agustus 2026)
- **Two-Way Cross-Component Sync:** Mengimplementasikan mesin pencocokan KKS pintar (normalisasi awalan unit `10` vs `20` dan sinonim istilah `DRAUGHT`/`DRAFT`/`ID FAN`) untuk menghubungkan checklist WO dengan lembar Actuator dan Instrument secara real-time.
- **Instant Background Toggle:** Menghubungkan centang sub-task frontend langsung ke endpoint `/api/quick_toggle_subtask` sehingga perubahan tersimpan ke Excel dan menyinkronkan komponen terkait secara instan tanpa reload.

### 🔹 Versi 1.3 (30 Agustus 2026)
- **Unit Switcher Relocation:** Memindahkan tombol pemilihan **UNIT 1** dan **UNIT 2** dari header atas ke dalam *Outage Banner* untuk menggantikan badge statis `🔥 Outage Active`.
- **Header Simplification:** Menyederhanakan sisi kanan header utama agar hanya memuat ikon tema `🌙/☀️` dan tombol `📑 Report`.

### 🔹 Versi 1.2 (30 Agustus 2026)
- **Centered Header Layout:** Menerapkan tata letak CSS Grid 3-kolom (`1fr auto 1fr`) sehingga judul *Outage Work Order Monitoring System* dan subjudul berada tepat di tengah layar.
- **Label Button Standardization:** Mengubah label tombol laporan menjadi `📑 Report`.
- **Icon-Only Theme Toggle:** Menghilangkan teks "Dark Mode / Light Mode" sehingga tombol tema menjadi ikon bulat bersih dengan efek hover.

### 🔹 Versi 1.1 (29 Agustus 2026)
- **Startup Auto-Load Fix:** Menghapus referensi callback `updateSummaryUI` yang sebelumnya menyebabkan kendala data Unit 1 tidak otomatis termuat saat pertama kali dibuka.
- **WhatsApp Modal Payload Fix:** Memperbaiki pemetaan kunci objek `grand_pct`, `grand_done`, dan `grand_total` pada fungsi `generateWaText()`.

### 🔹 Versi 1.0 (29 Agustus 2026)
- **Pusat Laporan & Tools Terpadu:** Memindahkan seluruh tombol aksi sekunder (**📈 Kurva-S & Tren**, **📱 Format WA**, dan **📥 Unduh Excel**) ke bagian footer modal Laporan sehingga header aplikasi tetap rapi.
- **Sticky Summary Bar Integration:** Menyederhanakan tombol aksi pada sticky summary bar dengan tombol pintas `📑 Menu Laporan & Tools`.

### 🔹 Versi 0.9 (29 Agustus 2026)
- **Implementasi 5 Fitur Unggulan:**
  1. *Sticky Summary Bar:* Bar ringkasan progres melayang saat scroll > 180px.
  2. *Floating Back-to-Top Button:* Tombol kembali ke puncak halaman dengan scroll halus.
  3. *Batch Action Sub-Task:* Tombol `✓ Selesai Semua` dan `↺ Reset` di setiap kartu WO.
  4. *Live Excel Export:* Endpoint `/api/export_excel?unit=X` untuk mengunduh spreadsheet utuh terkini.
  5. *Kurva-S & Generator WhatsApp:* Visualisasi grafik SVG progres harian dan generator ringkasan pesan WhatsApp.

### 🔹 Versi 0.8 (29 Agustus 2026)
- **Comprehensive Daily Report (Opsi 1):** Memperluas cakupan Laporan Harian agar tidak hanya mencatat pembaruan Work Order, tetapi juga mencantumkan aktivitas inspeksi Actuator Valve dan kalibrasi Instrument yang diselesaikan pada tanggal terkait.

### 🔹 Versi 0.7 (29 Agustus 2026)
- **Automatic Finish Date Synchronization:** Menambahkan logika backend dan frontend untuk otomatis mengisi tanggal hari ini (`DD/MM/YYYY`) sebagai tanggal selesai WO/Actuator/Instrument saat checklist mencapai 100%, dan otomatis mengosongkannya kembali saat ada checklist yang belum tuntas.

### 🔹 Versi 0.6 (29 Agustus 2026)
- **Actuator Indexing & String Bug Fix:** Memperbaiki pergeseran indeks kolom openpyxl pada handler `save_actuator_update` dan menangani error `(item.status).replace is not a function` dengan safe string wrapper.

### 🔹 Versi 0.5 (29 Agustus 2026)
- **Dual Theme Support:** Menambahkan switcher Dark Mode / Light Mode berbasis CSS custom properties (`data-theme="light"` / `data-theme="dark"`).
- **Light Mode UI Fixes:** Memperbaiki kontras teks pada pill Master PIC, date badge, dan kartu laporan saat berada dalam tema terang.

### 🔹 Versi 0.4 (29 Agustus 2026)
- **Preserve Accordion State:** Mengimplementasikan `openCardIds` state tracking agar kartu yang sedang dibuka tidak tertutup otomatis saat pengguna menekan tombol "Simpan Perubahan".

### 🔹 Versi 0.3 (29 Agustus 2026)
- **PIC Master Persistence:** Menyeragamkan dropdown PIC di seluruh kartu data dan memperbaiki persistensi pemilihan PIC pada master scope.

### 🔹 Versi 0.2 (29 Agustus 2026)
- **Initial Backend & SPA Architecture:** Membangun core HTTP Server (`server.py`), REST API (`/api/data`, `/api/update_wo`, dll), openpyxl reader/writer, serta antarmuka monitoring Unit 1 & Unit 2.

### 🔹 Versi 0.1 (29 Agustus 2026)
- **Initial Draft PRD:** Perumusan spesifikasi kebutuhan unifikasi 4 spreadsheet terpisah ke dalam satu template terpadu.

---

## 9. Petunjuk Pengoperasian & Alur Kerja Lapangan

1. **Memulai Aplikasi:**
   - Cukup klik ganda file `server.exe` di folder kerja.
   - Browser akan otomatis terbuka dan menampilkan dashboard monitoring Unit 1.
2. **Memperbarui Progres Work Order:**
   - Klik kartu Work Order yang ingin diperbarui.
   - Centang sub-task yang telah selesai dikerjakan teknisi di lapangan (tanggal selesai akan terisi otomatis).
   - Klik tombol **`💾 Simpan Perubahan`** jika melakukan perubahan data umum (PIC atau catatan).
3. **Mencetak Laporan / Mengirim Ringkasan WhatsApp:**
   - Klik tombol **`📑 Report`** di header atas.
   - Pilih salah satu dari 4 jenis laporan, sesuaikan tanggal filter jika diperlukan, lalu klik **`🖨️ Cetak / Simpan PDF`**.
   - Untuk membagikan progres ke grup WhatsApp, klik tombol **`📱 Format WA`** di bagian bawah modal lalu klik **`📋 Salin ke Clipboard`**.
4. **Melihat Kurva Capaian:**
   - Klik tombol **`📈 Kurva-S & Tren`** di bagian bawah modal Report untuk melihat grafik kumulatif realisasi vs target harian.

---

<div align="center">

**PLTU MSW &bull; Section Electric, Instrument & Control**  
*Excellence in Operation & Maintenance Execution*

</div>
