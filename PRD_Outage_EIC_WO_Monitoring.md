# Product Requirements Document (PRD)
## Outage EIC Work Order Monitoring System
**PLTU MSW – Section Electric, Instrument & Control (EIC)**

| | |
|---|---|
| Versi | 0.1 (Draft) |
| Disusun oleh | M. Farhan Tandia (EIC & IT Supervisor) |
| Tanggal | 29 Agustus 2026 |
| Status | Draft untuk review |

---

## 1. Latar Belakang

Setiap periode outage (mayor/minor) unit PLTU, tim EIC mengelola progress pekerjaan melalui beberapa file Excel terpisah yang dibuat oleh PIC berbeda:

1. **Progress_Outage_Unit_1_EIC.xlsx** — tracker WO (Work Order) utama section EIC, berisi daftar WO dengan sub-task checklist (mis. inspection frame motor, regreasing, solo run, dll), status (SCHED-OK / IN PROGRESS / FINISH), PIC, dan % estimasi progress per WO.
2. **AMP-MSW-Progress_Actuator_Unit_1_2026.xlsx** — tracker khusus actuator valve, per area (Boiler, ID Fan, STG, PA Fan, dst), dengan dua sub-task standar (General Inspection & Function Test) dan status/% progress.
3. **JAPA-MSW-Progress_Transmitter_&_Switch_Unit_2_2026.xls** — tracker verifikasi & kalibrasi field instrument, terdiri dari 3 sheet: Pressure Transmitter, Temperature Transmitter, dan Pressure Switch (yang punya sub-data AS FOUND/AS LEFT untuk titik set HIGH/LOW).
4. **PIC_Outage_Unit_1_2026.xlsx** — master pembagian scope pekerjaan per kategori (A. Electric Utama & Bantu, B. Instrument, C. Rotating, D. Actuator Valve, E. Turbine Major Outage), dipisah antara **Vendor Scope** dan **MSW Scope**, masing-masing dengan PIC penanggung jawab.

Masalah yang muncul:
- Data tersebar di 4+ file Excel berbeda struktur, diisi manual oleh beberapa PIC → **rawan versi ganda, sulit dikonsolidasi real-time**.
- Manajemen/atasan tidak punya **satu dashboard** untuk melihat progress keseluruhan outage EIC (WO, valve, instrument, dan siapa PIC-nya) tanpa membuka & merekap manual dari banyak file.
- Perhitungan agregat (jumlah finish/in-progress, % progress) saat ini dihitung manual/semi-formula di tiap file, rawan salah rekap saat digabung.
- Tidak ada riwayat perubahan (siapa update apa, kapan) — sulit audit saat outage selesai.

## 2. Tujuan Produk

Membangun **web-based Outage EIC Work Order Monitoring System** yang:
- Menyatukan ke-4 sumber data di atas ke dalam satu dashboard progress real-time.
- Berjalan **lokal** dari folder yang di-share via **OneDrive Desktop**, sehingga bisa diakses beberapa PIC sekaligus (maks. ±2 orang bersamaan) tanpa perlu hosting VPS/cloud.
- File Excel dibaca & ditulis langsung oleh aplikasi web — pengguna tidak perlu membuka Excel manual untuk update progress.
- Menyediakan drill-down per WO / per valve / per instrument / per PIC.
- Menghasilkan laporan progress outage (harian/final) secara otomatis, tanpa rekap manual.

### 2.1 Non-Goals (di luar cakupan versi awal)
- Tidak menggantikan proses approval WO di CMMS/ERP resmi perusahaan — belum dibutuhkan integrasi apa pun ke sistem resmi untuk saat ini.
- Tidak mengelola workflow persetujuan vendor/kontrak (procurement).
- Tidak di-hosting di server/VPS — versi ini murni aplikasi lokal (dijalankan dari komputer/folder share OneDrive).
- Tidak ada mekanisme login/PIN sesi.

## 3. Target Pengguna & Akses

**Versi ini tidak menggunakan login/otentikasi.** Semua orang yang mengakses URL sistem (PIC EIC, koordinator vendor, manajemen) memiliki hak akses setara — dapat melihat dashboard **dan** mengedit data (WO, valve, instrument, temuan, upload foto) layaknya admin. Ini dipilih karena kecepatan penggunaan di lapangan saat outage lebih diprioritaskan dibanding kontrol akses berjenjang.

| Jenis pengguna (informal, bukan role sistem) | Aktivitas tipikal |
|---|---|
| PIC / Teknisi EIC | Update status WO, valve, instrument; isi Temuan; upload foto |
| Koordinator Vendor | Update progress scope vendor miliknya |
| Manajemen (Plant Manager, Ops Manager) | Melihat dashboard progress & statistik |

> **Catatan risiko (perlu disadari):** Karena tanpa login, sistem tidak bisa mencatat *siapa* yang mengubah data secara otomatis (kecuali dicatat lewat kolom PIC pada baris data itu sendiri, yang kini berbentuk dropdown — lihat Bagian 4.6). Karena aplikasi berjalan lokal dari folder share OneDrive (bukan diakses lewat jaringan publik), risiko akses tak sah relatif kecil.

## 4. Model Data & Template Excel

Template Excel final (`Template_Outage_EIC_Monitoring.xlsx`, terlampir bersama PRD ini) berisi 8 sheet berikut. **Setiap sheet level equipment/komponen ditambahkan 3 kolom baru sesuai permintaan:** `Temuan`, `Tindak_Lanjut`, `Jumlah_Foto` (kolom terakhir ini otomatis terisi oleh web saat foto diupload, tidak diisi manual).

### 4.1 Sheet `WorkOrder` — dari `Progress_Outage_Unit_1_EIC.xlsx`
`No, No_WO, Unit (1/2), Job_Description, Area, Tanggal_Schedule, Tanggal_Actual_Start, Tanggal_Finish, Status (SCHED-OK/IN PROGRESS/FINISH), PIC, N_Task, Persen_Progress, Scope, Remarks, Temuan, Tindak_Lanjut, Jumlah_Foto`

### 4.2 Sheet `WorkOrder_Checklist` — sub-task per WO
`No_WO, Sub_Task_Description, Tanggal, PIC_Task, Selesai_TRUE_FALSE, Temuan, Tindak_Lanjut, Jumlah_Foto`

### 4.3 Sheet `ActuatorValve` — dari `AMP-MSW-Progress_Actuator...xlsx`
`Equipment_ID, Area, Equipment_Description, KKS, Unit (1/2), PIC, Status, Persen_Progress, Finish_Date, General_Inspection_TRUE_FALSE, Function_Test_TRUE_FALSE, Remarks, Temuan, Tindak_Lanjut, Jumlah_Foto`

**Penanganan "list sama, progress beda per unit":** setiap equipment actuator/valve dibuat **2 baris** dengan `Equipment_ID` & `Equipment_Description` identik, dibedakan hanya oleh kolom `Unit` (1 dan 2) — masing-masing punya kolom Status/Progress/Temuan/Foto sendiri-sendiri. Dashboard nanti bisa menampilkan perbandingan progress Unit 1 vs Unit 2 untuk equipment yang sama.

### 4.4 Sheet `Instrument_PressureTX`, `Instrument_TemperatureTX` — dari `JAPA-...Transmitter...xls`
`No, Area, Equipment, KKS, Unit (1/2), Range, Tanggal/Finish_Date, Status_WDONE_TRUE_FALSE, Remarks, Temuan, Tindak_Lanjut, Jumlah_Foto`

### 4.5 Sheet `Instrument_PressureSwitch`
`No, Area, Description, KKS, Unit, Sub_Area, Set_Point, Contact_Type_NO_NC, AsFound_Set, AsFound_Reset, AsLeft_Set, AsLeft_Reset, Status_OK_NotOK, Status_WDONE_TRUE_FALSE, Dated, Finish_Date, Remarks, Temuan, Tindak_Lanjut, Jumlah_Foto`

### 4.6 Sheet `PIC_Scope_Master` — dari `PIC_Outage_Unit_1_2026.xlsx`
`Kategori (A–E), Nama_Equipment_Scope, Tipe_Scope_Vendor_MSW, Scope_Kerja_ME_SI_SE, Deskripsi_Aktivitas, PIC, Unit`

**Kolom `PIC` di semua sheet progress (WorkOrder, WorkOrder_Checklist, ActuatorValve, 3 sheet Instrument, dan PIC_Scope_Master) ditampilkan di web sebagai dropdown**, bukan isian teks bebas — daftar nama diambil dari kumpulan nama PIC unik yang sudah pernah muncul di data (atau daftar master nama yang disiapkan terpisah), supaya penulisan nama tetap konsisten antar baris/sheet. Tidak ada auto-link by nama equipment antara sheet `PIC_Scope_Master` dan sheet progress lainnya — keduanya tetap berdiri sendiri di v1, hanya kolom PIC yang diseragamkan lewat dropdown.

### 4.7 Sheet `Dashboard_Summary`
Statistik sederhana dihitung otomatis dengan formula (`COUNTIFS`) langsung di dalam file Excel, per Unit 1 & Unit 2: jumlah SCHED-OK/IN PROGRESS/FINISH per kategori (WO, Actuator Valve, Instrument), serta jumlah baris yang punya Temuan terisi per sheet. Sheet ini bersifat read-only (jangan diedit manual) dan menjadi basis tampilan statistik ringkas di web.

### 4.8 Sheet `README`
Legend warna kolom (kuning = wajib isi manual, hijau muda = contoh baris), penjelasan tiap sheet, dan aturan pengisian Unit 1/2 untuk sheet ActuatorValve.

> **Catatan:** Inkonsistensi penamaan Unit 1/Unit 2 pada file JAPA asli (judul file menyebut Unit 2, isi sheet campur "OUTAGE U-1"/"U-2") sudah dirapikan di template baru ini — setiap baris punya kolom `Unit` eksplisit, tidak lagi bergantung pada judul file atau teks di dalam sheet.

## 5. Alur Data & Arsitektur: Aplikasi Lokal via OneDrive Desktop

Arsitektur final **bukan aplikasi berbasis server/VPS**, melainkan **aplikasi lokal** yang dijalankan dari komputer (mis. komputer Supervisor/PIC utama) dan foldernya di-share lewat **OneDrive Desktop sync**, sehingga PIC lain yang punya akses ke folder OneDrive yang sama bisa membuka & menggunakan aplikasi ini dari komputer masing-masing (maks. ±2 pengguna bersamaan).

**Mekanisme:**
1. Seluruh aset aplikasi — file Excel (`Template_Outage_EIC_Monitoring.xlsx`), folder `Finding/`, dan aplikasi web itu sendiri (backend + frontend) — berada dalam **satu folder** yang disinkronkan OneDrive.
2. Aplikasi web dijalankan secara lokal (localhost) di komputer yang membuka folder tersebut; PIC lain yang perlu akses dapat menjalankan salinan aplikasi yang sama dari folder OneDrive ter-sync di komputernya, membaca **file Excel yang sama** (karena sudah tersinkron OneDrive).
3. **Web membaca & menulis langsung ke file Excel** — saat pengguna update status/progress/Temuan di form, backend langsung menulis ke sel yang sesuai dan menyimpan ulang file (memakai `openpyxl`/library setara).
4. Karena OneDrive sync punya jeda (bukan realtime instan), ada kemungkinan kecil dua PIC menyimpan perubahan hampir bersamaan pada file yang sama — untuk mengantisipasi ini, sistem menerapkan **locking sederhana saat proses tulis** (durasi singkat per simpan) dan menampilkan peringatan bila file terdeteksi sedang di-lock oleh proses lain.
5. **Foto temuan tidak diupload dari HP/browser** — foto sudah tersedia di komputer/folder share (folder lokal), dan pengguna cukup **klik untuk memilih file dari folder tersebut langsung di halaman monitoring web** (file picker lokal), lalu sistem menyalin/mereferensikan file itu ke folder `Finding/<Nama WO atau Komponen>/` (lihat Bagian 6.3).
6. Sheet `Dashboard_Summary` tetap dihitung otomatis via formula Excel dan/atau dihitung ulang di level aplikasi setiap kali dashboard dibuka.

**Implikasi desain (dibanding versi PRD sebelumnya):**
- Tidak perlu VPS, tidak perlu pertimbangan kapasitas storage cloud — storage memakai kapasitas lokal/OneDrive yang sudah ada.
- Tidak perlu login/PIN sesi — akses dibatasi secara alami karena hanya orang yang punya akses ke folder OneDrive yang bisa membuka aplikasinya.
- Locking tulis-ke-Excel cukup sederhana karena hanya melayani maks. ±2 pengguna bersamaan, bukan skala banyak pengguna.

## 6. Functional Requirements

### 6.1 Baca & Tulis Data (Excel sebagai Database File)
- FR-1: Sistem membaca seluruh sheet (WorkOrder, WorkOrder_Checklist, ActuatorValve, 3 sheet Instrument, PIC_Scope_Master) dari satu file Excel di folder lokal/OneDrive, tanpa proses upload berulang.
- FR-2: Setiap perubahan yang dilakukan pengguna di web (update status, %progress, tanggal, isi Temuan, dsb.) **langsung ditulis kembali** ke sel yang bersesuaian di file Excel.
- FR-3: Sistem mendukung Unit 1 dan Unit 2 sekaligus dalam satu file/tampilan, dengan filter Unit di setiap halaman.
- FR-4: Untuk sheet ActuatorValve, sistem menampilkan satu daftar equipment (tidak duplikat di UI) dengan dua kolom progress berdampingan (Unit 1 | Unit 2) meski di file Excel tersimpan sebagai 2 baris terpisah.
- FR-5: Backup otomatis harian dari file Excel (agar ada titik pemulihan jika file rusak/salah edit) — disimpan di sub-folder lokal, mis. `Backup/`.
- FR-6: Semua pengguna (tanpa login) memiliki hak edit penuh — tidak ada pembatasan role pada versi ini.
- FR-7: Kolom `PIC` pada semua sheet progress ditampilkan sebagai **dropdown pilihan nama** di form web (bukan input teks bebas), agar penamaan PIC konsisten.

### 6.2 Dashboard & Statistik Sederhana
- FR-7: Dashboard ringkasan per Unit: total WO/valve/instrument, breakdown status (SCHED-OK/IN PROGRESS/FINISH), % selesai keseluruhan — mengacu pada perhitungan yang sama dengan sheet `Dashboard_Summary`.
- FR-8: Filter berdasarkan Unit, Area, PIC, Status, Kategori scope.
- FR-9: Halaman detail per WO menampilkan checklist sub-task dan status masing-masing.
- FR-10: Halaman detail per valve/instrument menampilkan tag KKS, range, status kalibrasi, dan (khusus pressure switch) hasil AS FOUND vs AS LEFT.
- FR-11: Ringkasan jumlah temuan (finding) — total & per kategori (WO/valve/instrument) — sebagai bagian statistik dashboard.

### 6.3 Temuan (Finding) & Foto dari Folder Lokal
- FR-12: Setiap baris equipment/komponen (WO, sub-task checklist, valve, instrument) memiliki field `Temuan` dan `Tindak_Lanjut` yang dapat diisi bebas lewat form web.
- FR-13: Setiap baris menyediakan tombol **"Pilih Foto"** yang membuka file picker lokal (bukan form upload dari HP) — pengguna langsung memilih file foto yang sudah ada di komputer/folder share, bisa lebih dari satu foto per baris/temuan.
- FR-14: Saat foto dipilih, sistem secara otomatis menyalin file tersebut ke struktur folder berikut (di dalam folder aplikasi yang sama, ikut ter-sync OneDrive):
  ```
  Finding/
    └── <Nama WO atau Nama Komponen>/
          ├── deskripsi.txt   (berisi teks Temuan & Tindak Lanjut yang diinput di web)
          ├── foto_1.jpg
          ├── foto_2.jpg
          └── ...
  ```
  Nama folder mengikuti `No_WO` (untuk sheet WorkOrder/Checklist) atau `Equipment_ID`/`Equipment_Description` (untuk valve/instrument), disanitasi agar aman sebagai nama folder (spasi/simbol khusus diganti underscore).
- FR-15: Setiap kali `deskripsi.txt` atau foto diperbarui melalui web, kolom `Jumlah_Foto` pada baris terkait di Excel otomatis diperbarui mengikuti jumlah file foto di folder tersebut.
- FR-16: Pengguna dapat melihat kembali foto & deskripsi temuan langsung dari halaman detail WO/valve/instrument di web (klik langsung dari halaman monitoring, tanpa perlu membuka folder file secara manual — galeri kecil + teks ditampilkan inline).

### 6.4 Pelaporan
- FR-17: Export laporan progress (PDF/Excel) harian & final outage, termasuk lampiran ringkasan temuan.
- FR-18: Sistem dapat menghasilkan daftar "task belum selesai" (outstanding items) per PIC untuk follow-up.

## 7. Non-Functional Requirements
- **Arsitektur lokal**: aplikasi dijalankan lokal (localhost) dari folder yang di-share via OneDrive Desktop — bukan di-hosting di VPS/cloud.
- **Stack teknis yang disarankan**: aplikasi ringan yang mudah dijalankan lokal oleh non-technical user, misalnya backend Python (Flask/FastAPI) yang langsung membaca-tulis file Excel di folder yang sama, dengan frontend web sederhana (React atau HTML biasa) — dijalankan lewat satu file *executable*/script starter agar PIC tinggal klik untuk menjalankan, tanpa perlu instalasi rumit di tiap komputer.
- **Performa**: baca-tulis file Excel per aksi edit harus selesai < 3 detik.
- **Konkurensi tulis**: locking singkat saat proses simpan ke file Excel, cukup untuk skala ±2 pengguna bersamaan (sesuai konfirmasi kebutuhan).
- **Tanpa otentikasi**: tidak ada login maupun PIN sesi; akses secara alami dibatasi karena hanya pengguna yang punya akses folder OneDrive yang bisa menjalankan aplikasinya.
- **Ketahanan data**: backup otomatis (harian atau per sekian kali simpan) dari file Excel + folder `Finding/` ke sub-folder `Backup/` lokal.
- **Penyimpanan foto**: memakai storage lokal/OneDrive yang sudah tersedia — tidak perlu estimasi kapasitas VPS terpisah.
- **Kompatibilitas**: mendukung baca-tulis `.xlsx` (format `.xls` lama dikonversi ke `.xlsx` saat setup awal template); aplikasi ditargetkan berjalan di lingkungan Windows (umum dipakai di lapangan plant).

## 8. Metrik Keberhasilan
- Waktu rekap progress outage EIC dari manual (saat ini, per file) turun signifikan menjadi cukup 1x upload per update.
- Manajemen bisa melihat status progress outage tanpa meminta rekap manual ke PIC.
- Zero selisih data antara laporan final dan Excel sumber (validasi akurasi parsing).

## 9. Fase Pengembangan (Roadmap Usulan)

| Fase | Cakupan |
|---|---|
| **Fase 1 – MVP** | Aplikasi lokal baca-tulis langsung ke file Excel (WorkOrder, ActuatorValve, 3 sheet Instrument, PIC_Scope_Master), dashboard ringkasan per Unit, halaman detail per WO/valve/instrument, dropdown PIC, form isi Temuan + pilih foto dari folder lokal dengan struktur folder `Finding/` |
| **Fase 2** | Halaman beban kerja per PIC (lintas sumber data), export laporan PDF/Excel (termasuk lampiran temuan & foto), galeri foto temuan terpusat |
| **Fase 3** | Notifikasi (mis. WO mendekati deadline outage, task belum update > X hari) |
| **Fase 4 (opsional)** | Jika suatu saat dibutuhkan hosting terpusat (bukan lokal) atau integrasi ke CMMS resmi perusahaan — dievaluasi ulang saat kebutuhan itu muncul |

## 10. Asumsi & Pertanyaan Terbuka (Terjawab)
1. **Format 4 template Excel** — dikonfirmasi tetap sama untuk saat ini.
2. **Jumlah pengguna bersamaan** — maksimal 2 orang (PIC) bersamaan; locking sederhana cukup.
3. **Foto temuan** — tidak diupload dari HP, melainkan dipilih langsung dari folder lokal/share lewat file picker di halaman monitoring; storage memakai folder lokal/OneDrive yang sudah ada.
4. **Link master Scope & PIC ke data progress** — tidak perlu auto-link; cukup kolom `PIC` di semua sheet progress berbentuk dropdown pilihan nama.
5. **Keamanan akses** — tidak perlu jaringan internal/VPN maupun PIN sesi; aplikasi berjalan lokal dari folder OneDrive yang aksesnya sudah dibatasi lewat sharing OneDrive itu sendiri.
6. **Integrasi CMMS resmi** — belum dibutuhkan untuk saat ini.

### Pertanyaan baru yang muncul dari keputusan ini
- Siapa yang akan berperan sebagai "komputer utama" penyimpan folder aplikasi (yang lain menyinkron via OneDrive), untuk memastikan tidak ada kebingungan folder master yang mana.
- Apakah semua PIC yang akan pakai aplikasi ini sudah familiar menjalankan aplikasi lokal (mis. double-click file starter/`.bat`), atau perlu panduan instalasi/on-boarding singkat.
- Perlu dipastikan setup Python/runtime yang dipakai bisa berjalan tanpa instalasi tambahan di komputer PIC lain (misal dengan aplikasi dikemas jadi single executable).

## 11. Lampiran
- `Template_Outage_EIC_Monitoring.xlsx` — template Excel final (8 sheet: Dashboard_Summary, README, WorkOrder, WorkOrder_Checklist, ActuatorValve, Instrument_PressureTX, Instrument_TemperatureTX, Instrument_PressureSwitch, PIC_Scope_Master), lengkap dengan formula statistik otomatis dan contoh pengisian data.
