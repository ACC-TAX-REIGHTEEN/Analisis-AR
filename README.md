# 📊 Automated AR (Accounts Receivable) Analytics System

![Python](https://img.shields.io/badge/Python-3.8%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-Data_Pipeline-150458?style=for-the-badge&logo=pandas&logoColor=white)
![XlsxWriter](https://img.shields.io/badge/XlsxWriter-Excel_Reporting-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Status](https://img.shields.io/badge/Status-Production_Ready-success?style=for-the-badge)

Sistem otomatisasi *pipeline* ETL (*Extract, Transform, Load*) dan pelaporan analitik piutang pelanggan (*Aging Schedule*) berbasis Python. Dirancang khusus untuk mengolah ekspor data piutang mentah dari sistem akuntansi/ERP (seperti Accurate), menstandardisasi format data secara cerdas, memisahkan data per depo/area operasional menggunakan konfigurasi eksternal (`.conf`), serta menghasilkan laporan eksekutif Excel (`.xlsx`) lengkap dengan *Pivot Summary*, peringkat piutang macet, grafik visual, dan rincian transaksi per wiraniaga (*salesman*).

---

## 📑 Daftar Isi

- [Fitur Utama](#-fitur-utama)
- [Arsitektur & Alur Kerja](#-arsitektur--alur-kerja)
- [Teknologi yang Digunakan](#-teknologi-yang-digunakan)
- [Struktur Direktori](#-struktur-direktori)
- [Prasyarat Sistem](#-prasyarat-sistem)
- [Panduan Instalasi & Konfigurasi](#-panduan-instalasi--konfigurasi)
- [Cara Penggunaan](#-cara-penggunaan)
- [Referensi Konfigurasi (`.conf`)](#-referensi-konfigurasi-conf)
- [Penanganan Masalah Umum (Troubleshooting)](#-penanganan-masalah-umum-troubleshooting)
- [Kontribusi & Lisensi](#-kontribusi--lisensi)

---

## 🛠️ Fitur Utama

### 1. Orkestrator Pipeline Satu Pintu (`Analisis_AR.py`)
* **Validasi Pra-Eksekusi Ketat**: Memeriksa keberadaan berkas sumber `Piutang.xls`, direktori pemrosesan `dapur/`, serta kelengkapan 3 skrip inti dan 4 berkas konfigurasi sebelum menjalankan proses.
* **Manajemen Siklus Hidup Berkas Temporer**: Membersihkan sisa berkas sementara (`*temp.xlsx` dan salinan `Piutang.xls`) sebelum dan sesudah eksekusi di dalam blok `finally` sehingga ruang kerja selalu bersih meskipun terjadi interupsi.
* **Menu CLI Interaktif**: Menyediakan antarmuka pemilihan 4 profil area operasional (**Selatan**, **Utara**, **IRC ZN**, dan **Customized**) dan menamai laporan keluaran secara otomatis sesuai area yang dipilih.

### 2. Pembersihan Data & Deteksi Heuristik (`1_CleanerAcc.py`)
* **Pembacaan Multi-Format (`CSV` & `Excel`)**: Membaca berkas `Piutang.xls` dimulai dari baris *header* ke-4 (`header=3`) dengan mekanisme *fallback* otomatis dari `pd.read_csv()` ke `pd.read_excel()`.
* **Deteksi Kolom `Kode Pelanggan` Dinamis**: Jika kolom `Kode Pelanggan` tidak bernama eksplisit akibat pergeseran kolom ekspor ERP, sistem menelusuri kolom secara mundur dari posisi `No. Faktur` dan melewati kolom mata uang (`IDR`) secara otomatis.
* **Normalisasi Hierarki & Nilai Numerik**: Melakukan *forward-fill* (`ffill`) pada `Kode Pelanggan` untuk mengisi baris grup pelanggan, membuang baris sub-total/kosong tanpa `No. Faktur`, serta membersihkan akhiran desimal (`.0`, `,00`, `.00`).
* **Deteksi Otomatis Dua Mode `Umur JT`**:
  * **Mode Baru**: Mengenali data umur jatuh tempo yang mengandung teks `"hari"` (contoh: `"15 hari"`).
  * **Mode Lama**: Mengenali format angka desimal mentah, membersihkan pemisah ribuan/desimal, dan mengonversinya menjadi bilangan bulat.

### 3. Pemisahan Depo & Penyaringan Produk Dinamis (`2_Split.py`)
* **Pemetaan *Prefix* Multi-Sheet**: Memisahkan baris transaksi ke masing-masing lembar kerja (*worksheet*) Excel berdasarkan kecocokan awalan (*prefix*) pada kolom `Nama Penjual` menggunakan aturan `[PREFIX_TO_SHEET]` atau `[FILTER_PREFIX]`.
* **Penyaringan Kata Kunci Berbasis *Regex* (`[PRODUCT_FILTERS]`)**: Memfilter baris data berdasarkan satu atau beberapa kata kunci pada kolom target tertentu (misalnya menyaring kata `CASH` pada `Nama Pelanggan` atau `IRC, ZN` pada `Nama Kontak`).

### 4. Analitik *Aging Schedule*, Filter Fraud & Visualisasi Excel (`3_Analytics.py`)
* **Penyaringan Pengecualian Fraud/Kasus (`gen_fraud`)**: Mengecualikan transaksi dengan pola nama penjual tertentu (misalnya `FRAUD` atau `FRAUD|CASE`) menggunakan pencocokan ekspresi reguler (*regex* *case-insensitive*).
* **Normalisasi Tanggal Bahasa Indonesia**: Menerjemahkan singkatan bulan bahasa Indonesia (`Mei`, `Agu`/`Ags`, `Okt`, `Nop`, `Des`) dan mengenali berbagai format tanggal (`dd/mm/yyyy`, `dd MMM yyyy`, `yyyy-mm-dd HH:MM:SS`, `dd-mm-yyyy`) agar pengurutan kronologis faktur akurat.
* **Lembar Analisis Jatuh Tempo Opsional (`<Sheet> JT`)**: Jika parameter `Aktif = YES` pada seksi `[JT]`, sistem otomatis menduplikasi analisis khusus untuk faktur yang sudah jatuh tempo (`Umur_Int >= 0`) ke lembar kerja terpisah.
* **Laporan Eksekutif Terpadu di Setiap Lembar Kerja**:
  * **Ringkasan Performa Sales**: Tabel *pivot* sisa piutang per penjual yang terbagi dalam 4 keranjang umur piutang (`-30 - 0 Hari`, `1-31 Hari`, `32-60 Hari`, `> 60 Hari`) beserta kolom `Grand Total` berpenanda kuning (`#FFFF00`).
  * **Top 10 Outlet Macet (`> 60 Hari`)**: Peringkat 10 toko/pelanggan dengan akumulasi sisa piutang lewat jatuh tempo di atas 60 hari tertinggi.
  * **Grafik Lingkaran (*Pie Chart*)**: Visualisasi porsi piutang per penjual berukuran `550x400 px` lengkap dengan label nilai dan persentase.
  * **Rincian Detil Transaksi Terkelompok**: Daftar faktur yang dikelompokkan per penjual, diurutkan berdasarkan tanggal faktur terlama, dilengkapi sub-header total piutang per penjual serta penanda warna merah muda (`#FFC7CE`) untuk faktur dalam kategori `> 60 Hari`.

---

## 🔄 Arsitektur & Alur Kerja

```text
+-------------------------------------------------------------------------+
|                        INPUT: Piutang.xls (Root)                        |
|        (Ekspor Data Piutang Mentah ERP / Accurate - Header Baris 4)     |
+------------------------------------+------------------------------------+
                                     |
                                     v
+-------------------------------------------------------------------------+
|                    ORKESTRATOR: Analisis_AR.py                          |
|  1. Validasi keberadaan Piutang.xls & 7 berkas wajib di folder dapur/   |
|  2. Bersihkan berkas *temp.xlsx lama & salin Piutang.xls ke dapur/      |
|  3. Tampilkan Menu CLI (1: Selatan, 2: Utara, 3: IRC ZN, 4: Customized) |
+------------------------------------+------------------------------------+
                                     |
                                     v
+-------------------------------------------------------------------------+
|               TAHAP 1: dapur/1_CleanerAcc.py                            |
|  - Baca CSV/Excel (header=3) & validasi 9 kolom wajib                   |
|  - Deteksi heuristik kolom 'Kode Pelanggan' (lewati kolom 'IDR')        |
|  - Forward-fill (ffill) Kode Pelanggan & hapus baris tanpa No. Faktur   |
|  - Deteksi Mode Umur JT (MODE BARU berteks 'hari' vs MODE LAMA angka)   |
|  - Output sementara: dapur/Piutang_cleantemp.xlsx                       |
+------------------------------------+------------------------------------+
                                     |
                                     v
+-------------------------------------------------------------------------+
|               TAHAP 2: dapur/2_Split.py <file.conf>                     |
|  - Muat aturan [PREFIX_TO_SHEET] / [FILTER_PREFIX] & [PRODUCT_FILTERS]  |
|  - Saring baris berdasarkan regex kata kunci pada kolom target (opsional)|
|  - Pisahkan data ke multi-sheet berdasarkan prefix 'Nama Penjual'       |
|  - Output sementara: dapur/cleandepotemp.xlsx                           |
+------------------------------------+------------------------------------+
                                     |
                                     v
+-------------------------------------------------------------------------+
|               TAHAP 3: dapur/3_Analytics.py <file.conf>                 |
|  - Saring pengecualian pola 'Nama Penjual' berdasarkan [GENERAL]        |
|    gen_fraud (contoh: FRAUD|CASE)                                       |
|  - Buat sheet tambahan '<NamaSheet> JT' (Umur >= 0) bila [JT] Aktif=YES |
|  - Konversi tanggal Indonesia, nominal Rupiah, & 4 kategori Aging       |
|  - Susun Pivot Sales, Top 10 Outlet (>60 Hari), Pie Chart, & Rincian    |
|  - Output sementara: dapur/Laporan_Analisis_Piutang.xlsx                |
+------------------------------------+------------------------------------+
                                     |
                                     v
+-------------------------------------------------------------------------+
|                  FINALISASI & PEMBERSIHAN OTOMATIS                      |
|  - Salin ke root: Laporan_Analisis_Piutang_<Area>.xlsx                  |
|  - Hapus seluruh file temporer di dalam folder dapur/                   |
+-------------------------------------------------------------------------+
```

---

## 💻 Teknologi yang Digunakan

| Komponen / Pustaka | Peran dalam Sistem |
| :--- | :--- |
| **Python 3.x** | Bahasa pemrograman utama untuk orkestrasi proses dan manipulasi data. |
| **`pandas`** | Mesin pemrosesan tabel data (*DataFrame*), pembacaan CSV/Excel, *forward-fill*, *filtering regex*, dan pembuatan *pivot table* agregasi piutang. |
| **`numpy`** | Operasi numerik dasar dan penanganan tipe data *floating-point* / *NaN*. |
| **`xlsxwriter`** | Mesin pembuat laporan Excel tingkat lanjut untuk pengaturan lebar kolom otomatis (*auto-fit*), format mata uang Rupiah, warna *conditional cell*, penggabungan sel (*merge range*), dan penyematan grafik *Pie Chart*. |
| **`openpyxl` / `xlrd`** | *Engine* pembaca berkas Excel (`.xlsx` dan `.xls`) yang digunakan oleh `pandas.read_excel()`. |
| **`configparser`** | Modul standar Python untuk mengurai parameter konfigurasi operasional berbasis berkas `.conf` (format INI). |
| **`subprocess` & `shutil`** | Eksekusi berantai antar-skrip secara terisolasi dan manajemen penyalinan/penghapusan berkas lintas direktori. |

---

## 📁 Struktur Direktori

Pastikan susunan berkas di dalam repositori mengikuti struktur berikut agar validasi `Analisis_AR.py` berjalan lancar:

```text
📂 project-root/
│
├── 📄 __init__.py                      # Penanda paket & metadata maintainer (me@ihfandica.com)
├── 📄 Analisis_AR.py                   # Orkestrator CLI utama (Entry Point)
├── 📄 Piutang.xls                      # Berkas input data piutang mentah (diletakkan oleh pengguna)
├── 📄 README.md                        # Dokumentasi repositori
│
├── 📊 Laporan_Analisis_Piutang_*.xlsx  # Berkas hasil analisis akhir (dihasilkan otomatis)
│
└── 📂 dapur/                           # Direktori internal pemrosesan pipeline & konfigurasi
    ├── 📄 1_CleanerAcc.py              # Modul Tahap 1: Pembersihan data & deteksi kolom/umur JT
    ├── 📄 2_Split.py                   # Modul Tahap 2: Pemisahan sheet depo & filter produk
    ├── 📄 3_Analytics.py               # Modul Tahap 3: Kalkulasi aging, filter fraud & pembuatan laporan Excel
    ├── ⚙️ selatan.conf                 # Konfigurasi Area Selatan (Yogya, Solo, Magelang)
    ├── ⚙️ utara.conf                   # Konfigurasi Area Utara (Kendari, Makassar, Pati, Palu, Purwokerto, Semarang, Tegal)
    ├── ⚙️ irczn.conf                   # Konfigurasi khusus tim sales SR-* & produk IRC/ZN
    └── ⚙️ customized.conf              # Konfigurasi kustom untuk penyaringan transaksi/area spesifik
```

---

## 📋 Prasyarat Sistem

1. **Lingkungan Runtime**:
   * **Python**: Versi `3.8` atau lebih baru.
   * **Sistem Operasi**: Windows, macOS, atau Linux.
2. **Spesifikasi Berkas Input (`Piutang.xls`)**:
   * Berkas harus diberi nama persis **`Piutang.xls`** dan diletakkan di direktori akar (*root directory*) sejajar dengan `Analisis_AR.py`.
   * Nama kolom tabel harus berada pada **baris ke-4** (`header=3`, indeks berbasis 0).
   * Wajib memiliki **9 kolom bernama** berikut (nama peka huruf besar/kecil sesudah pemangkasan spasi):
     * `No. Faktur`
     * `Tgl Faktur`
     * `Jatuh Tempo`
     * `Nilai Faktur`
     * `Sisa Piutang`
     * `Umur JT`
     * `Nama Pelanggan`
     * `Nama Penjual`
     * `Nama Kontak`
   * Memiliki kolom **`Kode Pelanggan`** (baik bernama eksplisit `"Kode Pelanggan"` maupun kolom tanpa nama di sebelah kiri kolom `"No. Faktur"`).

---

## ⚙️ Panduan Instalasi & Konfigurasi

1. **Kloning Repositori**
   ```bash
   git clone <url-repositori-anda>
   cd <nama-folder-repositori>
   ```

2. **Buat dan Aktifkan *Virtual Environment* (Disarankan)**
   ```bash
   # Windows
   python -m venv venv
   venv\Scripts\activate

   # macOS / Linux
   python3 -m venv venv
   source venv/bin/activate
   ```

3. **Instal Pustaka Dependensi**
   ```bash
   pip install pandas numpy xlsxwriter openpyxl xlrd
   ```
   > **Catatan**: Pustaka `xlrd` disarankan turut diinstal apabila berkas `Piutang.xls` Anda berformat biner Excel 97-2003 (`.xls` murni, bukan CSV berlabel `.xls`).

4. **Verifikasi Kelengkapan Folder `dapur/`**
   Pastikan ketujuh berkas (`1_CleanerAcc.py`, `2_Split.py`, `3_Analytics.py`, `selatan.conf`, `utara.conf`, `irczn.conf`, dan `customized.conf`) berada di dalam sub-folder `dapur/`.

---

## 🚀 Cara Penggunaan

### 1. Eksekusi Otomatis Melalui Orkestrator (Direkomendasikan)

1. Salin atau letakkan berkas data mentah **`Piutang.xls`** ke folder utama (*root*).
2. Jalankan skrip orkestrator dari terminal atau *Command Prompt*:
   ```bash
   python Analisis_AR.py
   ```
3. Pilih nomor area operasional yang ingin dianalisis saat menu interaktif muncul:
   ```text
   --> Membersihkan file sementara di folder dapur...
   --> Menyalin Piutang.xls ke folder dapur...
   --> 
   Pilih Area Operasional:
   --> 1. Selatan
   --> 2. Utara
   --> 3. IRC ZN
   --> 4. Customized -> Menggunakan aturan konfigurasi pada dapur/customized.conf
   --> Masukkan pilihan (1/2/3/4): 
   ```
4. Ketik angka pilihan (`1`, `2`, `3`, atau `4`) lalu tekan **Enter**.
5. Setelah proses selesai, berkas laporan akhir akan tersedia di folder utama dengan nama sesuai pilihan Anda:
   * Pilihan `1`: `Laporan_Analisis_Piutang_Selatan.xlsx`
   * Pilihan `2`: `Laporan_Analisis_Piutang_Utara.xlsx`
   * Pilihan `3`: `Laporan_Analisis_Piutang_IRC_ZN.xlsx`
   * Pilihan `4`: `Laporan_Analisis_Piutang_Customized.xlsx`

### 2. Eksekusi Modul Secara Manual (Untuk Debugging)

Jika Anda ingin menguji từng tahapan secara terpisah di dalam folder `dapur/`:

```bash
cd dapur

# Tahap 1: Menghasilkan Piutang_cleantemp.xlsx dari Piutang.xls
python 1_CleanerAcc.py

# Tahap 2: Menghasilkan cleandepotemp.xlsx menggunakan konfigurasi tertentu
python 2_Split.py selatan.conf

# Tahap 3: Menghasilkan Laporan_Analisis_Piutang.xlsx menggunakan konfigurasi tertentu
python 3_Analytics.py selatan.conf
```

---

## 🔧 Referensi Konfigurasi (`.conf`)

Setiap berkas `.conf` di dalam folder `dapur/` menggunakan struktur standar INI. Berikut adalah rincian seluruh bagian dan parameter yang didukung oleh `2_Split.py` dan `3_Analytics.py`:

### 1. Parameter Konfigurasi

| Seksi | Kunci / Parameter | Digunakan Oleh | Wajib? | Deskripsi & Perilaku |
| :--- | :--- | :--- | :---: | :--- |
| **`[PREFIX_TO_SHEET]`** atau **`[FILTER_PREFIX]`** | `<Prefix> = <Nama Sheet>` | `2_Split.py` | **Ya** (Salah satu) | Memetakan awalan teks pada kolom `Nama Penjual` ke nama *worksheet* tujuan. Peka huruf besar/kecil (*case-sensitive*). Beberapa *prefix* berbeda dapat diarahkan ke nama *sheet* yang sama (contoh: `SL- = Solo` dan `SL. = Solo`). Catatan: Jika kedua seksi ada di satu berkas, `[PREFIX_TO_SHEET]` diprioritaskan. |
| **`[PRODUCT_FILTERS]`** | `COLUMN_NAME` | `2_Split.py` | Tidak | Nama kolom yang akan disaring sebelum pemisahan depo (contoh: `Nama Pelanggan` atau `Nama Kontak`). |
| **`[PRODUCT_FILTERS]`** | `KEYWORDS` | `2_Split.py` | Tidak | Daftar kata kunci yang dipisahkan tanda koma (contoh: `IRC, ZN` atau `CASH`). Dikonversi menjadi pola *regex* `OR` (`IRC\|ZN`). Hanya baris yang memuat kata kunci tersebut yang akan diproses. |
| **`[JT]`** | `Aktif` | `3_Analytics.py` | Tidak | Nilai boolean (`YES`/`NO`, `True`/`False`, `1`/`0`). Jika `YES`, sistem menambahkan lembar kerja `<NamaSheet> JT` yang hanya berisi faktur dengan `Umur JT >= 0` hari. Default: `False`. |
| **`[GENERAL]`** | `gen_fraud` | `3_Analytics.py` | Tidak | Pola *regex* pengecualian pada kolom `Nama Penjual` (*case-insensitive*). Baris penjual yang cocok dengan pola ini (contoh: `FRAUD` atau `FRAUD\|CASE`) dihapus dari perhitungan laporan. |

### 2. Ringkasan Profil Konfigurasi Bawaan

| Berkas Konfigurasi | Pemetaan Sheet (`Nama Penjual`) | Filter Kolom (`[PRODUCT_FILTERS]`) | Mode Sheet JT (`[JT]`) | Filter Pengecualian (`gen_fraud`) |
| :--- | :--- | :--- | :---: | :--- |
| **`selatan.conf`** | `YY-` → **Yogya**<br>`SL-`, `SL.` → **Solo**<br>`MGL.` → **Magelang** | *Tidak aktif* | `YES` | `FRAUD` |
| **`utara.conf`** | `KDI-` → **Kendari**, `MKS-` → **Makassar**, `PA-` → **Pati**, `PLU-` → **Palu**, `PW-` → **Purwokerto**, `SMG-` → **Semarang**, `TGL-` → **Tegal** | *Tidak aktif* | `YES` | `FRAUD` |
| **`irczn.conf`** | 9 *prefix* `SR-*` (`SR-ADM KTR`, `SR-AGUS`, `SR-AGUS 1`, `SR-AGUS SULISTIONO`, `SR-ALBERTUS`, `SR-BETA`, `SR-DWI IRC`, `SR-ENDI`, `SR-WAHYU`) → **IRC ZN** | Kolom: `Nama Kontak`<br>Kata Kunci: `IRC, ZN` | `YES` | `FRAUD` |
| **`customized.conf`** | `YY-` → **Yogya**<br>`SL-`, `SL.` → **Solo**<br>`MGL.` → **Magelang** | Kolom: `Nama Pelanggan`<br>Kata Kunci: `CASH` | `No` | `FRAUD\|CASE` |

### 3. Contoh Penulisan Berkas `.conf`

```ini
[PREFIX_TO_SHEET]
YY- = Yogya
SL- = Solo
SL. = Solo
MGL. = Magelang

[PRODUCT_FILTERS]
COLUMN_NAME = Nama Pelanggan
KEYWORDS = CASH

[JT]
Aktif = YES

[GENERAL]
gen_fraud = FRAUD|CASE
```

---

## 🛠️ Penanganan Masalah Umum (Troubleshooting)

| Pesan Error / Gejala | Penyebab Utama | Solusi Langkah-demi-Langkah |
| :--- | :--- | :--- |
| `--> Error: File 'Piutang.xls' tidak ditemukan di folder utama.` | Berkas sumber belum diletakkan di direktori akar atau namanya tidak sesuai. | Pastikan berkas bernama persis `Piutang.xls` (bukan `piutang.xlsx` atau `Piutang (1).xls`) dan berada di folder yang sama dengan `Analisis_AR.py`. |
| `ValueError: Kolom berikut tidak ditemukan di file sumber: [...]` | Posisi baris *header* di berkas `Piutang.xls` bukan berada di baris ke-4 atau ada nama kolom yang berubah saat ekspor. | Buka `Piutang.xls`, pastikan judul kolom (`No. Faktur`, `Tgl Faktur`, dll.) berada tepat di **baris ke-4** (terdapat 3 baris judul/kop di atasnya). |
| `--> Peringatan: Tidak ada data yang cocok dengan konfigurasi ini.` | Tidak ada baris data yang memenuhi syarat *prefix* `Nama Penjual` dan/atau syarat `KEYWORDS` pada `[PRODUCT_FILTERS]`. | 1. Periksa isi kolom `Nama Penjual` pada `Piutang.xls` apakah *prefix*-nya sesuai dengan `.conf`.<br>2. Jika menggunakan `customized.conf`, periksa apakah `[PRODUCT_FILTERS]` (`KEYWORDS = CASH`) terlalu membatasi data; kosongkan nilai `COLUMN_NAME` dan `KEYWORDS` bila ingin memproses seluruh pelanggan. |
| `[PREFIX_TO_SHEET] tidak berubah saat mengedit [FILTER_PREFIX]` pada `customized.conf` / `irczn.conf` | Di dalam `2_Split.py`, jika sebuah berkas `.conf` memiliki seksi `[PREFIX_TO_SHEET]` sekaligus `[FILTER_PREFIX]` kosong, maka sistem hanya membaca `[PREFIX_TO_SHEET]`. | Tambahkan atau ubah daftar *prefix* langsung di bawah blok `[PREFIX_TO_SHEET]`, atau hapus blok `[PREFIX_TO_SHEET]` jika ingin menggunakan `[FILTER_PREFIX]`. |
| `ImportError: Missing optional dependency 'xlrd'` | Berkas `Piutang.xls` berformat Excel `.xls` lawas tetapi pustaka `xlrd` belum terpasang. | Jalankan perintah `pip install xlrd` di dalam lingkungan Python Anda. |
| `PermissionError: [Errno 13] Permission denied: 'Laporan_Analisis_Piutang_*.xlsx'` | Berkas laporan hasil analisis sebelumnya sedang terbuka di aplikasi Microsoft Excel. | Tutup jendela Microsoft Excel yang sedang membuka berkas laporan tersebut, lalu jalankan ulang `python Analisis_AR.py`. |

---

## 🤝 Kontribusi & Lisensi

### Panduan Kontribusi
1. Lakukan *Fork* pada repositori ini.
2. Buat *branch* fitur atau perbaikan baru (`git checkout -b fitur/nama-fitur`).
3. Pastikan perubahan telah diuji terhadap keempat profil konfigurasi (`selatan.conf`, `utara.conf`, `irczn.conf`, dan `customized.conf`).
4. Lakukan *Commit* (`git commit -m "Menambahkan fitur X"`) dan *Push* ke *branch* Anda (`git push origin fitur/nama-fitur`).
5. Buat *Pull Request* ke *branch* utama repositori ini.

## 📜 Lisensi

Proyek ini dikembangkan untuk keperluan internal perusahaan. Silakan sesuaikan dengan kebutuhan organisasi Anda.

---

*Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)*
