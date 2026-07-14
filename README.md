# 📊 Automated AR (Accounts Receivable) Analytics System

Sistem otomatisasi pipeline data untuk membersihkan, memisahkan, dan menganalisis data piutang pelanggan (*Aging Schedule*) berdasarkan area operasional secara dinamis menggunakan Python.

Orkestrator utama dari sistem ini adalah **`Analisis_AR.py`**, yang mengatur seluruh jalannya skrip pemrosesan data di latar belakang.

---

## 🛠️ Fitur Utama

* **Otomatisasi Pipeline (All-in-One):** Cukup jalankan satu file konduktor untuk mengeksekusi pembersihan, penyaringan, hingga pembuatan laporan.
* **Data Cleaning Otomatis:** Mengurai file ekspor mentah, menangani *missing values*, melakukan *forward fill* kode pelanggan, dan membersihkan format angka.
* **Penyaringan Berbasis Area (.conf):** Memisahkan data berdasarkan prefix nama penjual/sales menggunakan file konfigurasi eksternal.
* **Laporan Analitik Interaktif:** Menghasilkan file Excel (`.xlsx`) yang dilengkapi dengan:
    * Ringkasan performa *aging* per Salesman (Kategori: -30-0 hari, 1-31 hari, 32-60 hari, >60 hari).
    * Top 10 Outlet dengan piutang macet (>60 Hari).
    * Grafik lingkaran (*Pie Chart*) porsi piutang per sales.
    * Rincian detail transaksi per sales yang terkelompok rapi.

---

## 📁 Struktur Folder

Untuk menjalankan sistem ini dengan sukses, pastikan struktur direktori kamu mengikuti format berikut:
```text
📂 project-root/
│
├── 📄 Analisis_AR.py                   # <-- Orkestrator Utama (Jalankan file ini)
├── 📄 Piutang.xls                   # <-- File Data Mentah Mandiri (Input)
│
└── 📂 dapur/                           # Folder internal pemrosesan
    ├── 📄 1_clean_all_piutang_withNOPEL.py
    ├── 📄 2_separate_depo_customized.py
    ├── 📄 2_separate_depo_IRCZN.py
    ├── 📄 2_separate_depo_selatan.py
    ├── 📄 2_separate_depo_utara.py
    ├── 📄 3_analytics_salesman.py
    ├── 📄 customized.conf
    ├── 📄 irczn.conf
    ├── 📄 selatan.conf
    └── 📄 utara.conf
```
---

## 🔄 Alur Kerja Sistem (Pipeline Workflow)

1. **`Analisis_AR.py` (Orkestrator)** memvalidasi ketersediaan file input (`Piutang.xls`) dan kelengkapan skrip di folder `dapur/`.
2. File mentah disalin ke folder `dapur/` untuk diolah.
3. User memilih Area Operasional lewat menu interaktif (Selatan, Utara, IRC ZN, atau Customized).
4. **Skrip Tahap 1 (`1_clean_...`)** membersihkan data mentah menjadi `Piutang_cleantemp.xlsx`.
5. **Skrip Tahap 2 (`2_separate_...`)** memfilter data berdasarkan aturan file konfigurasi (`.conf`) terkait dan menghasilkan `cleandepotemp.xlsx`.
6. **Skrip Tahap 3 (`3_analytics_...`)** melakukan agregasi data, membuat grafik, format mata uang, dan menyusun laporan akhir.
7. Orkestrator memindahkan hasil akhir ke folder utama dengan penamaan area (contoh: `Laporan_Analisis_Piutang_Selatan.xlsx`) dan membersihkan file temporer secara otomatis.

---

## 💻 Prasyarat & Instalasi

Pastikan Anda sudah menginstal Python 3.x dan pustaka (*libraries*) yang dibutuhkan.

```bash
pip install pandas numpy xlsxwriter openpyxl

```

---

## 🚀 Cara Penggunaan

1. Letakkan file data piutang mentah Anda di **folder utama (root)** dengan nama **`Piutang.xls`**.
2. Jalankan terminal / command prompt di folder utama tersebut.
3. Eksekusi skrip konduktor:
```bash
python Analisis_AR.py

```


4. Pilih nomor area operasional yang ingin dianalisis saat diminta di terminal:
```text
Pilih Area Operasional:
1. Selatan
2. Utara
3. IRC ZN
4. Customized.
Masukkan pilihan (1/2/3/4): 

```


5. Tunggu hingga muncul pesan sukses. Hasil analisis akan muncul di folder utama Anda!

---

## ⚙️ Kustomisasi Konfigurasi (`.conf`)

Anda dapat mengatur pembagian *sheet* operasional maupun menyalakan/mematikan fitur Analisis Jatuh Tempo (JT) melalui file `.conf` di dalam folder `dapur`.

Contoh isi `customized.conf`:

* **`[PREFIX_TO_SHEET]`**: Menentukan sales dengan awalan tertentu masuk ke sheet mana (Contoh: `YY- = Yogya`).
* **`[PRODUCT_FILTERS]`**: Memfilter kata kunci khusus pada kolom tertentu (Contoh: menyaring kata `CASH`).
* **`[JT]`**: Mengaktifkan (`YES`) atau menonaktifkan (`NO`) lembar analisis jatuh tempo tambahan di laporan akhir.
