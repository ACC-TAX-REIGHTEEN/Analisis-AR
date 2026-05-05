import pandas as pd
import numpy as np

def clean_data_autofit(input_file, output_file):
    print(f"Sedang memproses file: {input_file}...")
    
    # 1. Membaca file (Coba CSV dulu, lalu Excel)
    try:
        df = pd.read_csv(input_file, header=3)
    except:
        df = pd.read_excel(input_file, header=3)

    # 2. Memilih kolom yang relevan (Index C=2, D=3, dst...)
    target_indices = [2, 3, 4, 7, 8, 10, 11, 12, 13, 14]
    df_clean = df.iloc[:, target_indices].copy()

    # 3. Memberi nama kolom baru
    new_columns = [
        'Kode Pelanggan', # C
        'No. Faktur',     # D
        'Tgl Faktur',     # E
        'Jatuh Tempo',    # H
        'Nilai Faktur',   # I
        'Sisa Piutang',   # K
        'Umur JT',        # L
        'Nama Pelanggan', # M
        'Nama Penjual',   # N
        'Nama Kontak'     # O
    ]
    df_clean.columns = new_columns

    # 4. Logika Pengisian Kode Pelanggan (Forward Fill)
    df_clean['Kode Pelanggan'] = df_clean['Kode Pelanggan'].ffill()

    # 5. Membersihkan baris yang bukan data transaksi (Hapus jika No. Faktur kosong)
    df_clean = df_clean.dropna(subset=['No. Faktur'])

    # 6. Fungsi Pembersihan Desimal (,00 atau .0)
    def format_clean(val):
        if pd.isna(val):
            return ""
        s = str(val)
        # Hapus .0 di belakang (format standar)
        if s.endswith('.0'):
            return s[:-2]
        # Hapus ,00 di belakang (format Indonesia)
        if s.endswith(',00'):
            return s[:-3]
        return s

    # 7. Terapkan pembersihan ke kolom target
    cols_to_clean = ['Kode Pelanggan', 'Nilai Faktur', 'Sisa Piutang']
    for col in cols_to_clean:
        df_clean[col] = df_clean[col].apply(format_clean)

    # Reset nomor baris agar rapi
    df_clean.reset_index(drop=True, inplace=True)

    # 8. Simpan ke Excel dengan AUTO-FIT Column Width
    try:
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            # Tulis Dataframe ke Excel
            df_clean.to_excel(writer, index=False, sheet_name='Data Bersih')
            
            # Akses Workbook dan Worksheet untuk formatting
            workbook = writer.book
            worksheet = writer.sheets['Data Bersih']
            
            # Loop setiap kolom untuk mengatur lebarnya
            for i, col in enumerate(df_clean.columns): 
                panjang_data = df_clean[col].dropna().astype(str).map(len)
                max_data_len = panjang_data.max() if not panjang_data.empty else 0
                max_len = max(int(max_data_len), len(str(col))) + 2
                worksheet.set_column(i, i, max_len)    
                
        print(f"SUKSES! File tersimpan rapi di: {output_file}")
        
    except Exception as e:
        print(f"Error saat menyimpan file: {e}")

    return df_clean

# --- BAGIAN EKSEKUSI ---
# Ganti nama file input sesuai file asli Anda
input_filename = 'ExportFile.xls' 
output_filename = 'ExportFile_cleantemp.xlsx'

# Jalankan fungsi
if __name__ == "__main__":
    clean_data_autofit(input_filename, output_filename)