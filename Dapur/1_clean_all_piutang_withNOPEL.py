import pandas as pd
import numpy as np

def clean_data_autofit(input_file, output_file):
    print(f"Sedang memproses file: {input_file}...")
    
    try:
        df = pd.read_csv(input_file, header=3)
    except:
        df = pd.read_excel(input_file, header=3)

    target_indices = [2, 3, 4, 7, 8, 10, 11, 12, 13, 14]
    df_clean = df.iloc[:, target_indices].copy()

    new_columns = [
        'Kode Pelanggan', 
        'No. Faktur',     
        'Tgl Faktur',     
        'Jatuh Tempo',    
        'Nilai Faktur',   
        'Sisa Piutang',   
        'Umur JT',        
        'Nama Pelanggan', 
        'Nama Penjual',   
        'Nama Kontak'    
    ]
    df_clean.columns = new_columns
    df_clean['Kode Pelanggan'] = df_clean['Kode Pelanggan'].ffill()
    df_clean = df_clean.dropna(subset=['No. Faktur'])
    
    def format_clean(val):
        if pd.isna(val):
            return ""
        s = str(val)

        if s.endswith('.0'):
            return s[:-2]
     
        if s.endswith(',00'):
            return s[:-3]
        return s

    cols_to_clean = ['Kode Pelanggan', 'Nilai Faktur', 'Sisa Piutang']
    for col in cols_to_clean:
        df_clean[col] = df_clean[col].apply(format_clean)

    df_clean.reset_index(drop=True, inplace=True)
    
    try:
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
         
            df_clean.to_excel(writer, index=False, sheet_name='Data Bersih')
    
            workbook = writer.book
            worksheet = writer.sheets['Data Bersih']
       
            for i, col in enumerate(df_clean.columns): 
                panjang_data = df_clean[col].dropna().astype(str).map(len)
                max_data_len = panjang_data.max() if not panjang_data.empty else 0
                max_len = max(int(max_data_len), len(str(col))) + 2
                worksheet.set_column(i, i, max_len)    
                
        print(f"SUKSES! File tersimpan rapi di: {output_file}")
        
    except Exception as e:
        print(f"Error saat menyimpan file: {e}")

    return df_clean

input_filename = 'ExportFile.xls' 
output_filename = 'ExportFile_cleantemp.xlsx'

if __name__ == "__main__":
    clean_data_autofit(input_filename, output_filename)
