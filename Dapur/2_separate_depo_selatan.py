import pandas as pd
import configparser
import os

def split_depo_config(input_file, output_file, config_file):
    if not os.path.exists(config_file):
        print(f"Error: File konfigurasi '{config_file}' tidak ditemukan.")
        return

    config = configparser.ConfigParser()
    config.optionxform = str
    config.read(config_file)

    if 'FILTER_PREFIX' not in config:
        print("Error: Section [FILTER_PREFIX] tidak ditemukan di file config.")
        return

    sheet_map = {}
    for prefix, sheet_name in config['FILTER_PREFIX'].items():
        if sheet_name not in sheet_map:
            sheet_map[sheet_name] = []
        sheet_map[sheet_name].append(prefix)

    try:
        df = pd.read_csv(input_file)
    except:
        try:
            df = pd.read_excel(input_file)
        except Exception as e:
            print(f"Gagal membaca file input: {e}")
            return

    target_col = 'Nama Penjual'
    if target_col not in df.columns:
        print(f"Error: Kolom '{target_col}' tidak ditemukan.")
        return

    try:
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            for sheet_name, prefixes in sheet_map.items():
                tuple_prefixes = tuple(prefixes)
                
                mask = df[target_col].astype(str).str.strip().str.startswith(tuple_prefixes)
                df_filtered = df[mask].copy()

                df_filtered.to_excel(writer, sheet_name=sheet_name, index=False)
                
                print(f"Sheet '{sheet_name}': {len(df_filtered)} baris (Prefix: {prefixes})")

                worksheet = writer.sheets[sheet_name]
                for i, col in enumerate(df_filtered.columns): 
                    panjang_data = df_filtered[col].dropna().astype(str).map(len)
                    max_data_len = panjang_data.max() if not panjang_data.empty else 0
                    max_len = max(int(max_data_len), len(str(col))) + 2
                    worksheet.set_column(i, i, max_len)

        print(f"\nSUKSES! File tersimpan sebagai: {output_file}")
        
    except Exception as e:
        print(f"Terjadi kesalahan: {e}")

input_filename = 'Piutang_cleantemp.xlsx' 
output_filename = 'cleandepotemp.xlsx'
config_filename = 'selatan.conf'

if __name__ == "__main__":
    split_depo_config(input_filename, output_filename, config_filename)