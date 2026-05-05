import pandas as pd
import configparser
import os
import sys

def load_config(config_file):
    if not os.path.exists(config_file):
        print(f"File konfigurasi '{config_file}' tidak ditemukan.")
        sys.exit(1)

    config = configparser.ConfigParser()
    config.optionxform = str
    config.read(config_file)

    if 'PREFIX_TO_SHEET' not in config:
        print("Bagian [PREFIX_TO_SHEET] tidak ditemukan di file konfigurasi.")
        sys.exit(1)

    grouped_filters = {}
    for prefix, sheet_name in config['PREFIX_TO_SHEET'].items():
        clean_prefix = prefix.strip()
        clean_sheet = sheet_name.strip().replace('"', '').replace("'", "")
        
        if clean_sheet not in grouped_filters:
            grouped_filters[clean_sheet] = []
        grouped_filters[clean_sheet].append(clean_prefix)

    prefix_map = {k: tuple(v) for k, v in grouped_filters.items()}

    product_settings = {}
    if 'PRODUCT_FILTERS' in config:
        product_settings['col_name'] = config['PRODUCT_FILTERS'].get('COLUMN_NAME', 'Nama Barang')
        
        raw_keywords = config['PRODUCT_FILTERS'].get('KEYWORDS', '')
        keyword_list = [k.strip() for k in raw_keywords.split(',') if k.strip()]
        product_settings['regex_pattern'] = '|'.join(keyword_list)
    else:
        product_settings['col_name'] = None
        product_settings['regex_pattern'] = None

    return prefix_map, product_settings

def split_depo_final(input_file, output_file, config_file):
    prefix_map, product_filters = load_config(config_file)

    try:
        df = pd.read_csv(input_file)
    except:
        try:
            df = pd.read_excel(input_file)
        except:
            print("Gagal membaca file input.")
            return

    search_col = product_filters['col_name']
    search_regex = product_filters['regex_pattern']

    if search_col and search_regex:
        if search_col in df.columns:
            filter_mask = df[search_col].astype(str).str.contains(search_regex, na=False, regex=True)
            df = df[filter_mask]
        else:
            print(f"Kolom '{search_col}' tidak ditemukan di data.")

    target_col = 'Nama Penjual'
    if target_col not in df.columns:
        print(f"Kolom '{target_col}' tidak ditemukan.")
        return

    try:
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            total_saved = 0
            
            for sheet_name, prefixes in prefix_map.items():
                mask = df[target_col].astype(str).str.strip().str.startswith(prefixes)
                df_filtered = df[mask].copy()

                if not df_filtered.empty:
                    df_filtered.to_excel(writer, sheet_name=sheet_name, index=False)
                    total_saved += 1

                    worksheet = writer.sheets[sheet_name]
                    for i, col in enumerate(df_filtered.columns): 
                        panjang_data = df_filtered[col].dropna().astype(str).map(len)
                        max_data_len = panjang_data.max() if not panjang_data.empty else 0
                        max_len = max(int(max_data_len), len(str(col))) + 2
                        worksheet.set_column(i, i, max_len)
            
            if total_saved > 0:
                print(f"Sukses! File tersimpan: {output_file}")
            else:
                print("Tidak ada data yang cocok.")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    input_filename = 'ExportFile_cleantemp.xlsx' 
    output_filename = 'cleandepotemp.xlsx'
    config_filename = 'customized.conf'

    split_depo_final(input_filename, output_filename, config_filename)