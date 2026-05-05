import pandas as pd
import xlsxwriter
from datetime import datetime
import os
import configparser

def cek_status_jt(conf_filename):
    """Mengecek apakah fitur [JT] diaktifkan (YES) di dalam file config"""
    if not os.path.exists(conf_filename):
        return False
    
    try:
        config = configparser.ConfigParser()
        config.read(conf_filename)
   
        if config.has_section('JT'):
         
            return config.getboolean('JT', 'Aktif', fallback=False)
            
    except Exception as e:
        print(f"Peringatan: Gagal membaca {conf_filename} -> {e}")
    
    return False

def analisa_piutang_satu_file_fix(input_file, output_file, is_jt_active):
    print(f"Membaca file sumber: {input_file}...")
    
    try:
        all_sheets = pd.read_excel(input_file, sheet_name=None)
    except Exception as e:
        print(f"Error membaca file: {e}")
        return

    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        workbook = writer.book
  
        fmt_header      = workbook.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        fmt_sub_header  = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'font_color': 'black'})
        fmt_currency    = workbook.add_format({'num_format': '_-Rp* #,##0_-;-Rp* #,##0_-;_-Rp* "-"_-;_-@_-', 'border': 1})
        fmt_date        = workbook.add_format({'num_format': 'dd mmm yyyy', 'border': 1, 'align': 'center'})
        fmt_text        = workbook.add_format({'border': 1})
        fmt_text_center = workbook.add_format({'border': 1, 'align': 'center'})
        fmt_title       = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#203764'})
        fmt_bold_total  = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#FFFF00', 'num_format': '_-Rp* #,##0_-'})

        for sheet_name, df_original in all_sheets.items():
            print(f"Processing Data: {sheet_name}...")
        
            sheets_to_process = [(sheet_name, df_original)]
            
            if is_jt_active:
                print(f" -> Membuat Sheet Analisis JT untuk: {sheet_name}...")
                df_jt = df_original.copy()
             
                col_sales_idx = 8 
                nama_sales_col = df_jt.columns[col_sales_idx]
                df_jt = df_jt[~df_jt[nama_sales_col].astype(str).str.contains('FRAUD', case=False, na=False)]
          
                col_umur_idx = 6
                def temporary_clean_umur(val):
                    try: return int(float(str(val).lower().replace(' hari', '').replace(' ', '')))
                    except: return 0
                
                df_jt['Umur_Temp'] = df_jt.iloc[:, col_umur_idx].apply(temporary_clean_umur)
                df_jt = df_jt[df_jt['Umur_Temp'] >= 0]
                df_jt = df_jt.drop(columns=['Umur_Temp'])
         
                sheets_to_process.append((f"{sheet_name} JT", df_jt))

            for current_sheet_name, df in sheets_to_process:
          
                if df.empty:
                    continue
 
                col_tgl_idx = 2 
                tgl_clean = df.iloc[:, col_tgl_idx].astype(str)
                replacements = {
                    ' Jan ': ' Jan ', ' Feb ': ' Feb ', ' Mar ': ' Mar ', ' Apr ': ' Apr ',
                    ' Mei ': ' May ', ' Jun ': ' Jun ', ' Jul ': ' Jul ', ' Agu ': ' Aug ',
                    ' Sep ': ' Sep ', ' Okt ': ' Oct ', ' Nop ': ' Nov ', ' Des ': ' Dec '
                }
                for indo, eng in replacements.items():
                    tgl_clean = tgl_clean.str.replace(indo, eng, case=False, regex=False)
                
                df['Tgl_Sort'] = pd.to_datetime(tgl_clean, errors='coerce', dayfirst=True)

                col_umur_idx = 6 
                def clean_umur(val):
                    try: return int(float(str(val).lower().replace(' hari', '').replace(' ', '')))
                    except: return 0
                df['Umur_Int'] = df.iloc[:, col_umur_idx].apply(clean_umur)

                def get_bucket(days):
                    if days <= 0: return '-30 - 0 Hari'
                    elif days <= 31: return '1-31 Hari'
                    elif days <= 60: return '32-60 Hari'
                    else: return '> 60 Hari'
                df['Kategori'] = df['Umur_Int'].apply(get_bucket)

                col_sisa_idx = 5
              
                df.iloc[:, col_sisa_idx] = pd.to_numeric(df.iloc[:, col_sisa_idx], errors='coerce').fillna(0)

                col_sales_name = df.columns[8] 
                col_sisa_name  = df.columns[5] 
                col_outlet_name = df.columns[7]
                
                pivot = df.pivot_table(index=col_sales_name, columns='Kategori', values=col_sisa_name, aggfunc='sum').fillna(0)
                desired_order = ['-30 - 0 Hari', '1-31 Hari', '32-60 Hari', '> 60 Hari']
                cols_order = [c for c in desired_order if c in pivot.columns]
                pivot = pivot[cols_order]
                pivot['Grand Total'] = pivot.sum(axis=1)

                df_overdue_60 = df[df['Umur_Int'] > 60].copy()
                df_overdue_60[col_sisa_name] = pd.to_numeric(df_overdue_60[col_sisa_name], errors='coerce').fillna(0)
             
                if not df_overdue_60.empty:
                    top_10_outlet = df_overdue_60.groupby([col_outlet_name, col_sales_name])[col_sisa_name].sum().nlargest(10).reset_index()
                else:
                    top_10_outlet = pd.DataFrame(columns=[col_outlet_name, col_sales_name, col_sisa_name])

                worksheet = workbook.add_worksheet(current_sheet_name)
                worksheet.write(0, 0, f"LAPORAN ANALITIK PIUTANG - {current_sheet_name.upper()}", fmt_title)

                col_widths = {}
                def update_w(c, val, is_currency=False):
                    length = (len(f"{val:,.0f}") + 8) if is_currency else (len(str(val)) + 3)
                    if c not in col_widths or length > col_widths[c]:
                        col_widths[c] = length

                start_row = 2
                worksheet.write(start_row, 0, "RINGKASAN PERFORMA SALES", fmt_sub_header)
                worksheet.write(start_row+1, 0, "Nama Penjual", fmt_header)
                update_w(0, "Nama Penjual")
                
                for i, c in enumerate(pivot.columns):
                    worksheet.write(start_row+1, i+1, c, fmt_header)
                    update_w(i+1, c)
                
                for r, (sales, row_data) in enumerate(pivot.iterrows()):
                    worksheet.write(start_row+2+r, 0, sales, fmt_text)
                    update_w(0, sales)
                    for c, col_name in enumerate(pivot.columns):
                        fmt = fmt_bold_total if col_name == 'Grand Total' else fmt_currency
                        worksheet.write(start_row+2+r, c+1, row_data[col_name], fmt)
                        update_w(c+1, row_data[col_name], is_currency=True)
                
                last_summary_row = start_row + 2 + len(pivot)

                top_10_start_col = 7
                worksheet.write(start_row, top_10_start_col, "TOP 10 OUTLET (JATUH TEMPO > 60 HARI)", fmt_sub_header)
                worksheet.write(start_row+1, top_10_start_col, "Nama Outlet/Toko", fmt_header)
                worksheet.write(start_row+1, top_10_start_col+1, "Nama Sales", fmt_header)
                worksheet.write(start_row+1, top_10_start_col+2, "Sisa Piutang", fmt_header)

                update_w(top_10_start_col, "Nama Outlet/Toko")
                update_w(top_10_start_col+1, "Nama Sales")
                update_w(top_10_start_col+2, "Sisa Piutang")

                for r, row in top_10_outlet.iterrows():
                    worksheet.write(start_row+2+r, top_10_start_col, row[col_outlet_name], fmt_text)
                    update_w(top_10_start_col, row[col_outlet_name])
                    
                    worksheet.write(start_row+2+r, top_10_start_col+1, row[col_sales_name], fmt_text)
                    update_w(top_10_start_col+1, row[col_sales_name])
                    
                    worksheet.write(start_row+2+r, top_10_start_col+2, row[col_sisa_name], fmt_currency)
                    update_w(top_10_start_col+2, row[col_sisa_name], is_currency=True)

                chart = workbook.add_chart({'type': 'pie'})
                chart.add_series({
                    'name':       'Porsi Piutang per Sales',
                    'categories': [current_sheet_name, start_row+2, 0, last_summary_row-1, 0], 
                    'values':     [current_sheet_name, start_row+2, len(pivot.columns), last_summary_row-1, len(pivot.columns)],
                    'data_labels': {'value': True, 'percentage': True, 'leader_lines': True, 'position': 'best_fit', 'num_format': '#,##0'},
                })
                chart.set_title({'name': f'Porsi Total Piutang per Sales ({current_sheet_name})'})
                chart.set_legend({'position': 'bottom'})
                chart.set_size({'width': 550, 'height': 400}) 
                worksheet.insert_chart(2, top_10_start_col + 4, chart)

                list_start_row = max(last_summary_row + 3, 18)
                worksheet.write(list_start_row, 0, "RINCIAN DETIL TRANSAKSI", fmt_title)
                list_start_row += 2
                
                headers_show = list(df.columns[:9]) + ['Kategori Umur'] 
                for idx, h in enumerate(headers_show):
                    worksheet.write(list_start_row, idx, h, fmt_header)
                    update_w(idx, h)
                
                current_row = list_start_row + 1
                unique_sales = sorted(df[col_sales_name].dropna().unique())
                
                for sales in unique_sales:
                    df_sales = df[df[col_sales_name] == sales].copy()
                    df_sales = df_sales.sort_values(by=['Tgl_Sort'], ascending=True)
                    total_p = df_sales[col_sisa_name].sum()
                    
                    worksheet.merge_range(current_row, 0, current_row, len(headers_show)-1, 
                                          f"SALES: {sales}  (Total Piutang: Rp {total_p:,.0f})", fmt_sub_header)
                    current_row += 1
                    
                    for _, row in df_sales.iterrows():
                        for c_idx in range(9):
                            val = row.iloc[c_idx]
                            if c_idx == 2: 
                                worksheet.write_datetime(current_row, c_idx, row['Tgl_Sort'], fmt_date) if pd.notna(row['Tgl_Sort']) else worksheet.write(current_row, c_idx, str(val), fmt_text_center)
                                update_w(c_idx, "12 Dec 2024") 
                            elif c_idx == 5 or c_idx == 4: 
                                worksheet.write(current_row, c_idx, val, fmt_currency)
                                update_w(c_idx, val, is_currency=True)
                            else:
                                worksheet.write(current_row, c_idx, str(val), fmt_text)
                                update_w(c_idx, val)
                        
                        fmt_kat = workbook.add_format({'border': 1, 'align': 'center', 'bg_color': '#FFC7CE'}) if row['Umur_Int'] > 60 else fmt_text_center
                        worksheet.write(current_row, 9, row['Kategori'], fmt_kat)
                        update_w(9, row['Kategori'])
                        current_row += 1
                    current_row += 1

                for col_idx, width in col_widths.items():
                    worksheet.set_column(col_idx, col_idx, min(width, 50))

    print(f"Laporan berhasil diperbarui: {output_file}")

input_filename = 'cleandepotemp.xlsx'
output_filename = 'Laporan_Analisis_Piutang.xlsx'
config_filename = 'selatan.conf'

if __name__ == "__main__":
 
    is_jt_active = cek_status_jt(config_filename)
    if is_jt_active:
        print("-> Mode Analisis Jatuh Tempo (JT) AKTIF berdasarkan config.")
        
    analisa_piutang_satu_file_fix(input_filename, output_filename, is_jt_active)
