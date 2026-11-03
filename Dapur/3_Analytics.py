import configparser
from datetime import datetime
import os
import sys
import pandas as pd
import xlsxwriter


def cek_status_jt(conf_filename):
    if not os.path.exists(conf_filename):
        return False

    try:
        config = configparser.ConfigParser()
        config.read(conf_filename)

        if config.has_section("JT"):
            return config.getboolean("JT", "Aktif", fallback=False)

    except Exception as e:
        print(
            f"--> Peringatan: Gagal membaca status JT dari {conf_filename} ->"
            f" {e}"
        )

    return False


def baca_gen_fraud(conf_filename):
    if not os.path.exists(conf_filename):
        return ""

    try:
        config = configparser.ConfigParser()
        config.read(conf_filename)

        if config.has_section("GENERAL"):
            nilai = config.get("GENERAL", "gen_fraud", fallback="").strip()
            return nilai

    except Exception as e:
        print(
            f"--> Peringatan: Gagal membaca gen_fraud dari {conf_filename} ->"
            f" {e}"
        )

    return ""


def deteksi_mode_umur(series_umur):
    sampel = series_umur.dropna().astype(str).str.strip()
    ada_hari = sampel.str.contains("hari", case=False, na=False).any()
    return "BARU" if ada_hari else "LAMA"


def parse_umur_ke_int(val):
    if pd.isna(val):
        return 0
    s = str(val).lower().replace("hari", "").strip()
    if not s:
        return 0
    if s.endswith(",00") or s.endswith(".00"):
        s = s[:-3]
    elif s.endswith(".0"):
        s = s[:-2]
    s = s.replace(" ", "").replace(",", ".")
    try:
        return int(float(s))
    except ValueError:
        return 0


def parse_nominal_rupiah(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s:
        return 0.0
    if s.endswith(",00") or s.endswith(".00"):
        s = s[:-3]
    elif s.endswith(".0"):
        s = s[:-2]
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def paksa_jadi_tanggal(val):
    if pd.isna(val):
        return pd.NaT

    val_str = str(val).strip()

    bulan_indo_eng = {
        "jan": "Jan",
        "feb": "Feb",
        "mar": "Mar",
        "apr": "Apr",
        "mei": "May",
        "jun": "Jun",
        "jul": "Jul",
        "agu": "Aug",
        "ags": "Aug",
        "sep": "Sep",
        "okt": "Oct",
        "nop": "Nov",
        "nov": "Nov",
        "des": "Dec",
    }

    val_lower = val_str.lower()
    for indo, eng in bulan_indo_eng.items():
        if indo in val_lower:
            val_str = val_lower.replace(indo, eng)
            break

    for fmt in ("%d/%m/%Y", "%d %b %Y", "%Y-%m-%d %H:%M:%S", "%d-%m-%Y"):
        try:
            return pd.to_datetime(val_str, format=fmt)
        except ValueError:
            continue

    return pd.to_datetime(val_str, errors="coerce", dayfirst=True)


def analisa_piutang_satu_file_fix(
    input_file, output_file, is_jt_active, gen_fraud_pattern
):
    print(f"--> Membaca file sumber: {input_file}...")

    try:
        all_sheets = pd.read_excel(input_file, sheet_name=None)
    except Exception as e:
        print(f"--> Error membaca file: {e}")
        return

    COL_KODE = "Kode Pelanggan"
    COL_FAKTUR = "No. Faktur"
    COL_TGL = "Tgl Faktur"
    COL_JT = "Jatuh Tempo"
    COL_NILAI = "Nilai Faktur"
    COL_SISA = "Sisa Piutang"
    COL_UMUR = "Umur JT"
    COL_OUTLET = "Nama Pelanggan"
    COL_SALES = "Nama Penjual"
    COL_KONTAK = "Nama Kontak"

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        workbook = writer.book

        fmt_header = workbook.add_format({
            "bold": True,
            "bg_color": "#4472C4",
            "font_color": "white",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
        })
        fmt_sub_header = workbook.add_format({
            "bold": True,
            "bg_color": "#D9E1F2",
            "border": 1,
            "font_color": "black",
        })
        fmt_currency = workbook.add_format({
            "num_format": '_-Rp* #,##0_-;-Rp* #,##0_-;_-Rp* "-"_-;_-@_-',
            "border": 1,
        })
        fmt_date = workbook.add_format({
            "num_format": "dd mmm yyyy",
            "border": 1,
            "align": "center",
        })
        fmt_text = workbook.add_format({"border": 1})
        fmt_text_center = workbook.add_format({"border": 1, "align": "center"})
        fmt_title = workbook.add_format({
            "bold": True,
            "font_size": 14,
            "font_color": "#203764",
        })
        fmt_bold_total = workbook.add_format({
            "bold": True,
            "border": 1,
            "bg_color": "#FFFF00",
            "num_format": "_-Rp* #,##0_-",
        })
        fmt_overdue = workbook.add_format({
            "border": 1,
            "align": "center",
            "bg_color": "#FFC7CE",
        })

        for sheet_name, df_original in all_sheets.items():
            print(f"--> Processing Data: {sheet_name}...")
            df_original.columns = [
                str(c).strip() for c in df_original.columns
            ]

            required_cols = [
                COL_TGL,
                COL_SISA,
                COL_UMUR,
                COL_OUTLET,
                COL_SALES,
            ]
            missing = [c for c in required_cols if c not in df_original.columns]
            if missing:
                print(
                    f"--> Peringatan: Sheet '{sheet_name}' dilewati karena"
                    f" kolom berikut tidak ditemukan: {missing}"
                )
                continue

            if gen_fraud_pattern:
                baris_sebelum = len(df_original)
                mask_fraud = (
                    df_original[COL_SALES]
                    .astype(str)
                    .str.contains(
                        gen_fraud_pattern, case=False, na=False, regex=True
                    )
                )
                df_original = df_original[~mask_fraud].copy()
                baris_terhapus = baris_sebelum - len(df_original)
                if baris_terhapus > 0:
                    print(
                        f"--> -> [{sheet_name}] Memfilter {baris_terhapus}"
                        " baris Nama Penjual yang memuat pola:"
                        f" '{gen_fraud_pattern}'"
                    )
            else:
                print(
                    f"--> -> [{sheet_name}] gen_fraud kosong, memproses seluruh"
                    " data Nama Penjual."
                )

            mode_terdeteksi = deteksi_mode_umur(df_original[COL_UMUR])
            print(f"--> -> Mode Data pada {sheet_name}: Mode {mode_terdeteksi}")

            sheets_to_process = [(sheet_name, df_original)]

            if is_jt_active:
                print(
                    f"--> -> Membuat Sheet Analisis JT untuk: {sheet_name}..."
                )
                df_jt = df_original.copy()

                df_jt["Umur_Temp"] = df_jt[COL_UMUR].apply(parse_umur_ke_int)
                df_jt = df_jt[df_jt["Umur_Temp"] >= 0].copy()
                df_jt.drop(columns=["Umur_Temp"], inplace=True)

                sheets_to_process.append((f"{sheet_name} JT", df_jt))

            for current_sheet_name, df in sheets_to_process:
                if df.empty:
                    continue

                df = df.copy()

                df["Tgl_Sort"] = df[COL_TGL].apply(paksa_jadi_tanggal)
                df["Umur_Int"] = df[COL_UMUR].apply(parse_umur_ke_int)
                df[COL_SISA] = df[COL_SISA].apply(parse_nominal_rupiah)
                if COL_NILAI in df.columns:
                    df[COL_NILAI] = df[COL_NILAI].apply(parse_nominal_rupiah)

                if mode_terdeteksi == "LAMA":
                    df[COL_UMUR] = df["Umur_Int"].astype(str)

                def get_bucket(days):
                    if days <= 0:
                        return "-30 - 0 Hari"
                    elif days <= 31:
                        return "1-31 Hari"
                    elif days <= 60:
                        return "32-60 Hari"
                    else:
                        return "> 60 Hari"

                df["Kategori"] = df["Umur_Int"].apply(get_bucket)

                pivot = df.pivot_table(
                    index=COL_SALES,
                    columns="Kategori",
                    values=COL_SISA,
                    aggfunc="sum",
                ).fillna(0)
                desired_order = [
                    "-30 - 0 Hari",
                    "1-31 Hari",
                    "32-60 Hari",
                    "> 60 Hari",
                ]
                cols_order = [c for c in desired_order if c in pivot.columns]
                pivot = pivot[cols_order]
                pivot["Grand Total"] = pivot.sum(axis=1)

                df_overdue_60 = df[df["Umur_Int"] > 60].copy()
                if not df_overdue_60.empty:
                    top_10_outlet = (
                        df_overdue_60.groupby([COL_OUTLET, COL_SALES])[
                            COL_SISA
                        ]
                        .sum()
                        .nlargest(10)
                        .reset_index()
                    )
                else:
                    top_10_outlet = pd.DataFrame(
                        columns=[COL_OUTLET, COL_SALES, COL_SISA]
                    )

                worksheet = workbook.add_worksheet(current_sheet_name)
                worksheet.write(
                    0,
                    0,
                    f"LAPORAN ANALITIK PIUTANG - {current_sheet_name.upper()}",
                    fmt_title,
                )

                col_widths = {}

                def update_w(c, val, is_currency=False):
                    length = (
                        (len(f"{val:,.0f}") + 8)
                        if is_currency
                        else (len(str(val)) + 3)
                    )
                    if c not in col_widths or length > col_widths[c]:
                        col_widths[c] = length

                start_row = 2
                worksheet.write(
                    start_row, 0, "RINGKASAN PERFORMA SALES", fmt_sub_header
                )
                worksheet.write(start_row + 1, 0, COL_SALES, fmt_header)
                update_w(0, COL_SALES)

                for i, c in enumerate(pivot.columns):
                    worksheet.write(start_row + 1, i + 1, c, fmt_header)
                    update_w(i + 1, c)

                for r, (sales, row_data) in enumerate(pivot.iterrows()):
                    worksheet.write(start_row + 2 + r, 0, sales, fmt_text)
                    update_w(0, sales)
                    for c, col_name in enumerate(pivot.columns):
                        fmt = (
                            fmt_bold_total
                            if col_name == "Grand Total"
                            else fmt_currency
                        )
                        worksheet.write(
                            start_row + 2 + r, c + 1, row_data[col_name], fmt
                        )
                        update_w(c + 1, row_data[col_name], is_currency=True)

                last_summary_row = start_row + 2 + len(pivot)

                top_10_start_col = 7
                worksheet.write(
                    start_row,
                    top_10_start_col,
                    "TOP 10 OUTLET (JATUH TEMPO > 60 HARI)",
                    fmt_sub_header,
                )
                worksheet.write(
                    start_row + 1,
                    top_10_start_col,
                    "Nama Outlet/Toko",
                    fmt_header,
                )
                worksheet.write(
                    start_row + 1,
                    top_10_start_col + 1,
                    "Nama Sales",
                    fmt_header,
                )
                worksheet.write(
                    start_row + 1, top_10_start_col + 2, COL_SISA, fmt_header
                )

                update_w(top_10_start_col, "Nama Outlet/Toko")
                update_w(top_10_start_col + 1, "Nama Sales")
                update_w(top_10_start_col + 2, COL_SISA)

                for r, row in top_10_outlet.iterrows():
                    worksheet.write(
                        start_row + 2 + r,
                        top_10_start_col,
                        row[COL_OUTLET],
                        fmt_text,
                    )
                    update_w(top_10_start_col, row[COL_OUTLET])

                    worksheet.write(
                        start_row + 2 + r,
                        top_10_start_col + 1,
                        row[COL_SALES],
                        fmt_text,
                    )
                    update_w(top_10_start_col + 1, row[COL_SALES])

                    worksheet.write(
                        start_row + 2 + r,
                        top_10_start_col + 2,
                        row[COL_SISA],
                        fmt_currency,
                    )
                    update_w(
                        top_10_start_col + 2, row[COL_SISA], is_currency=True
                    )

                chart = workbook.add_chart({"type": "pie"})
                chart.add_series({
                    "name": "Porsi Piutang per Sales",
                    "categories": [
                        current_sheet_name,
                        start_row + 2,
                        0,
                        last_summary_row - 1,
                        0,
                    ],
                    "values": [
                        current_sheet_name,
                        start_row + 2,
                        len(pivot.columns),
                        last_summary_row - 1,
                        len(pivot.columns),
                    ],
                    "data_labels": {
                        "value": True,
                        "percentage": True,
                        "separator": "\n",
                        "leader_lines": True,
                        "position": "best_fit",
                        # "num_format": "#,##0",
                    },
                })
                chart.set_title({
                    "name": (
                        f"Porsi Total Piutang per Sales ({current_sheet_name})"
                    )
                })
                chart.set_legend({"position": "bottom"})
                chart.set_size({"width": 550, "height": 400})
                worksheet.insert_chart(2, top_10_start_col + 4, chart)

                list_start_row = max(last_summary_row + 3, 18)
                worksheet.write(
                    list_start_row, 0, "RINCIAN DETIL TRANSAKSI", fmt_title
                )
                list_start_row += 2

                urutan_kolom_rincian = [
                    COL_KODE,
                    COL_FAKTUR,
                    COL_TGL,
                    COL_JT,
                    COL_NILAI,
                    COL_SISA,
                    COL_UMUR,
                    COL_OUTLET,
                    COL_SALES,
                ]
                kolom_rincian_aktif = [
                    c for c in urutan_kolom_rincian if c in df.columns
                ]
                headers_show = kolom_rincian_aktif + ["Kategori Umur"]

                for idx, h in enumerate(headers_show):
                    worksheet.write(list_start_row, idx, h, fmt_header)
                    update_w(idx, h)

                current_row = list_start_row + 1
                unique_sales = sorted(df[COL_SALES].dropna().unique())

                for sales in unique_sales:
                    df_sales = df[df[COL_SALES] == sales].copy()
                    df_sales = df_sales.sort_values(
                        by=["Tgl_Sort"], ascending=True
                    )
                    total_p = df_sales[COL_SISA].sum()

                    worksheet.merge_range(
                        current_row,
                        0,
                        current_row,
                        len(headers_show) - 1,
                        f"SALES: {sales}  (Total Piutang: Rp {total_p:,.0f})",
                        fmt_sub_header,
                    )
                    current_row += 1

                    for _, row in df_sales.iterrows():
                        for c_idx, col_name in enumerate(kolom_rincian_aktif):
                            val = row[col_name]

                            if col_name == COL_TGL:
                                if pd.notna(row["Tgl_Sort"]):
                                    worksheet.write_datetime(
                                        current_row,
                                        c_idx,
                                        row["Tgl_Sort"],
                                        fmt_date,
                                    )
                                else:
                                    worksheet.write(
                                        current_row,
                                        c_idx,
                                        str(val),
                                        fmt_text_center,
                                    )
                                update_w(c_idx, "12 Dec 2026")
                            elif col_name in (COL_NILAI, COL_SISA):
                                worksheet.write(
                                    current_row, c_idx, val, fmt_currency
                                )
                                update_w(c_idx, val, is_currency=True)
                            elif col_name == COL_UMUR:
                                worksheet.write(
                                    current_row, c_idx, str(val), fmt_text_center
                                )
                                update_w(c_idx, val)
                            else:
                                worksheet.write(
                                    current_row, c_idx, str(val), fmt_text
                                )
                                update_w(c_idx, val)

                        fmt_kat = (
                            fmt_overdue
                            if row["Umur_Int"] > 60
                            else fmt_text_center
                        )
                        worksheet.write(
                            current_row,
                            len(kolom_rincian_aktif),
                            row["Kategori"],
                            fmt_kat,
                        )
                        update_w(len(kolom_rincian_aktif), row["Kategori"])
                        current_row += 1
                    current_row += 1

                for col_idx, width in col_widths.items():
                    worksheet.set_column(col_idx, col_idx, min(width, 50))

    print(f"--> Laporan berhasil diperbarui: {output_file}")


input_filename = "cleandepotemp.xlsx"
output_filename = "Laporan_Analisis_Piutang.xlsx"
config_filename = "customized.conf"

if __name__ == "__main__":
    if len(sys.argv) > 1:
        config_filename = sys.argv[1]
        print(f"--> Menjalankan analisis menggunakan config: {config_filename}")
    else:
        print(
            f"--> Tidak ada input file. Menggunakan default: {config_filename}"
        )

    is_jt_active = cek_status_jt(config_filename)
    if is_jt_active:
        print(
            f"--> Mode Analisis Jatuh Tempo (JT) AKTIF untuk {config_filename}."
        )
    else:
        print(
            f"--> Mode Analisis Jatuh Tempo (JT) MATI untuk {config_filename}."
        )

    gen_fraud_pattern = baca_gen_fraud(config_filename)
    if gen_fraud_pattern:
        print(
            f"--> Aturan gen_fraud ditemukan: '{gen_fraud_pattern}' (Akan"
            " memfilter Nama Penjual)"
        )
    else:
        print(
            "--> Aturan gen_fraud kosong/tidak ada. Seluruh data Nama Penjual"
            " akan diproses."
        )

    analisa_piutang_satu_file_fix(
        input_filename, output_filename, is_jt_active, gen_fraud_pattern
    )