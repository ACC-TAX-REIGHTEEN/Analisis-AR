import re
import numpy as np
import pandas as pd


def deteksi_dan_bersihkan_umur_jt(series_umur):
    sampel_valid = series_umur.dropna().astype(str).str.strip()
    sampel_valid = sampel_valid[sampel_valid != ""]

    ada_teks_hari = sampel_valid.str.contains(
        "hari", case=False, na=False
    ).any()

    if ada_teks_hari:
        mode_data = "MODE BARU (Mengandung teks 'hari')"
        print(f"--> Deteksi Umur JT: {mode_data}")

        def clean_umur_baru(val):
            if pd.isna(val):
                return ""
            return str(val).strip()

        return series_umur.apply(clean_umur_baru), "BARU"
    else:
        mode_data = "MODE LAMA (Tanpa teks 'hari' - Format Angka)"
        print(
            f"--> Deteksi Umur JT: {mode_data}. Membersihkan desimal (,00 / .0)..."
        )

        def clean_umur_lama(val):
            if pd.isna(val):
                return ""
            s = str(val).strip()
            if not s:
                return ""
            if s.endswith(",00"):
                s = s[:-3]
            elif s.endswith(".00"):
                s = s[:-3]
            elif s.endswith(".0"):
                s = s[:-2]
            try:
                angka = float(s.replace(".", "").replace(",", "."))
                return str(int(angka))
            except ValueError:
                return s

        return series_umur.apply(clean_umur_lama), "LAMA"


def cari_kolom_kode_pelanggan(df, nama_kolom_faktur):
    if "Kode Pelanggan" in df.columns:
        return "Kode Pelanggan"

    idx_faktur = df.columns.get_loc(nama_kolom_faktur)
    for i in range(idx_faktur - 1, -1, -1):
        col_candidate = df.columns[i]
        sampel = df[col_candidate].dropna().astype(str).str.strip()
        if not sampel.empty and not sampel.str.upper().eq("IDR").all():
            return col_candidate

    return df.columns[max(0, idx_faktur - 1)]


def clean_data_autofit(input_file, output_file):
    print(f"--> Sedang memproses file: {input_file}...")

    try:
        df = pd.read_csv(input_file, header=3)
    except Exception:
        df = pd.read_excel(input_file, header=3)

    df.columns = [str(col).strip() for col in df.columns]

    required_named_columns = [
        "No. Faktur",
        "Tgl Faktur",
        "Jatuh Tempo",
        "Nilai Faktur",
        "Sisa Piutang",
        "Umur JT",
        "Nama Pelanggan",
        "Nama Penjual",
        "Nama Kontak",
    ]

    missing_cols = [c for c in required_named_columns if c not in df.columns]
    if missing_cols:
        raise ValueError(
            f"Kolom berikut tidak ditemukan di file sumber: {missing_cols}"
        )

    col_kode_asli = cari_kolom_kode_pelanggan(df, "No. Faktur")

    selected_cols = [col_kode_asli] + required_named_columns
    df_clean = df[selected_cols].copy()

    df_clean.rename(columns={col_kode_asli: "Kode Pelanggan"}, inplace=True)

    df_clean["Kode Pelanggan"] = df_clean["Kode Pelanggan"].ffill()
    df_clean = df_clean.dropna(subset=["No. Faktur"]).copy()

    def format_clean(val):
        if pd.isna(val):
            return ""
        s = str(val).strip()
        if s.endswith(".0"):
            return s[:-2]
        if s.endswith(",00"):
            return s[:-3]
        return s

    cols_to_clean = [
        "Kode Pelanggan",
        "No. Faktur",
        "Nilai Faktur",
        "Sisa Piutang",
    ]
    for col in cols_to_clean:
        if col in df_clean.columns:
            df_clean[col] = df_clean[col].apply(format_clean)

    df_clean["Umur JT"], _ = deteksi_dan_bersihkan_umur_jt(df_clean["Umur JT"])

    df_clean.reset_index(drop=True, inplace=True)

    try:
        with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
            df_clean.to_excel(writer, index=False, sheet_name="Data Bersih")

            worksheet = writer.sheets["Data Bersih"]
            for i, col in enumerate(df_clean.columns):
                panjang_data = df_clean[col].dropna().astype(str).map(len)
                max_data_len = (
                    panjang_data.max() if not panjang_data.empty else 0
                )
                max_len = max(int(max_data_len), len(str(col))) + 2
                worksheet.set_column(i, i, max_len)

        print(f"--> SUKSES! File tersimpan rapi di: {output_file}")

    except Exception as e:
        print(f"--> Error saat menyimpan file: {e}")

    return df_clean


input_filename = "Piutang.xls"
output_filename = "Piutang_cleantemp.xlsx"

if __name__ == "__main__":
    clean_data_autofit(input_filename, output_filename)