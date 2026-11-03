import configparser
import os
import sys
import pandas as pd


def load_config(config_file):
    if not os.path.exists(config_file):
        print(f"--> Error: File konfigurasi '{config_file}' tidak ditemukan.")
        sys.exit(1)

    config = configparser.ConfigParser()
    config.optionxform = str
    config.read(config_file)

    section_name = None
    if "PREFIX_TO_SHEET" in config:
        section_name = "PREFIX_TO_SHEET"
    elif "FILTER_PREFIX" in config:
        section_name = "FILTER_PREFIX"
    else:
        print(
            "--> Error: Bagian [PREFIX_TO_SHEET] atau [FILTER_PREFIX] tidak"
            f" ditemukan di {config_file}."
        )
        sys.exit(1)

    grouped_filters = {}
    for prefix, sheet_name in config[section_name].items():
        clean_prefix = prefix.strip()
        clean_sheet = sheet_name.strip().replace('"', "").replace("'", "")

        if clean_sheet not in grouped_filters:
            grouped_filters[clean_sheet] = []
        grouped_filters[clean_sheet].append(clean_prefix)

    prefix_map = {k: tuple(v) for k, v in grouped_filters.items()}

    product_settings = {"col_name": None, "regex_pattern": None}
    if "PRODUCT_FILTERS" in config:
        product_settings["col_name"] = (
            config["PRODUCT_FILTERS"].get("COLUMN_NAME", "").strip()
        )
        raw_keywords = config["PRODUCT_FILTERS"].get("KEYWORDS", "")
        keyword_list = [
            k.strip() for k in raw_keywords.split(",") if k.strip()
        ]
        if keyword_list:
            product_settings["regex_pattern"] = "|".join(keyword_list)

    return prefix_map, product_settings


def pastikan_umur_jt_bersih(series_umur):
    sampel_valid = series_umur.dropna().astype(str).str.strip()
    ada_teks_hari = sampel_valid.str.contains(
        "hari", case=False, na=False
    ).any()

    if ada_teks_hari:
        return series_umur.astype(str).str.strip()
    else:

        def bersihkan_angka(val):
            if pd.isna(val):
                return ""
            s = str(val).strip()
            if s.endswith(",00") or s.endswith(".00"):
                s = s[:-3]
            elif s.endswith(".0"):
                s = s[:-2]
            try:
                return str(int(float(s.replace(".", "").replace(",", "."))))
            except ValueError:
                return s

        return series_umur.apply(bersihkan_angka)


def split_depo_final(input_file, output_file, config_file):
    print(f"--> Memisahkan data menggunakan konfigurasi: '{config_file}'")
    prefix_map, product_filters = load_config(config_file)

    try:
        df = pd.read_csv(input_file)
    except Exception:
        try:
            df = pd.read_excel(input_file)
        except Exception as e:
            print(f"--> Gagal membaca file input: {e}")
            return

    df.columns = [str(c).strip() for c in df.columns]

    if "Umur JT" in df.columns:
        df["Umur JT"] = pastikan_umur_jt_bersih(df["Umur JT"])

    search_col = product_filters["col_name"]
    search_regex = product_filters["regex_pattern"]

    if search_col and search_regex:
        if search_col in df.columns:
            filter_mask = (
                df[search_col]
                .astype(str)
                .str.contains(search_regex, na=False, regex=True)
            )
            df = df[filter_mask].copy()
        else:
            print(
                "--> Peringatan: Kolom filter"
                f" '{search_col}' tidak ditemukan di data."
            )

    target_col = "Nama Penjual"
    if target_col not in df.columns:
        print(f"--> Error: Kolom '{target_col}' tidak ditemukan.")
        return

    try:
        with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
            total_saved = 0

            for sheet_name, prefixes in prefix_map.items():
                mask = (
                    df[target_col]
                    .astype(str)
                    .str.strip()
                    .str.startswith(prefixes)
                )
                df_filtered = df[mask].copy()

                if not df_filtered.empty:
                    df_filtered.to_excel(
                        writer, sheet_name=sheet_name, index=False
                    )
                    total_saved += 1
                    print(
                        f"--> Sheet '{sheet_name}': {len(df_filtered)} baris"
                        f" (Prefix: {list(prefixes)})"
                    )

                    worksheet = writer.sheets[sheet_name]
                    for i, col in enumerate(df_filtered.columns):
                        panjang_data = (
                            df_filtered[col].dropna().astype(str).map(len)
                        )
                        max_data_len = (
                            panjang_data.max() if not panjang_data.empty else 0
                        )
                        max_len = max(int(max_data_len), len(str(col))) + 2
                        worksheet.set_column(i, i, max_len)

            if total_saved > 0:
                print(f"--> SUKSES! Data terbagi rapi di: {output_file}\n")
            else:
                print(
                    "--> Peringatan: Tidak ada data yang cocok dengan"
                    " konfigurasi ini.\n"
                )

    except Exception as e:
        print(f"--> Terjadi kesalahan saat menyimpan file: {e}")


if __name__ == "__main__":
    input_filename = "Piutang_cleantemp.xlsx"
    output_filename = "cleandepotemp.xlsx"

    if len(sys.argv) > 1:
        config_filename = sys.argv[1]
    else:
        config_filename = "customized.conf"

    split_depo_final(input_filename, output_filename, config_filename)