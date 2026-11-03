import glob
import os
import shutil
import subprocess
import sys


def run_workflow():
    base_dir = os.getcwd()
    dapur_dir = os.path.join(base_dir, "dapur")
    source_file = os.path.join(base_dir, "Piutang.xls")

    required_dapur_files = [
        "1_CleanerAcc.py",
        "2_Split.py",
        "3_Analytics.py",
        "customized.conf",
        "irczn.conf",
        "selatan.conf",
        "utara.conf",
    ]

    if not os.path.exists(source_file):
        print("--> Error: File 'Piutang.xls' tidak ditemukan di folder utama.")
        return

    if not os.path.exists(dapur_dir):
        print("--> Error: Folder 'dapur' tidak ditemukan.")
        return

    missing_files = [
        f
        for f in required_dapur_files
        if not os.path.exists(os.path.join(dapur_dir, f))
    ]
    if missing_files:
        print("--> Error: File berikut tidak ditemukan di dalam folder 'dapur':")
        for f in missing_files:
            print(f"--> - {f}")
        return

    print("--> Membersihkan file sementara di folder dapur...")
    dapur_export_file = os.path.join(dapur_dir, "Piutang.xls")
    if os.path.exists(dapur_export_file):
        os.remove(dapur_export_file)

    for f in glob.glob(os.path.join(dapur_dir, "*temp.xlsx")):
        os.remove(f)

    print("--> Menyalin Piutang.xls ke folder dapur...")
    shutil.copy(source_file, dapur_export_file)

    print("--> \nPilih Area Operasional:")
    print("--> 1. Selatan")
    print("--> 2. Utara")
    print("--> 3. IRC ZN")
    print(
        "--> 4. Customized -> Menggunakan aturan konfigurasi pada"
        " dapur/customized.conf"
    )

    choice = input("--> Masukkan pilihan (1/2/3/4): ").strip()

    pilihan_map = {
        "1": ("_Selatan", "selatan.conf"),
        "2": ("_Utara", "utara.conf"),
        "3": ("_IRC_ZN", "irczn.conf"),
        "4": ("_Customized", "customized.conf"),
    }

    if choice not in pilihan_map:
        print("--> Pilihan tidak valid.")
        return

    suffix_nama, conf_file = pilihan_map[choice]

    commands_to_run = [
        [sys.executable, "1_CleanerAcc.py"],
        [sys.executable, "2_Split.py", conf_file],
        [sys.executable, "3_Analytics.py", conf_file],
    ]

    os.chdir(dapur_dir)

    try:
        for cmd in commands_to_run:
            print("--> --------------------------------------------------")
            print(f"--> Menjalankan script: {' '.join(cmd[1:])}")
            subprocess.run(cmd, check=True)

        original_result = "Laporan_Analisis_Piutang.xlsx"
        final_result = f"Laporan_Analisis_Piutang{suffix_nama}.xlsx"

        if os.path.exists(original_result):
            print(
                f"--> \nMenyalin hasil ke folder utama sebagai: {final_result}..."
            )
            shutil.copy(original_result, os.path.join(base_dir, final_result))
        else:
            print(f"--> \nWarning: {original_result} tidak ditemukan.")

    except subprocess.CalledProcessError as e:
        print(f"--> \nError: Terjadi kesalahan saat menjalankan {e.cmd}.")
    except Exception as e:
        print(f"--> \nError tidak terduga: {e}")
    finally:
        if os.path.exists("Piutang.xls"):
            os.remove("Piutang.xls")
        if os.path.exists("Laporan_Analisis_Piutang.xlsx"):
            os.remove("Laporan_Analisis_Piutang.xlsx")
        for t in glob.glob("*temp.xlsx"):
            os.remove(t)
        os.chdir(base_dir)


if __name__ == "__main__":
    try:
        run_workflow()
    except Exception as e:
        print(f"--> System Error: {e}")

    input("--> \nTekan Enter untuk keluar...")