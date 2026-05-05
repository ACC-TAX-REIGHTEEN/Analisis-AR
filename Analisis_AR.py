import os
import shutil
import subprocess
import glob
import sys

def run_workflow():
    base_dir = os.getcwd()
    dapur_dir = os.path.join(base_dir, 'dapur')
    source_file = os.path.join(base_dir, 'ExportFile.xls')
    
    required_dapur_files = [
        '1_clean_all_piutang_withNOPEL.py',
        '2_separate_depo_customized.py',
        '2_separate_depo_IRCZN.py',
        '2_separate_depo_selatan.py',
        '2_separate_depo_utara.py',
        '3_analytics_salesman.py',
        'customized.conf',
        'irczn.conf',
        'selatan.conf',
        'utara.conf'
    ]

    if not os.path.exists(source_file):
        print("Error: File 'ExportFile.xls' tidak ditemukan di folder ini.")
        return

    if not os.path.exists(dapur_dir):
        print("Error: Folder 'dapur' tidak ditemukan.")
        return

    missing_files = []
    for f in required_dapur_files:
        if not os.path.exists(os.path.join(dapur_dir, f)):
            missing_files.append(f)
    
    if missing_files:
        print("Error: File berikut tidak ditemukan di dalam folder 'dapur':")
        for f in missing_files:
            print(f"- {f}")
        return

    print("Membersihkan folder dapur...")
    dapur_export_file = os.path.join(dapur_dir, 'ExportFile.xls')
    if os.path.exists(dapur_export_file):
        os.remove(dapur_export_file)

    temp_files = glob.glob(os.path.join(dapur_dir, '*temp.xlsx'))
    for f in temp_files:
        os.remove(f)

    print("Menyalin ExportFile.xls ke dapur...")
    shutil.copy(source_file, dapur_export_file)

    print("\nPilih Area Operasional:")
    print("1. Selatan")
    print("2. Utara")
    print("3. IRC ZN")
    print("4. Customized. -> Pastikan sudah konfigurasikan customized.conf di folder Dapur")
    
    choice = input("Masukkan pilihan (1/2/3/4): ")
    
    scripts_to_run = []
    suffix_nama = ""
    
    if choice == '1':
        suffix_nama = "_Selatan"
        scripts_to_run = [
            '1_clean_all_piutang_withNOPEL.py',
            '2_separate_depo_selatan.py',
            '3_analytics_salesman.py'
        ]
    elif choice == '2':
        suffix_nama = "_Utara"
        scripts_to_run = [
            '1_clean_all_piutang_withNOPEL.py',
            '2_separate_depo_utara.py',
            '3_analytics_salesman.py'
        ]
    elif choice == '3':
        suffix_nama = "_IRC_ZN"
        scripts_to_run = [
            '1_clean_all_piutang_withNOPEL.py',
            '2_separate_depo_IRCZN.py',
            '3_analytics_salesman.py'
        ]
    elif choice == '4':
        suffix_nama = "_Customized"
        scripts_to_run = [
            '1_clean_all_piutang_withNOPEL.py',
            '2_separate_depo_customized.py',
            '3_analytics_salesman.py'
        ]    
    else:
        print("Pilihan tidak valid.")
        return

    os.chdir(dapur_dir)
    
    try:
        for script in scripts_to_run:
            print(f"--------------------------------------------------")
            print(f"Menjalankan: {script}")
            subprocess.run([sys.executable, script], check=True)
            
        original_result_filename = 'Laporan_Analisis_Piutang.xlsx'
        final_result_filename = f'Laporan_Analisis_Piutang{suffix_nama}.xlsx'

        if os.path.exists(original_result_filename):
            print(f"\nMenyalin hasil ke folder utama sebagai: {final_result_filename}...")
            shutil.copy(original_result_filename, os.path.join(base_dir, final_result_filename))
        else:
            print(f"\nWarning: {original_result_filename} tidak ditemukan setelah proses selesai.")

    except subprocess.CalledProcessError as e:
        print(f"\nError: Terjadi kesalahan saat menjalankan script {e.cmd}.")
    except Exception as e:
        print(f"\nError tidak terduga: {e}")
    finally:
        if os.path.exists('ExportFile.xls'):
            os.remove('ExportFile.xls')
            
        if os.path.exists('Laporan_Analisis_Piutang.xlsx'):
            os.remove('Laporan_Analisis_Piutang.xlsx')    
        
        temps = glob.glob('*temp.xlsx')
        for t in temps:
            os.remove(t)
            
        os.chdir(base_dir)

if __name__ == "__main__":
    try:
        run_workflow()
    except Exception as e:
        print(f"System Error: {e}")
    
    input("\nTekan Enter untuk keluar...")
