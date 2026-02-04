#!/usr/bin/env python3
"""
Sistem Absensi Karyawan - Versi Command Line
Untuk sistem tanpa GUI/Tkinter
"""

import os
from datetime import datetime
from attendance_processor import AttendanceProcessor


def clear_screen():
    """Clear terminal screen"""
    os.system('clear' if os.name == 'posix' else 'cls')


def print_header():
    """Print header aplikasi"""
    clear_screen()
    print("=" * 70)
    print("📊 SISTEM ABSENSI KARYAWAN - COMMAND LINE VERSION".center(70))
    print("=" * 70)
    print()


def print_menu(title, options):
    """Print menu dengan pilihan"""
    print(f"\n{title}")
    print("-" * 50)
    for i, option in enumerate(options, 1):
        print(f"{i}. {option}")
    print()


def get_choice(max_option):
    """Dapatkan pilihan user"""
    while True:
        try:
            choice = int(input(f"Pilih (1-{max_option}): "))
            if 1 <= choice <= max_option:
                return choice
            else:
                print(f"❌ Pilihan harus antara 1-{max_option}")
        except ValueError:
            print("❌ Input harus berupa angka")


def select_file():
    """Pilih file CSV"""
    print("\n📁 PILIH FILE DATA ABSENSI")
    print("-" * 50)
    
    # Cari file CSV di direktori saat ini
    csv_files = [f for f in os.listdir('.') if f.endswith('.csv')]
    
    # Tambahkan file dari folder contoh
    if os.path.exists('contoh'):
        contoh_files = [f'contoh/{f}' for f in os.listdir('contoh') if f.endswith('.csv')]
        csv_files.extend(contoh_files)
    
    if not csv_files:
        print("❌ Tidak ada file CSV ditemukan!")
        return None
    
    print("\nFile CSV yang tersedia:")
    for i, file in enumerate(csv_files, 1):
        print(f"{i}. {file}")
    
    print(f"{len(csv_files) + 1}. Input manual path file")
    
    choice = get_choice(len(csv_files) + 1)
    
    if choice == len(csv_files) + 1:
        file_path = input("\nMasukkan path file CSV: ").strip()
        return file_path if os.path.exists(file_path) else None
    else:
        return csv_files[choice - 1]


def select_employee(employee_list):
    """Pilih karyawan"""
    print("\n👥 PILIH KARYAWAN")
    print("-" * 50)
    
    # Tampilkan 20 karyawan per halaman
    page_size = 20
    total_pages = (len(employee_list) + page_size - 1) // page_size
    current_page = 0
    
    while True:
        start_idx = current_page * page_size
        end_idx = min(start_idx + page_size, len(employee_list))
        
        print(f"\nHalaman {current_page + 1}/{total_pages}")
        print("-" * 50)
        
        for i in range(start_idx, end_idx):
            print(f"{i + 1}. {employee_list[i]}")
        
        print("\nNavigasi:")
        if current_page > 0:
            print("P. Halaman sebelumnya")
        if current_page < total_pages - 1:
            print("N. Halaman selanjutnya")
        print("S. Cari karyawan")
        
        choice = input("\nPilih nomor / P / N / S: ").strip().upper()
        
        if choice == 'P' and current_page > 0:
            current_page -= 1
        elif choice == 'N' and current_page < total_pages - 1:
            current_page += 1
        elif choice == 'S':
            search = input("Masukkan nama (sebagian): ").strip()
            filtered = [e for e in employee_list if search.lower() in e.lower()]
            if filtered:
                print(f"\nDitemukan {len(filtered)} karyawan:")
                for i, emp in enumerate(filtered[:10], 1):
                    print(f"{i}. {emp}")
                try:
                    idx = int(input("\nPilih nomor: ")) - 1
                    if 0 <= idx < len(filtered):
                        return filtered[idx]
                except ValueError:
                    pass
            else:
                print("❌ Tidak ditemukan")
                input("Tekan Enter untuk melanjutkan...")
        else:
            try:
                idx = int(choice) - 1
                if start_idx <= idx < end_idx:
                    return employee_list[idx]
            except ValueError:
                pass


def select_period():
    """Pilih periode (bulan dan tahun)"""
    print("\n📅 PILIH PERIODE LAPORAN")
    print("-" * 50)
    
    months = {
        1: 'Januari', 2: 'Februari', 3: 'Maret', 4: 'April',
        5: 'Mei', 6: 'Juni', 7: 'Juli', 8: 'Agustus',
        9: 'September', 10: 'Oktober', 11: 'November', 12: 'Desember'
    }
    
    # Pilih bulan
    print("\nBulan:")
    for i in range(1, 13):
        print(f"{i:2d}. {months[i]}")
    
    month = get_choice(12)
    
    # Pilih tahun
    print("\nTahun:")
    current_year = datetime.now().year
    years = list(range(current_year - 2, current_year + 5))
    
    for i, year in enumerate(years, 1):
        print(f"{i}. {year}")
    
    year_choice = get_choice(len(years))
    year = years[year_choice - 1]
    
    return month, year, months[month]


def main():
    """Main function"""
    processor = AttendanceProcessor()
    
    while True:
        print_header()
        print("Selamat datang di Sistem Absensi Karyawan!")
        print("Versi Command Line - Sederhana dan Mudah Digunakan\n")
        
        # Step 1: Pilih file
        file_path = select_file()
        
        if not file_path:
            print("\n❌ File tidak valid!")
            input("\nTekan Enter untuk keluar...")
            break
        
        # Load file
        print(f"\n⏳ Memuat file: {file_path}")
        success, message = processor.load_csv(file_path)
        
        if not success:
            print(f"\n❌ Error: {message}")
            input("\nTekan Enter untuk keluar...")
            break
        
        print(f"✓ {message}")
        input("\nTekan Enter untuk melanjutkan...")
        
        # Step 2: Pilih karyawan
        print_header()
        employee_list = processor.get_employee_list()
        selected_employee = select_employee(employee_list)
        
        if not selected_employee:
            print("\n❌ Tidak ada karyawan yang dipilih!")
            input("\nTekan Enter untuk keluar...")
            break
        
        print(f"\n✓ Karyawan terpilih: {selected_employee}")
        input("\nTekan Enter untuk melanjutkan...")
        
        # Step 3: Pilih periode
        print_header()
        month, year, month_name = select_period()
        
        print(f"\n✓ Periode terpilih: {month_name} {year}")
        input("\nTekan Enter untuk melanjutkan...")
        
        # Step 4: Nama file output
        print_header()
        print("\n📄 NAMA FILE OUTPUT")
        print("-" * 50)
        
        default_filename = f"Rekap_Absensi_{month_name}_{year}.xlsx"
        print(f"\nNama file default: {default_filename}")
        
        custom = input("\nGunakan nama default? (Y/n): ").strip().upper()
        
        if custom == 'N':
            output_file = input("Masukkan nama file: ").strip()
            if not output_file.endswith('.xlsx'):
                output_file += '.xlsx'
        else:
            output_file = default_filename
        
        # Step 5: Proses data
        print_header()
        print("\n⏳ MEMPROSES DATA...")
        print("-" * 50)
        print(f"\nKaryawan : {selected_employee}")
        print(f"Periode  : {month_name} {year}")
        print(f"Output   : {output_file}")
        print()
        
        success, message, df = processor.process_employee_attendance(
            selected_employee, month, year
        )
        
        if not success:
            print(f"\n❌ Error: {message}")
            input("\nTekan Enter untuk keluar...")
            break
        
        print(f"✓ {message}")
        print(f"✓ Total hari diproses: {len(df)}")
        
        # Save to Excel
        print("\n⏳ Menyimpan ke file Excel...")
        save_success, save_message = processor.save_to_excel(df, output_file)
        
        if save_success:
            print(f"\n✅ SUKSES! {save_message}")
            print("\nRingkasan:")
            print(f"- Karyawan   : {selected_employee}")
            print(f"- Periode    : {month_name} {year}")
            print(f"- Total hari : {len(df)}")
            print(f"- File output: {output_file}")
            print(f"- Ukuran file: {os.path.getsize(output_file) / 1024:.2f} KB")
        else:
            print(f"\n❌ Error: {save_message}")
        
        # Tanya apakah ingin proses lagi
        print("\n" + "=" * 70)
        again = input("\nProses data lain? (y/N): ").strip().upper()
        
        if again != 'Y':
            print("\n👋 Terima kasih telah menggunakan Sistem Absensi Karyawan!")
            print("=" * 70)
            break


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️  Program dihentikan oleh user")
        print("👋 Terima kasih!")
    except Exception as e:
        print(f"\n\n❌ Error tidak terduga: {str(e)}")
        input("\nTekan Enter untuk keluar...")
