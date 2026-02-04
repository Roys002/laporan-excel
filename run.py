#!/usr/bin/env python3
"""
Sistem Absensi Karyawan - Auto Launcher
Otomatis memilih versi GUI atau CLI
"""

import sys
import os


def main():
    """Main launcher function"""
    
    print("=" * 70)
    print("📊 SISTEM ABSENSI KARYAWAN".center(70))
    print("=" * 70)
    print()
    
    # Cek apakah Tkinter tersedia
    try:
        import tkinter as tk
        has_tkinter = True
    except ImportError:
        has_tkinter = False
    
    # Cek apakah pandas tersedia
    try:
        import pandas
        has_pandas = True
    except ImportError:
        has_pandas = False
    
    # Jika pandas tidak ada, instruksi install
    if not has_pandas:
        print("❌ Pandas belum terinstall!")
        print("\nSilakan install dependencies terlebih dahulu:")
        print("  pip install pandas openpyxl")
        print("\natau:")
        print("  pip3 install pandas openpyxl")
        print("\natau:")
        print("  python -m pip install pandas openpyxl")
        input("\nTekan Enter untuk keluar...")
        return
    
    # Pilihan mode
    print("Pilih mode aplikasi:\n")
    
    if has_tkinter:
        print("1. Mode GUI (Recommended)")
        print("   - Tampilan grafis yang modern dan intuitif")
        print("   - Mudah digunakan dengan mouse")
        print()
    
    print("2. Mode Command Line")
    print("   - Tampilan text-based sederhana")
    print("   - Cocok untuk server atau sistem tanpa GUI")
    print()
    
    if has_tkinter:
        print("0. Keluar")
        print()
        
        while True:
            choice = input("Pilihan Anda (1/2/0): ").strip()
            
            if choice == '1':
                print("\n🚀 Meluncurkan mode GUI...")
                try:
                    from absensi_app import main as gui_main
                    gui_main()
                except Exception as e:
                    print(f"\n❌ Error meluncurkan GUI: {str(e)}")
                    input("\nTekan Enter untuk keluar...")
                break
            elif choice == '2':
                print("\n🚀 Meluncurkan mode Command Line...")
                try:
                    from absensi_cli import main as cli_main
                    cli_main()
                except Exception as e:
                    print(f"\n❌ Error meluncurkan CLI: {str(e)}")
                    input("\nTekan Enter untuk keluar...")
                break
            elif choice == '0':
                print("\n👋 Terima kasih!")
                break
            else:
                print("❌ Pilihan tidak valid. Silakan pilih 1, 2, atau 0")
    else:
        print("ℹ️  Tkinter tidak tersedia. Meluncurkan mode Command Line...")
        input("\nTekan Enter untuk melanjutkan...")
        try:
            from absensi_cli import main as cli_main
            cli_main()
        except Exception as e:
            print(f"\n❌ Error meluncurkan CLI: {str(e)}")
            input("\nTekan Enter untuk keluar...")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️  Program dihentikan")
    except Exception as e:
        print(f"\n\n❌ Error: {str(e)}")
        input("\nTekan Enter untuk keluar...")
