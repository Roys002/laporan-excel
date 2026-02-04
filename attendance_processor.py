"""
Attendance Data Processor
Sistem pengolahan data absensi dengan auto-detect format
"""

import pandas as pd
from datetime import datetime, timedelta
import calendar


class AttendanceProcessor:
    """Kelas untuk memproses data absensi"""
    
    # Hari libur nasional Indonesia 2026
    INDONESIAN_HOLIDAYS = {
        '2026': {
            '01-01': 'Tahun Baru 2026',
            '03-23': 'Isra Miraj',
            '03-31': 'Hari Suci Nyepi',
            '04-03': 'Wafat Yesus Kristus',
            '04-05': 'Paskah',
            '05-01': 'Hari Buruh',
            '05-04': 'Kenaikan Yesus Kristus',
            '05-14': 'Hari Raya Waisak',
            '06-01': 'Hari Lahir Pancasila',
            '06-17': 'Idul Fitri',
            '06-18': 'Idul Fitri',
            '08-17': 'Hari Kemerdekaan RI',
            '08-24': 'Idul Adha',
            '09-14': 'Tahun Baru Islam',
            '11-23': 'Maulid Nabi Muhammad SAW',
            '12-25': 'Hari Raya Natal',
            '12-26': 'Cuti Bersama Natal',
        }
    }
    
    def __init__(self):
        self.df_source = None
        self.employee_list = []
    
    def detect_csv_format(self, file_path):
        """Deteksi format CSV (semicolon atau comma)"""
        try:
            # Coba baca dengan semicolon
            df = pd.read_csv(file_path, sep=';', dtype=str, nrows=5)
            if len(df.columns) > 3:
                return ';'
            
            # Coba baca dengan comma
            df = pd.read_csv(file_path, sep=',', dtype=str, nrows=5)
            if len(df.columns) > 3:
                return ','
            
            return ';'  # default
        except Exception:
            return ';'
    
    def load_csv(self, file_path):
        """Load CSV file dengan auto-detect separator"""
        try:
            separator = self.detect_csv_format(file_path)
            self.df_source = pd.read_csv(file_path, sep=separator, dtype=str)
            
            # Bersihkan nama kolom
            self.df_source.columns = self.df_source.columns.astype(str).str.strip().str.lower()
            
            # Ekstrak daftar pegawai
            if 'nama' in self.df_source.columns:
                self.employee_list = self.df_source['nama'].dropna().tolist()
                return True, f"Berhasil memuat {len(self.employee_list)} karyawan"
            else:
                return False, "Kolom 'nama' tidak ditemukan dalam file CSV"
                
        except Exception as e:
            return False, f"Error membaca file: {str(e)}"
    
    def get_hari_indonesia(self, date_obj):
        """Mengembalikan nama hari dalam Bahasa Indonesia"""
        days = {
            'Monday': 'Senin', 'Tuesday': 'Selasa', 'Wednesday': 'Rabu',
            'Thursday': 'Kamis', 'Friday': 'Jumat', 'Saturday': 'Sabtu',
            'Sunday': 'Minggu'
        }
        return days[date_obj.strftime('%A')]
    
    def is_holiday(self, date_obj):
        """Cek apakah tanggal adalah hari libur nasional"""
        year = str(date_obj.year)
        date_str = date_obj.strftime('%m-%d')
        
        if year in self.INDONESIAN_HOLIDAYS:
            if date_str in self.INDONESIAN_HOLIDAYS[year]:
                return True, self.INDONESIAN_HOLIDAYS[year][date_str]
        
        return False, None
    
    def hitung_durasi(self, masuk_str, pulang_str):
        """Menghitung selisih waktu dari string jam"""
        try:
            fmt = '%H:%M:%S'
            t_masuk = datetime.strptime(str(masuk_str).strip(), fmt)
            t_pulang = datetime.strptime(str(pulang_str).strip(), fmt)
            
            selisih = t_pulang - t_masuk
            total_seconds = int(selisih.total_seconds())
            
            jam = total_seconds // 3600
            sisa_detik = total_seconds % 3600
            menit = sisa_detik // 60
            
            return f"{jam} jam {menit} menit"
        except Exception:
            return "-"
    
    def process_employee_attendance(self, employee_name, month=None, year=None):
        """Proses data absensi untuk karyawan tertentu"""
        
        if self.df_source is None:
            return False, "Data belum dimuat. Silakan load CSV terlebih dahulu.", None
        
        # Cari pegawai
        if 'nama' not in self.df_source.columns:
            return False, "Kolom 'nama' tidak ditemukan", None
        
        pegawai = self.df_source[self.df_source['nama'].str.contains(
            employee_name, case=False, na=False)]
        
        if pegawai.empty:
            return False, f"Karyawan '{employee_name}' tidak ditemukan", None
        
        # Kolom tanggal adalah kolom selain 'nama', 'nik', 'no', 'nomor'
        abaikan = ['nama', 'nik', 'no', 'nomor']
        date_columns = [col for col in self.df_source.columns if col not in abaikan]
        
        output_data = []
        
        for col in date_columns:
            # Parsing Header Tanggal
            try:
                # Coba format DD-MM-YYYY
                date_obj = datetime.strptime(col, '%d-%m-%Y')
            except ValueError:
                try:
                    # Coba format YYYY-MM-DD
                    date_obj = datetime.strptime(col, '%Y-%m-%d')
                except ValueError:
                    try:
                        # Coba format DD/MM/YYYY
                        date_obj = datetime.strptime(col, '%d/%m/%Y')
                    except ValueError:
                        continue
            
            # Filter berdasarkan bulan dan tahun jika ditentukan
            if month and date_obj.month != month:
                continue
            if year and date_obj.year != year:
                continue
            
            tgl_str_clean = date_obj.strftime('%Y-%m-%d')
            hari = self.get_hari_indonesia(date_obj)
            
            # Cek hari libur
            is_libur, nama_libur = self.is_holiday(date_obj)
            
            # Ambil nilai sel absensi
            raw_value = pegawai.iloc[0][col]
            
            jam_masuk = "-"
            jam_pulang = "-"
            durasi = "-"
            keterangan = "-"
            
            # Logika Pengisian Status
            val_str = str(raw_value).strip()
            
            # Cek data kosong
            if pd.isna(raw_value) or val_str == "" or val_str.lower() == "nan":
                if is_libur:
                    keterangan = f"Libur - {nama_libur}"
                elif hari in ['Sabtu', 'Minggu']:
                    keterangan = "Libur Akhir Pekan"
                else:
                    keterangan = "Tidak Hadir / Tanpa Keterangan"
            else:
                # Format: "06:59:29 - 16:03:04"
                parts = val_str.split('-')
                
                if len(parts) >= 2:
                    jam_masuk = parts[0].strip()
                    jam_pulang = parts[1].strip()
                    
                    # Hanya hitung durasi jika format jam valid
                    if ':' in jam_masuk and ':' in jam_pulang:
                        durasi = self.hitung_durasi(jam_masuk, jam_pulang)
                    
                    if is_libur:
                        keterangan = f"Hadir di Hari Libur ({nama_libur})"
                    else:
                        keterangan = "Hadir"
                        
                elif len(parts) == 1 and parts[0] != "":
                    jam_masuk = parts[0].strip()
                    keterangan = "Lupa Absen Pulang/Datang"
            
            output_data.append({
                'Tanggal': tgl_str_clean,
                'Hari': hari,
                'Jam Masuk': jam_masuk,
                'Jam Pulang': jam_pulang,
                'Durasi Kerja': durasi,
                'Keterangan': keterangan
            })
        
        if not output_data:
            return False, "Tidak ada data tanggal yang berhasil diproses", None
        
        df_result = pd.DataFrame(output_data)
        return True, "Data berhasil diproses", df_result
    
    def save_to_excel(self, df, output_file):
        """Simpan DataFrame ke file Excel"""
        try:
            df.to_excel(output_file, index=False)
            return True, f"File berhasil disimpan: {output_file}"
        except Exception as e:
            return False, f"Gagal menyimpan file: {str(e)}"
    
    def get_employee_list(self):
        """Dapatkan list karyawan"""
        return self.employee_list
    
    @staticmethod
    def get_month_name(month_number):
        """Dapatkan nama bulan dalam Bahasa Indonesia"""
        months = {
            1: 'Januari', 2: 'Februari', 3: 'Maret', 4: 'April',
            5: 'Mei', 6: 'Juni', 7: 'Juli', 8: 'Agustus',
            9: 'September', 10: 'Oktober', 11: 'November', 12: 'Desember'
        }
        return months.get(month_number, '')


if __name__ == "__main__":
    # Test processor
    processor = AttendanceProcessor()
    
    # Test dengan file yang ada
    success, message = processor.load_csv('absen jan 26 ekspor_csv.csv')
    print(f"Load CSV: {message}")
    
    if success:
        print(f"\nDaftar karyawan: {len(processor.get_employee_list())}")
        
        # Test proses untuk satu karyawan
        success, message, df = processor.process_employee_attendance(
            "I MADE BRAHMANDA SETYADI, S.Kom", 
            month=1, 
            year=2026
        )
        
        if success:
            print(f"\n{message}")
            print(f"Total hari: {len(df)}")
            print("\nSample data:")
            print(df.head(10))
