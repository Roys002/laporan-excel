import pandas as pd
from datetime import datetime
import os
import sys

# --- KONFIGURASI ---
# Pastikan nama ini SAMA PERSIS dengan yang ada di kolom 'nama' di CSV
NAMA_TARGET = "I MADE BRAHMANDA SETYADI, S.Kom"
FILE_SUMBER = 'absen jan 26 ekspor_csv.csv'
OUTPUT_FILE = 'Rekap_Absensi_Januari_2026.xlsx'

def hitung_durasi(masuk_str, pulang_str):
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

def get_hari_indonesia(date_obj):
    """Mengembalikan nama hari dalam Bahasa Indonesia"""
    days = {
        'Monday': 'Senin', 'Tuesday': 'Selasa', 'Wednesday': 'Rabu',
        'Thursday': 'Kamis', 'Friday': 'Jumat', 'Saturday': 'Sabtu',
        'Sunday': 'Minggu'
    }
    return days[date_obj.strftime('%A')]

def process_attendance():
    print(f"Sedang memproses data untuk: {NAMA_TARGET}...")

    # 1. BACA DATA SUMBER (KHUSUS CSV dengan pemisah ';')
    try:
        if os.path.exists(FILE_SUMBER):
            # Penting: sep=';' digunakan karena file CSV Anda dipisah titik koma
            # dtype=str memastikan semua data dibaca sebagai teks agar jam tidak berubah format
            df_source = pd.read_csv(FILE_SUMBER, sep=';', dtype=str)
        else:
            print(f"File {FILE_SUMBER} tidak ditemukan.")
            return
    except Exception as e:
        print(f"Gagal membaca file sumber: {e}")
        return

    # 2. BERSIHKAN NAMA KOLOM
    # Ubah header menjadi huruf kecil semua dan hapus spasi agar mudah diproses
    # Contoh: "Nama " -> "nama", "NIK" -> "nik"
    df_source.columns = df_source.columns.astype(str).str.strip().str.lower()

    # Debug: Tampilkan kolom yang terbaca untuk memastikan pemisah ';' bekerja
    print(f"Kolom terdeteksi ({len(df_source.columns)}): {df_source.columns.tolist()[:5]} ...")

    # 3. CARI PEGAWAI
    if 'nama' not in df_source.columns:
        print("Error: Kolom 'nama' tidak ditemukan. Pastikan file CSV menggunakan pemisah titik koma (;).")
        return

    # Cari baris data pegawai (case insensitive)
    pegawai = df_source[df_source['nama'].str.contains(NAMA_TARGET, case=False, na=False)]

    if pegawai.empty:
        print(f"Error: Pegawai atas nama '{NAMA_TARGET}' tidak ditemukan.")
        print("Coba cek penulisan nama di variabel NAMA_TARGET.")
        return

    print("Data pegawai ditemukan. Mengekstrak absensi...")

    # 4. OLAH DATA PER TANGGAL
    # Kolom tanggal adalah kolom selain 'nama' dan 'nik'
    abaikan = ['nama', 'nik', 'no', 'nomor']
    date_columns = [col for col in df_source.columns if col not in abaikan]

    output_data = []

    for col in date_columns:
        # Parsing Header Tanggal (Format di CSV Anda: 01-01-2026)
        try:
            # Coba format DD-MM-YYYY (paling umum di CSV Indonesia)
            date_obj = datetime.strptime(col, '%d-%m-%Y')
        except ValueError:
            try:
                # Coba format YYYY-MM-DD (cadangan)
                date_obj = datetime.strptime(col, '%Y-%m-%d')
            except ValueError:
                # Jika header kolom bukan tanggal (misal kolom 'Total'), lewati
                continue

        tgl_str_clean = date_obj.strftime('%Y-%m-%d')
        hari = get_hari_indonesia(date_obj)

        # Ambil nilai sel absensi
        raw_value = pegawai.iloc[0][col]

        jam_masuk = "-"
        jam_pulang = "-"
        durasi = "-"
        keterangan = "-"

        # Logika Pengisian Status
        val_str = str(raw_value).strip()

        # Cek data kosong (nan dalam string atau kosong)
        if pd.isna(raw_value) or val_str == "" or val_str.lower() == "nan":
            if hari in ['Sabtu', 'Minggu']:
                keterangan = "Libur Akhir Pekan"
            else:
                keterangan = "Tidak Hadir / Tanpa Keterangan"
        else:
            # Format di CSV Anda: "06:59:29 - 16:03:04"
            parts = val_str.split('-')

            if len(parts) >= 2:
                jam_masuk = parts[0].strip()
                jam_pulang = parts[1].strip()
                # Hanya hitung durasi jika format jam valid (XX:XX:XX)
                if ':' in jam_masuk and ':' in jam_pulang:
                    durasi = hitung_durasi(jam_masuk, jam_pulang)
                keterangan = "Hadir"
            elif len(parts) == 1 and parts[0] != "":
                # Kasus lupa absen pulang/datang
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

    # 5. SIMPAN HASIL
    if not output_data:
        print("Tidak ada data tanggal yang berhasil diproses.")
        return

    df_result = pd.DataFrame(output_data)

    try:
        df_result.to_excel(OUTPUT_FILE, index=False)
        print(f"\nSukses! Data absensi tersimpan di: {OUTPUT_FILE}")
        print(f"Total hari diproses: {len(df_result)}")
    except Exception as e:
        print(f"Gagal menyimpan file excel: {e}")

if __name__ == "__main__":
    process_attendance()