#!/bin/bash
# Quick Start Script untuk Sistem Absensi Karyawan

echo "======================================================================"
echo "   📊 SISTEM ABSENSI KARYAWAN - QUICK START"
echo "======================================================================"
echo ""

# Cek Python
echo "🔍 Memeriksa instalasi Python..."
if command -v python3 &> /dev/null; then
    PYTHON_CMD="python3"
    echo "✓ Python3 ditemukan: $(python3 --version)"
elif command -v python &> /dev/null; then
    PYTHON_CMD="python"
    echo "✓ Python ditemukan: $(python --version)"
else
    echo "❌ Python tidak ditemukan!"
    echo "   Silakan install Python 3.7+ terlebih dahulu"
    exit 1
fi

echo ""

# Cek pip
echo "🔍 Memeriksa pip..."
if $PYTHON_CMD -m pip --version &> /dev/null; then
    echo "✓ pip tersedia"
else
    echo "❌ pip tidak ditemukan!"
    echo "   Install dengan: sudo apt install python3-pip"
    exit 1
fi

echo ""

# Install dependencies
echo "📦 Menginstall dependencies..."
echo "   - pandas"
echo "   - openpyxl"
echo ""

$PYTHON_CMD -m pip install --user pandas openpyxl

if [ $? -eq 0 ]; then
    echo ""
    echo "✅ Dependencies berhasil diinstall!"
else
    echo ""
    echo "⚠️  Gagal menginstall dependencies"
    echo "   Coba manual: pip3 install pandas openpyxl"
fi

echo ""
echo "======================================================================"
echo "✅ SETUP SELESAI!"
echo "======================================================================"
echo ""
echo "Cara menjalankan aplikasi:"
echo "  ./run.py"
echo "  atau"
echo "  python3 run.py"
echo ""
echo "Untuk GUI (jika Tkinter tersedia):"
echo "  python3 absensi_app.py"
echo ""
echo "Untuk Command Line:"
echo "  python3 absensi_cli.py"
echo ""
echo "======================================================================"
