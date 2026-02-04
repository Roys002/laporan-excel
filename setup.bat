@echo off
REM Quick Start Script untuk Sistem Absensi Karyawan (Windows)

echo ======================================================================
echo    [92m📊 SISTEM ABSENSI KARYAWAN - QUICK START[0m
echo ======================================================================
echo.

REM Cek Python
echo [96m🔍 Memeriksa instalasi Python...[0m
python --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=python
    echo [92m✓ Python ditemukan[0m
    python --version
) else (
    python3 --version >nul 2>&1
    if %errorlevel% equ 0 (
        set PYTHON_CMD=python3
        echo [92m✓ Python3 ditemukan[0m
        python3 --version
    ) else (
        echo [91m❌ Python tidak ditemukan![0m
        echo    Silakan install Python 3.7+ dari python.org
        pause
        exit /b 1
    )
)

echo.

REM Cek pip
echo [96m🔍 Memeriksa pip...[0m
%PYTHON_CMD% -m pip --version >nul 2>&1
if %errorlevel% equ 0 (
    echo [92m✓ pip tersedia[0m
) else (
    echo [91m❌ pip tidak ditemukan![0m
    pause
    exit /b 1
)

echo.

REM Install dependencies
echo [96m📦 Menginstall dependencies...[0m
echo    - pandas
echo    - openpyxl
echo.

%PYTHON_CMD% -m pip install pandas openpyxl

if %errorlevel% equ 0 (
    echo.
    echo [92m✅ Dependencies berhasil diinstall![0m
) else (
    echo.
    echo [93m⚠️  Gagal menginstall dependencies[0m
    echo    Coba manual: pip install pandas openpyxl
)

echo.
echo ======================================================================
echo [92m✅ SETUP SELESAI![0m
echo ======================================================================
echo.
echo Cara menjalankan aplikasi:
echo   python run.py
echo.
echo Untuk GUI:
echo   python absensi_app.py
echo.
echo Untuk Command Line:
echo   python absensi_cli.py
echo.
echo ======================================================================
pause
