@echo off
echo ====================================================
echo  RATIG - Clean Build (EXE Production)
echo ====================================================

echo.
echo [1/4] Cleaning up previous builds...
if exist build     rmdir /s /q build
if exist dist      rmdir /s /q dist
if exist release   rmdir /s /q release
if exist __pycache__ rmdir /s /q __pycache__
echo     Done.

echo.
echo [2/4] Installing dependencies...
python -m pip install --upgrade pip >nul 2>&1
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)
echo     Done.

echo.
echo [3/4] Building EXE using PyInstaller (spec file)...
python -m PyInstaller --clean --noconfirm RATIG.spec
if errorlevel 1 (
    echo.
    echo [ERROR] PyInstaller build failed. Check output above.
    pause
    exit /b 1
)
echo     Build successful.

echo.
echo [4/4] Preparing release folder...
mkdir release\RATIG_Portable
if exist dist\RATIG.exe copy dist\RATIG.exe release\RATIG_Portable\ /Y
if exist dist\RATIG xcopy dist\RATIG release\RATIG_Portable /E /I /H /Y


:: Buat struktur folder yang dibutuhkan aplikasi
mkdir release\RATIG_Portable\data
mkdir release\RATIG_Portable\data\excel
mkdir release\RATIG_Portable\data\database
mkdir release\RATIG_Portable\reports
mkdir release\RATIG_Portable\logs
mkdir release\RATIG_Portable\config
mkdir release\RATIG_Portable\assets
mkdir release\RATIG_Portable\backups

:: Copy icon jika ada
if exist assets\ratig.ico copy assets\ratig.ico release\RATIG_Portable\assets\ /Y

echo     Release folder ready.

echo.
echo ====================================================
echo  Clean Build Selesai!
echo  File EXE ada di: release\RATIG_Portable\RATIG.exe
echo ====================================================
pause
