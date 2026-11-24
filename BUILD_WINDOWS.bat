@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"
set "LOG=%~dp0build_log.txt"
echo SINETECHNIC GIDER derleme - %date% %time% > "%LOG%"

echo.
echo ===== SINETECHNIC GIDER - EXE DERLEME =====
echo.

rem ---- 1) Python bul (py launcher > python > winget ile kur)
set "PY="
py -3 --version >nul 2>&1 && set "PY=py -3"
if not defined PY (
  python --version >nul 2>&1 && python -c "import sys" >nul 2>&1 && set "PY=python"
)
if not defined PY (
  echo [1/4] Python bulunamadi, winget ile Python 3.12 kuruluyor...
  echo Python yok, winget ile kuruluyor >> "%LOG%"
  winget install -e --id Python.Python.3.12 --scope user --accept-package-agreements --accept-source-agreements >> "%LOG%" 2>&1
  if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PY="%LOCALAPPDATA%\Programs\Python\Python312\python.exe""
  if not defined PY (
    py -3 --version >nul 2>&1 && set "PY=py -3"
  )
)
if not defined PY (
  echo.
  echo HATA: Python kurulamadi. https://www.python.org/downloads/ adresinden
  echo Python 3.12 kurun ^(kurulumda "Add python.exe to PATH" kutusunu isaretleyin^)
  echo ve bu dosyayi tekrar calistirin.
  echo HATA: Python yok >> "%LOG%"
  pause
  exit /b 1
)
echo [1/4] Python: %PY%
%PY% --version >> "%LOG%" 2>&1

rem ---- 2) Paketler
echo [2/4] Gerekli paketler yukleniyor (ilk seferde birkac dakika surebilir)...
%PY% -m pip install --upgrade pip >> "%LOG%" 2>&1
%PY% -m pip install PySide6 openpyxl Pillow reportlab qrcode pypdf pillow-heif pyinstaller opencv-python-headless rapidocr_onnxruntime pypdfium2 >> "%LOG%" 2>&1
if errorlevel 1 (
  echo HATA: Paketler yuklenemedi. Ayrintilar build_log.txt dosyasinda.
  pause
  exit /b 1
)
%PY% -m pip install winocr >> "%LOG%" 2>&1 || echo winocr yuklenemedi, OCR'siz devam >> "%LOG%"

rem ---- 3) Derle
echo [3/4] EXE derleniyor...
set "EXTRA="
%PY% -c "import winocr" >nul 2>&1 && set "EXTRA=--collect-all winocr --collect-all winrt"
%PY% -c "import pillow_heif" >nul 2>&1 && set "EXTRA=%EXTRA% --collect-all pillow_heif"
set "EXTRA=%EXTRA% --collect-all rapidocr_onnxruntime --collect-all onnxruntime --collect-submodules cv2 --collect-all pypdfium2 --collect-all pypdfium2_raw"
%PY% -m PyInstaller --noconfirm --clean --onefile --windowed --name "SINETECHNIC_GIDER" --icon "assets\icon.ico" --add-data "assets;assets" %EXTRA% gider.py >> "%LOG%" 2>&1
if not exist "dist\SINETECHNIC_GIDER.exe" (
  echo HATA: EXE olusmadi. Ayrintilar build_log.txt dosyasinda.
  pause
  exit /b 1
)

rem ---- 4) Indirilenler'e kopyala
copy /Y "dist\SINETECHNIC_GIDER.exe" "%USERPROFILE%\Downloads\SINETECHNIC_GIDER.exe" >> "%LOG%" 2>&1
echo [4/4] Tamam! SINETECHNIC_GIDER.exe Indirilenler klasorune kopyalandi.
echo BASARILI >> "%LOG%"
echo.
pause
