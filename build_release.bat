@echo off
chcp 65001 >nul
setlocal ENABLEDELAYEDEXPANSION

:: === НАСТРОЙКИ ===
set "PROJECT_PATH=%~dp0"
set "SRC_DIR=%PROJECT_PATH%scripts"
set "DIST_DIR=%PROJECT_PATH%release_exe"
set "LOGFILE=%PROJECT_PATH%build.log"

echo ==== %date% %time% ==== Начало сборки =====================================>>"%LOGFILE%"
echo 📁 Проект: %PROJECT_PATH%
echo 📂 Скрипты: %SRC_DIR%
echo [INFO] Сборка .exe без обфускации >>"%LOGFILE%"

:: === VENV ===
if not exist "venv" (
    python -m venv venv || goto :error
)
call "venv\Scripts\activate.bat"

:: === УСТАНОВКА ЗАВИСИМОСТЕЙ ===
python -m pip install --upgrade pip >>"%LOGFILE%" 2>&1
pip install -U pyinstaller >>"%LOGFILE%" 2>&1

:: === ОЧИСТКА ===
rd /s /q "%DIST_DIR%" 2>nul
mkdir "%DIST_DIR%"

:: === СБОРКА .EXE ДЛЯ КАЖДОГО СКРИПТА ===
for %%f in (%SRC_DIR%\*.py) do (
    set "FILENAME=%%~nxf"
    set "NAME=%%~nf"
    echo 🛠️  Сборка !FILENAME!...
    pyinstaller --onefile --noconfirm --noconsole --name "!NAME!" "%%f" >>"%LOGFILE%" 2>&1

    if exist "dist\!NAME!.exe" (
        copy "dist\!NAME!.exe" "%DIST_DIR%\!NAME!.exe" >nul
        echo [OK] !NAME!.exe собран >>"%LOGFILE%"
    ) else (
        echo [FAIL] !NAME!.exe НЕ собран >>"%LOGFILE%"
    )
)

:: === ОЧИСТКА ВРЕМЕННЫХ ФАЙЛОВ ===
:: === ОЧИСТКА ВРЕМЕННЫХ ФАЙЛОВ ===
rd /s /q "build" >nul 2>&1
rd /s /q "dist" >nul 2>&1

