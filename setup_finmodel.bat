@chcp 65001 >nul
@echo off
setlocal ENABLEDELAYEDEXPANSION

:: ---------- 0. Инициализация -------------------------------------------------
set "PROJECT_PATH=%~dp0"
set "LOGFILE=%PROJECT_PATH%install.log"
echo ==== %date% %time% ==== Установка начата ==================================>>"%LOGFILE%"
echo 📁 Рабочая папка: %PROJECT_PATH%
echo [INFO] Рабочая папка: %PROJECT_PATH%>>"%LOGFILE%"

:: ---------- 1. Проверка Python ----------------------------------------------
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo 🔍 Python не найден, скачиваю… & echo [INFO] Python not found>>"%LOGFILE%"
    curl -o python-installer.exe ^
         https://www.python.org/ftp/python/3.11.8/python-3.11.8-amd64.exe >>"%LOGFILE%" 2>&1
    echo 🧱 Устанавливаю Python… & echo [INFO] Installing Python>>"%LOGFILE%"
    python-installer.exe /quiet InstallAllUsers=0 PrependPath=1 SimpleInstall=1 >>"%LOGFILE%" 2>&1
    timeout /t 10 >nul
)
where python >nul 2>nul || (
    echo ❌ Python не установлен — см. install.log
    echo [ERR] Python still missing, abort>>"%LOGFILE%"
    pause & exit /b
)

:: ---------- 2. Меню ----------------------------------------------------------
:menu
cls
echo === 🚀 УСТАНОВКА ФИНМОДЕЛИ ===
echo 1. Создать виртуальное окружение (venv)
echo 2. Установить зависимости
echo 3. Установить надстройку xlwings
echo 4. Всё сразу \(1→2→3\)
echo 5. Выход
echo.
set /p choice=Выберите пункт (1-5): 

if "%choice%"=="1" (call :venv      && goto menu)
if "%choice%"=="2" (call :deps      && goto menu)
if "%choice%"=="3" (call :addin     && goto menu)
if "%choice%"=="4" (call :venv  silent & call :deps  silent & call :addin  silent & goto finish)
if "%choice%"=="5" exit
goto menu

:: ---------- 3. Подпроцедуры -------------------------------------------------
:venv
if exist venv (
    echo ⚠️ venv уже существует
    echo [INFO] venv exists>>"%LOGFILE%"
) else (
    echo 🐍 Создаю venv…
    echo [INFO] Creating venv>>"%LOGFILE%"
    python -m venv venv >>"%LOGFILE%" 2>&1
)
if "%~1" neq "silent" pause
exit /b

:deps
call ".\venv\Scripts\activate.bat"
echo 📦 Ставлю зависимости…
echo [INFO] Installing requirements>>"%LOGFILE%"
python -m pip install --upgrade pip >>"%LOGFILE%" 2>&1
pip install -r requirements.txt       >>"%LOGFILE%" 2>&1
if "%~1" neq "silent" pause
exit /b

:addin
echo 🧩 Ставлю надстройку xlwings…
echo [INFO] Installing xlwings add-in>>"%LOGFILE%"
xlwings addin install >>"%LOGFILE%" 2>&1
if "%~1" neq "silent" pause
exit /b

:: ---------- 4. Финал ---------------------------------------------------------
:finish
rem 4.1 Копируем шаблон Excel
if not exist "excel\Finmodel.xlsm" (
    if exist "excel\Finmodel_Template.xlsm" (
        echo 📄 Копирую шаблон…
        echo [INFO] Copying template>>"%LOGFILE%"
        copy "excel\Finmodel_Template.xlsm" "excel\Finmodel.xlsm" >>"%LOGFILE%" 2>&1
    ) else (
        echo ❌ Шаблон не найден: excel\Finmodel_Template.xlsm
        echo [ERR] Template missing>>"%LOGFILE%"
    )
) else (
    echo ℹ️  excel\Finmodel.xlsm уже существует
    echo [INFO] Main xlsm exists>>"%LOGFILE%"
)


:: 4.2 Создаём / перезаписываем .xlwings.conf
echo 🧾 Пишу .xlwings.conf…
echo [INFO] writing .xlwings.conf>>"%LOGFILE%"

>  ".xlwings.conf" echo PROJECT_PATH = %%(CURRENT_PATH)s
>> ".xlwings.conf" echo PYTHONPATH   = scripts
>> ".xlwings.conf" echo INTERPRETER  = %%(PROJECT_PATH)s\venv\Scripts\python.exe

echo [INFO] .xlwings.conf written>>"%LOGFILE%"




echo ✅ Установка завершена! Логи: install.log
echo ==== %date% %time% ==== Конец ============================================>>"%LOGFILE%"
pause
goto menu
