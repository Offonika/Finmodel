@echo off
chcp 65001 >nul
setlocal ENABLEDELAYEDEXPANSION

rem ╔══════════════════════════════════════════════════════════════════╗
rem ║   Finmodel – лёгкая установка для конечного пользователя        ║
rem ╚══════════════════════════════════════════════════════════════════╝

set "PROJECT_PATH=%~dp0"
set "LOGFILE=%PROJECT_PATH%install_user.log"
echo ==== %date% %time% ==== Старт установки =========================>>"%LOGFILE%"

echo 📁 Папка проекта: %PROJECT_PATH%

rem --- 1. Проверяем наличие шаблона Excel -----------------------------------
if not exist "%PROJECT_PATH%excel\Finmodel_Template.xlsm" (
    echo ❌ Шаблон excel\Finmodel_Template.xlsm не найден!
    echo [ERR] Template missing>>"%LOGFILE%"
    pause & exit /b 1
)

rem --- 2. Копируем Finmodel.xlsm, если ещё нет ------------------------------
if exist "%PROJECT_PATH%Finmodel.xlsm" (
    echo ℹ️  Файл Finmodel.xlsm уже существует (пропускаю копирование)
    echo [INFO] Finmodel.xlsm exists>>"%LOGFILE%"
) else (
    echo 📄 Создаю рабочую копию Excel-книги…
    copy "%PROJECT_PATH%excel\Finmodel_Template.xlsm" ^
         "%PROJECT_PATH%Finmodel.xlsm" >nul
    echo [INFO] Template copied to Finmodel.xlsm>>"%LOGFILE%"
)

rem --- 3. Создаём ярлык на рабочем столе ------------------------------------
set "LNK=%USERPROFILE%\Desktop\Finmodel.lnk"
if exist "%LNK%" (
    echo ℹ️  Ярлык уже есть на рабочем столе
    echo [INFO] Shortcut exists>>"%LOGFILE%"
) else (
    echo 🔗 Создаю ярлык на рабочем столе…
    powershell -NoLogo -NoProfile -Command ^
      "$s=(New-Object -COM WScript.Shell).CreateShortcut('%LNK%');" ^
      "$s.TargetPath='%PROJECT_PATH%Finmodel.xlsm';" ^
      "$s.WorkingDirectory='%PROJECT_PATH%';" ^
      "$s.Save()" 
    echo [INFO] Shortcut created>>"%LOGFILE%"
)

rem --- 4. Финал -------------------------------------------------------------
echo ✅ Установка завершена! Откройте ярлык «Finmodel» на рабочем столе.
echo ==== %date% %time% ==== Конец установки ==========================>>"%LOGFILE%"
pause
endlocal
