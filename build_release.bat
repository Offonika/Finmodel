@echo off
rem ═════════════════════════════════════════════════════════════════════════════
rem build_release.bat – сборка всех *.py из  папки /scripts в одинарные EXE
rem (PyInstaller, console-mode).  Лог в build.log, готовые EXE в /release_exe
rem ═════════════════════════════════════════════════════════════════════════════

chcp 65001 > nul
setlocal EnableDelayedExpansion

:: ────── папки проекта ─────────────────────────────────────────────────────────
set "PROJECT_DIR=%~dp0"
set "SRC_DIR=%PROJECT_DIR%scripts"
set "DIST_DIR=%PROJECT_DIR%release_exe"
set "LOG=%PROJECT_DIR%build.log"

:: ────── начало лога ───────────────────────────────────────────────────────────
echo ==== %date% %time% ==== НАЧАЛО СБОРКИ ====================================>> "%LOG%"
echo [INFO] Проект: %PROJECT_DIR%
echo [INFO] Скрипты: %SRC_DIR%
echo.

:: ────── виртуальное окружение (venv) ─────────────────────────────────────────
if not exist "%PROJECT_DIR%venv" (
    echo [SETUP] Создаю виртуальное окружение ...
    python -m venv "%PROJECT_DIR%venv" || goto :err
)

call "%PROJECT_DIR%venv\Scripts\activate.bat"

:: ────── зависимости ──────────────────────────────────────────────────────────
echo [SETUP] Обновляю pip и ставлю PyInstaller ...
python -m pip install -U pip        >> "%LOG%" 2>&1
python -m pip install -U pyinstaller>> "%LOG%" 2>&1

:: ────── очистка результатов прошлой сборки ───────────────────────────────────
echo [CLEAN] Очищаю старые build/dist...
rd /s /q "%DIST_DIR%"     2>nul
rd /s /q "%PROJECT_DIR%build" 2>nul
rd /s /q "%PROJECT_DIR%dist"  2>nul
md "%DIST_DIR%"

:: ────── сборка всех скриптов ─────────────────────────────────────────────────
echo.
echo ===== СТАРТ ПОЛНЫХ СБОРОК =================================================

for %%F in ("%SRC_DIR%\*.py") do (
    set "FILEPATH=%%~fF"
    set "FILENAME=%%~nxF"
    set "NAME=%%~nF"

    echo [BUILD] %%~nxF → !NAME!.exe
    echo [BUILD] %%~nxF >> "%LOG%"

    pyinstaller ^
        --onefile --noconfirm ^
        --name "!NAME!" ^
        "!FILEPATH!"            >> "%LOG%" 2>&1

    if exist "%PROJECT_DIR%dist\!NAME!.exe" (
        copy /y "%PROJECT_DIR%dist\!NAME!.exe" "%DIST_DIR%\!NAME!.exe" >nul
        echo      ✔ !NAME!.exe готов
        echo [OK] !NAME!.exe собран >> "%LOG%"
    ) else (
        echo      ✖ !NAME!.exe НЕ собран  (смотрите build.log)
        echo [FAIL] !NAME!.exe не собран >> "%LOG%"
    )
)

:: ────── финальная очистка build/dist от PyInstaller ──────────────────────────
rd /s /q "%PROJECT_DIR%build" 2>nul
rd /s /q "%PROJECT_DIR%dist"  2>nul

echo.
echo ✅ Сборка завершена. EXE-файлы: %DIST_DIR%
echo ==== %date% %time% ==== КОНЕЦ СБОРКИ =====================================>> "%LOG%"
goto :eof

:err
echo [ERROR] Сборка прервана. Подробности – в build.log
exit /b 1
