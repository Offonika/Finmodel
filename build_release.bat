@echo off
chcp 65001 >nul
setlocal ENABLEDELAYEDEXPANSION

rem Activate or create venv
if not exist "venv" (
    python -m venv venv || goto :error
)
call "venv\Scripts\activate.bat"

rem Ensure required packages
python -m pip install --upgrade pip >nul
pip install -U pyinstaller pyarmor==7.* >nul || goto :error

rem Obfuscate scripts
pyarmor gen --recursive -O obf_scripts scripts || goto :error

rem Build single-file executable
pyarmor pack -e "--onefile --name Finmodel" obf_scripts/Finmodel.py || goto :error

set exe_path=%CD%\dist\Finmodel.exe
if exist "%exe_path%" (
    echo Release created: %exe_path%
) else (
    echo Build failed: %exe_path% not found.
)

endlocal
exit /b 0

:error
echo Failed to build release.
endlocal
exit /b 1
