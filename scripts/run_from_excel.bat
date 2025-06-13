@echo off
REM Запускает Excel из правильного venv и сразу открывает книгу
SET BASE=%~dp0..
CALL "%BASE%\venv\Scripts\activate.bat"
START "" "%ProgramFiles%\Microsoft Office\root\Office16\EXCEL.EXE" ^
      "%BASE%\excel\Finmodel.xlsm"
