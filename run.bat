@echo off
rem run.bat — прогон генератора заявок: читает xlsx из input/, пишет docx в output/
chcp 65001 >nul
setlocal
cd /d "%~dp0"

set "PY=%~dp0.venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python"

"%PY%" -c "import docxtpl, openpyxl, yaml" 2>nul
if errorlevel 1 (
    echo Зависимости не установлены, ставлю requirements.txt...
    "%PY%" -m pip install --upgrade pip
    "%PY%" -m pip install -r requirements.txt
)

echo === Прогон файлов из input^/ ===
"%PY%" -m src.main --input input --config config/config.yaml --output output
set "RC=%ERRORLEVEL%"

if not "%RC%"=="0" echo Ошибка: код возврата %RC% 1>&2
endlocal & exit /b %RC%
