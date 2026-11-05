# setup.ps1 — установка окружения проекта (Windows, PowerShell)
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

Write-Host "==> Создаю виртуальное окружение .venv"
python -m venv .venv

& "$PSScriptRoot\.venv\Scripts\python.exe" -m pip install --upgrade pip
Write-Host "==> Ставлю зависимости"
& "$PSScriptRoot\.venv\Scripts\python.exe" -m pip install -r requirements.txt

Write-Host "==> Проверяю тесты"
& "$PSScriptRoot\.venv\Scripts\python.exe" -m pytest -v

Write-Host "Готово. Активация окружения: .\.venv\Scripts\Activate.ps1"
