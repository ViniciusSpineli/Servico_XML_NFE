@echo off
setlocal

cd /d "%~dp0"

if exist ".\venv\Scripts\python.exe" (
    ".\venv\Scripts\python.exe" ".\monitorar_baixa_xml.py"
) else (
    python ".\monitorar_baixa_xml.py"
)
