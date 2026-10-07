@echo off
chcp 65001 >nul
python "%~dp0resize.py" --size 1040x1040 %*
if errorlevel 1 pause
