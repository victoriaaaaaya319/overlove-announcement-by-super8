@echo off
chcp 65001 >nul
python "%~dp0resize.py" --size 960x1553 %*
if errorlevel 1 pause
