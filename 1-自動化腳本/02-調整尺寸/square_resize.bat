@echo off
chcp 65001 >nul
python "%~dp0square_resize.py" %*
if errorlevel 1 pause
