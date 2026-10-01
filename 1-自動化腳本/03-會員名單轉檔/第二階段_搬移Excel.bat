@echo off
chcp 65001 >nul
python "%~dp0member_csv.py" move
pause
