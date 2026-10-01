@echo off
chcp 65001 >nul
node "%~dp0upload.js"
pause
