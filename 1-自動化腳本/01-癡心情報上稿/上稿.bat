@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "PATH=C:\Program Files\nodejs;%PATH%"
echo.
echo ===== 癡心情報上稿 =====
echo.
node post.js %*
echo.
pause
