@echo off
chcp 65001 >nul

echo ========================================
echo       Update Log Viewer
echo ========================================
echo.

REM 查找最新的更新日志文件
set "LOG_FILE="
for /f "delims=" %%f in ('dir /b /o-d "%TEMP%\update_log_*.txt" 2^>nul') do (
    set "LOG_FILE=%TEMP%\%%f"
    goto :found
)

echo No update log file found in %TEMP%
echo.
pause
exit /b 0

:found
echo Latest update log: %LOG_FILE%
echo.
echo ========================================
echo File Contents:
echo ========================================
echo.
type "%LOG_FILE%"
echo.
echo ========================================
echo.
pause
