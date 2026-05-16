@echo off
chcp 65001 >nul
echo =====================================================
echo   公文自动格式化工具 v5.1.0 (Windows版^)
echo   完整支持：.docx, .doc, .wps
echo =====================================================
echo.

python "%~dp0src\main.py" %*

if errorlevel 1 pause
