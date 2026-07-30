@echo off
title DocuShift PRO - Word to PDF Converter
color 0B
echo ========================================================
echo   DocuShift PRO - Word to PDF Converter
echo ========================================================
echo   Starting local server...
echo.

set AUTO_OPEN=true

if exist "%~dp0bin\node.exe" (
    "%~dp0bin\node.exe" "%~dp0server.js"
) else (
    node "%~dp0server.js"
)

pause
