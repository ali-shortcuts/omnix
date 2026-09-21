@echo off
setlocal enabledelayedexpansion
title OMNIX Diagnose
color 0A
echo ============================================================
echo   OMNIX DIAGNOSE - one-click full status check
echo   Just read what appears below, or screenshot this whole
echo   window (scroll up if needed) and send it back.
echo ============================================================
echo.

echo --- 1) Install folder contents ---
echo Path: %LOCALAPPDATA%\Programs\OMNIX
if exist "%LOCALAPPDATA%\Programs\OMNIX" (
    dir /b "%LOCALAPPDATA%\Programs\OMNIX"
) else (
    echo    ^>^>^> FOLDER DOES NOT EXIST — OMNIX was never installed here.
)
echo.

echo --- 2) Registry: Excel 16.0 Addins\OMNIX ---
reg query "HKCU\Software\Microsoft\Office\16.0\Excel\Addins\OMNIX" 2>nul
if errorlevel 1 echo    ^>^>^> KEY NOT FOUND.
echo.

echo --- 3) Registry: Excel 15.0 Addins\OMNIX ---
reg query "HKCU\Software\Microsoft\Office\15.0\Excel\Addins\OMNIX" 2>nul
if errorlevel 1 echo    ^>^>^> KEY NOT FOUND.
echo.

echo --- 4) Registry: Excel 16.0 Resiliency\DisabledItems (raw) ---
reg query "HKCU\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems" 2>nul
if errorlevel 1 echo    ^>^>^> KEY NOT FOUND / EMPTY.
echo.

echo --- 5) VSTO Runtime x64 installed? ---
reg query "HKLM\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R" /v Version 2>nul
if errorlevel 1 echo    ^>^>^> NOT FOUND (x64 view).
echo.

echo --- 6) OMNIX logs folder ---
echo Path: %LOCALAPPDATA%\OMNIX\logs
if exist "%LOCALAPPDATA%\OMNIX\logs" (
    dir /b "%LOCALAPPDATA%\OMNIX\logs"
    echo.
    echo --- install-debug.log (full content) ---
    if exist "%LOCALAPPDATA%\OMNIX\logs\install-debug.log" (
        type "%LOCALAPPDATA%\OMNIX\logs\install-debug.log"
    ) else (
        echo    ^>^>^> install-debug.log NOT FOUND.
    )
    echo.
    echo --- startup-debug.log (full content) ---
    if exist "%LOCALAPPDATA%\OMNIX\logs\startup-debug.log" (
        type "%LOCALAPPDATA%\OMNIX\logs\startup-debug.log"
    ) else (
        echo    ^>^>^> startup-debug.log NOT FOUND — Excel never even attempted to load OMNIX.
    )
    echo.
    echo --- post-install-verify.log (automatic COM check done by Setup itself) ---
    if exist "%LOCALAPPDATA%\OMNIX\logs\post-install-verify.log" (
        type "%LOCALAPPDATA%\OMNIX\logs\post-install-verify.log"
    ) else (
        echo    ^>^>^> post-install-verify.log NOT FOUND — this OMNIX version predates this check, or install failed before ssPostInstall.
    )
    echo.
    echo --- runtime-journey.log (human-readable prompt/provider/tool timeline) ---
    if exist "%LOCALAPPDATA%\OMNIX\logs\runtime-journey.log" (
        powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-Content -LiteralPath $env:LOCALAPPDATA\OMNIX\logs\runtime-journey.log -Tail 200"
    ) else (
        echo    ^>^>^> runtime-journey.log NOT FOUND — run at least one OMNIX request first.
    )
    echo.
    echo --- runtime-events.jsonl (structured machine-readable runtime events) ---
    if exist "%LOCALAPPDATA%\OMNIX\logs\runtime-events.jsonl" (
        echo Path: %LOCALAPPDATA%\OMNIX\logs\runtime-events.jsonl
        powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-Content -LiteralPath $env:LOCALAPPDATA\OMNIX\logs\runtime-events.jsonl -Tail 80"
    ) else (
        echo    ^>^>^> runtime-events.jsonl NOT FOUND — run at least one OMNIX request first.
    )
) else (
    echo    ^>^>^> LOGS FOLDER DOES NOT EXIST AT ALL.
)
echo.

echo --- 7) Is Excel currently running? ---
tasklist /fi "imagename eq excel.exe" | find /i "excel.exe" >nul
if %errorlevel%==0 (
    echo    Excel IS running right now.
) else (
    echo    Excel is NOT running right now.
)
echo.

echo ============================================================
echo   DONE. Please screenshot this entire window (scroll up
echo   first if some lines are missing) and send it back.
echo ============================================================
pause
