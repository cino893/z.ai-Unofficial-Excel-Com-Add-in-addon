@echo off
REM Prefer PowerShell 7 (pwsh) if available, otherwise fall back to cscript
where pwsh >nul 2>nul
if %ERRORLEVEL%==0 (
    echo Found pwsh - launching build.ps1 with PowerShell 7
    pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0build.ps1"
    exit /b %ERRORLEVEL%
) else (
    echo pwsh not found - falling back to cscript
    cscript //nologo "%~dp0build.vbs"
)

echo.
pause
