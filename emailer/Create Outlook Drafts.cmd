@echo off
setlocal
cd /d "%~dp0"
title Outlook Draft Creator

echo.
echo Outlook Draft Creator (Guided)
echo ------------------------------
echo A few folder pickers will open.
echo - Select your exported outlook-draft-jobs.json file
echo - Select the Excel/attachment folders the app asks for
echo - Choose Dry Run or Create Drafts
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Create-OutlookDrafts.ps1" -Interactive
set "EXITCODE=%ERRORLEVEL%"

echo.
if not "%EXITCODE%"=="0" (
  echo The script finished with an error. Exit code: %EXITCODE%
) else (
  echo Done.
)
echo.
pause
exit /b %EXITCODE%
