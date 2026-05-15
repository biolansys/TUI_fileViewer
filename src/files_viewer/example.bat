@echo off
setlocal enabledelayedexpansion

set "APP_DIR=%~dp0"
set "LOG_FILE=%APP_DIR%example.log"

echo [%date% %time%] INFO starting demo>>"%LOG_FILE%"

echo Files in %APP_DIR%:
for %%F in ("%APP_DIR%*") do (
  if exist "%%~fF" echo - %%~nxF
)

echo Done.
endlocal
