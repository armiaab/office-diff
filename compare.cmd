@echo off
if "%~2"=="" goto usage
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0office-compare.ps1" -ByDate "%~1" "%~2"
if errorlevel 1 pause
exit /b

:usage
echo Usage: compare.cmd ^<file1^> ^<file2^>
echo    or: drag two files onto compare.cmd
pause
exit /b 1
