@echo off
rem Compare two Word / PowerPoint / Excel files without git.
rem   - drag two files onto this file, or
rem   - select two files > right-click > Send to > office-diff compare (see README), or
rem   - compare.cmd old.docx new.docx
rem The file with the older modified date is treated as the old version.
if "%~2"=="" goto usage
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0office-compare.ps1" -ByDate "%~1" "%~2"
if errorlevel 1 pause
exit /b

:usage
echo Usage: compare.cmd ^<file1^> ^<file2^>
echo    or: drag two files onto compare.cmd
pause
exit /b 1
