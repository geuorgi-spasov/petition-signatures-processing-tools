@echo off
rem Turns the raw CSV of signatures into Word files and then PDFs.
rem Double-click this file.
cd /d "%~dp0"

python --version >nul 2>&1
if errorlevel 1 goto no_python

echo Installing what the script needs...
python -m pip install --quiet --requirement requirements.txt
if errorlevel 1 goto failed

python split_signatures_into_folders.py
if errorlevel 1 goto failed

python convert_docx_to_pdf.py %*
if errorlevel 1 goto failed

echo.
echo Finished. Press any key to close this window.
pause >nul
exit /b 0

:no_python
echo.
echo Python is not installed, or it was installed without
echo "Add Python to PATH" ticked.
echo Install it from https://www.python.org/downloads/ and try again.
echo.
pause >nul
exit /b 1

:failed
echo.
echo Something went wrong - the message above says what.
echo.
pause >nul
exit /b 1
