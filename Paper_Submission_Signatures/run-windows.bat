@echo off
rem Turns the raw CSV of signatures into Word files and then PDFs.
rem Double-click this file.

rem pushd, not cd: it also works when the folder is on a network drive.
pushd "%~dp0"

if not exist "requirements.txt" goto no_files

python --version >nul 2>&1
if errorlevel 1 goto no_python

echo Installing what the program needs...
python -m pip install --quiet --requirement requirements.txt
if errorlevel 1 goto failed

rem Arguments go to the converter (--word, --libreoffice). The
rem splitting step has options of its own - run it directly:
rem     python split_signatures_into_folders.py --help
python split_signatures_into_folders.py
if errorlevel 1 goto failed

python convert_docx_to_pdf.py %*
if errorlevel 1 goto failed

echo.
echo Finished. Press any key to close this window.
pause >nul
popd
exit /b 0

:no_files
echo.
echo This file cannot find the rest of the program.
echo.
echo   This file is in:   %~dp0
echo   Looking in:        %CD%
echo.
echo requirements.txt should sit next to this file. If the two folders
echo above are different, tell Georgi - that is the problem.
echo If they are the same, the folder is missing its other files: close
echo the ZIP, right-click it, choose "Extract All", and run this file
echo from the extracted folder.
echo.
pause >nul
popd
exit /b 1

:no_python
echo.
echo Python is not installed, or it was installed without
echo "Add Python to PATH" ticked.
echo Install it from https://www.python.org/downloads/ and try again.
echo.
pause >nul
popd
exit /b 1

:failed
echo.
echo Something went wrong - the message above says what.
echo.
pause >nul
popd
exit /b 1
