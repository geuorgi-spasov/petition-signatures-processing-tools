@echo off
rem Builds the book of initials.
rem Double-click this file.

rem pushd, not cd: it also works when the folder is on a network drive.
pushd "%~dp0"

if not exist "requirements.txt" goto no_files

python --version >nul 2>&1
if errorlevel 1 goto no_python

echo Installing what the program needs...
python -m pip install --quiet --requirement requirements.txt
if errorlevel 1 goto failed

python generate_initials_book.py %*
if errorlevel 1 goto failed

echo.
echo Finished. Press any key to close this window.
pause >nul
popd
exit /b 0

:no_files
echo.
echo This file cannot find the rest of the program.
echo It looked in this folder:
echo     %CD%
echo.
echo If you double-clicked it from inside the ZIP file, that is the cause.
echo Close the ZIP, right-click it, choose "Extract All", and then run
echo this file from the extracted folder instead.
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
