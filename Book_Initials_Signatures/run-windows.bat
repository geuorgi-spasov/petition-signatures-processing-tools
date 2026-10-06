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

rem The font is not a Python dependency: the script only writes its *name*
rem into the document, and whatever opens the document has to have it. This
rem installs it for the current user only - %LOCALAPPDATA% plus an HKCU
rem registry entry - which needs no administrator rights.
set "FONT_FILE=oswald.ttf"
set "FONT_REG=HKCU\Software\Microsoft\Windows NT\CurrentVersion\Fonts"
set "FONT_DIR=%LOCALAPPDATA%\Microsoft\Windows\Fonts"

if not exist "%FONT_FILE%" goto font_done
reg query "%FONT_REG%" /v "Oswald (TrueType)" >nul 2>&1
if not errorlevel 1 goto font_done
if exist "%WINDIR%\Fonts\%FONT_FILE%" goto font_done

echo.
echo The book is set in Oswald, which is not installed yet.
echo The font is bundled here and free to install (SIL Open Font
echo License, see OFL.txt). It is installed for you only, so no
echo administrator rights are needed.
echo.
set "ANSWER="
set /p "ANSWER=Install it now? [Y/n] "
if /i "%ANSWER%"=="n" goto font_skipped
if /i "%ANSWER%"=="no" goto font_skipped

if not exist "%FONT_DIR%" mkdir "%FONT_DIR%" >nul 2>&1
copy /y "%FONT_FILE%" "%FONT_DIR%\%FONT_FILE%" >nul
if errorlevel 1 goto font_failed
reg add "%FONT_REG%" /v "Oswald (TrueType)" /t REG_SZ /d "%FONT_DIR%\%FONT_FILE%" /f >nul
if errorlevel 1 goto font_failed
echo Installed.
echo If Word is already open, close and reopen it so it sees the font.
goto font_done

:font_failed
echo.
echo Could not install the font automatically. Double-click oswald.ttf
echo in this folder and press Install, then run this file again.
echo The book will still be built below.
goto font_done

:font_skipped
echo.
echo Skipped. The book will still be built, but Word will substitute
echo another font and the spacing will be wrong.

:font_done
echo.

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
