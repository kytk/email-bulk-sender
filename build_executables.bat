@echo off
REM Email Bulk Sender Tool - Build Script for Windows

echo === Email Bulk Sender Build Script ===
echo.

REM Check and install required packages
echo Checking required packages...
pip list | findstr /C:"pyinstaller" >nul
if errorlevel 1 (
    echo PyInstaller is not installed. Installing...
    pip install pyinstaller
)

pip list | findstr /C:"customtkinter" >nul
if errorlevel 1 (
    echo CustomTkinter is not installed. Installing...
    pip install customtkinter
)

pip list | findstr /C:"chardet" >nul
if errorlevel 1 (
    echo chardet is not installed. Installing...
    pip install chardet
)

pip list | findstr /C:"openpyxl" >nul
if errorlevel 1 (
    echo openpyxl is not installed. Installing...
    pip install openpyxl
)

echo.
echo Starting build process...
echo.

REM Clean up existing build files
if exist build (
    echo Removing existing build directory...
    rmdir /s /q build
)

if exist dist (
    echo Removing existing dist directory...
    rmdir /s /q dist
)

if exist EmailBulkSender_win.spec (
    del EmailBulkSender_win.spec
)

if exist GmailBulkSender_win.spec (
    del GmailBulkSender_win.spec
)

echo.
echo 1/2: Building EmailBulkSender_win (Generic version)...
pyinstaller --onefile --windowed --name="EmailBulkSender_win" email_bulk_sender_gui.py

if errorlevel 1 (
    echo [FAILED] EmailBulkSender_win build failed
    pause
    exit /b 1
)
echo [SUCCESS] EmailBulkSender_win build completed

echo.
echo 2/2: Building GmailBulkSender_win (Gmail version)...
pyinstaller --onefile --windowed --name="GmailBulkSender_win" gmail_bulk_sender_gui.py

if errorlevel 1 (
    echo [FAILED] GmailBulkSender_win build failed
    pause
    exit /b 1
)
echo [SUCCESS] GmailBulkSender_win build completed

echo.
echo === Build Complete ===
echo.
echo Executables created in dist\ directory:
dir dist\*.exe
echo.

REM Create distribution packages with samples
echo.
echo Creating distribution packages...
echo.

REM Create temporary directories for packaging
if exist dist\EmailBulkSender_win_package rmdir /s /q dist\EmailBulkSender_win_package
if exist dist\GmailBulkSender_win_package rmdir /s /q dist\GmailBulkSender_win_package
mkdir dist\EmailBulkSender_win_package
mkdir dist\GmailBulkSender_win_package

REM Copy files for EmailBulkSender package
echo Copying files for EmailBulkSender package...
copy /Y dist\EmailBulkSender_win.exe dist\EmailBulkSender_win_package\
if errorlevel 1 (
    echo [ERROR] Failed to copy EmailBulkSender_win.exe
    pause
    exit /b 1
)
xcopy examples dist\EmailBulkSender_win_package\examples\ /E /I /Y
copy /Y README.md dist\EmailBulkSender_win_package\
copy /Y LICENSE dist\EmailBulkSender_win_package\

REM Copy files for GmailBulkSender package
echo Copying files for GmailBulkSender package...
copy /Y dist\GmailBulkSender_win.exe dist\GmailBulkSender_win_package\
if errorlevel 1 (
    echo [ERROR] Failed to copy GmailBulkSender_win.exe
    pause
    exit /b 1
)
xcopy examples dist\GmailBulkSender_win_package\examples\ /E /I /Y
copy /Y README.md dist\GmailBulkSender_win_package\
copy /Y LICENSE dist\GmailBulkSender_win_package\

REM Verify package contents
echo.
echo Verifying EmailBulkSender_win_package contents:
dir dist\EmailBulkSender_win_package
echo.
echo Verifying GmailBulkSender_win_package contents:
dir dist\GmailBulkSender_win_package
echo.

REM Wait a moment for file operations to complete
echo Waiting for file operations to complete...
timeout /t 3 /nobreak >nul

REM Create zip files (using PowerShell)
echo Creating EmailBulkSender_win.zip...
powershell -command "& {$ErrorActionPreference='Stop'; Compress-Archive -Path 'dist\EmailBulkSender_win_package\*' -DestinationPath 'dist\EmailBulkSender_win.zip' -Force}"
if errorlevel 1 (
    echo [ERROR] Failed to create EmailBulkSender_win.zip
    pause
    exit /b 1
)

echo Creating GmailBulkSender_win.zip...
powershell -command "& {$ErrorActionPreference='Stop'; Compress-Archive -Path 'dist\GmailBulkSender_win_package\*' -DestinationPath 'dist\GmailBulkSender_win.zip' -Force}"
if errorlevel 1 (
    echo [ERROR] Failed to create GmailBulkSender_win.zip
    pause
    exit /b 1
)

REM Clean up temporary directories
rmdir /s /q dist\EmailBulkSender_win_package
rmdir /s /q dist\GmailBulkSender_win_package

echo.
echo === Packaging Complete ===
echo.
echo Created distribution packages:
dir dist\*.zip
echo.
echo Package contents:
echo - Executable file (.exe)
echo - Sample files (examples/)
echo - README.md
echo - LICENSE
echo.
echo Windows executable files (.exe) and distribution packages (.zip) have been created.
echo You can distribute the .zip files to users without Python.
echo.
pause
