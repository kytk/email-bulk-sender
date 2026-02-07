@echo off
REM ==========================================
REM Windows ビルドスクリプト
REM メール一括送信ツール (Email Bulk Sender)
REM
REM 使い方:
REM   cd email-bulk-sender
REM   installer\build_win.bat
REM
REM PyInstaller で exe を生成し、Inno Setup でインストーラーを作成します。
REM ==========================================

setlocal

set APP_NAME=EmailBulkSender
set SCRIPT=email_bulk_sender_gui.py

REM スクリプトの存在確認
if not exist "%SCRIPT%" (
    echo エラー: %SCRIPT% が見つかりません。プロジェクトのルートディレクトリで実行してください。
    exit /b 1
)

REM 依存パッケージのインストール
echo.
echo --- 依存パッケージを確認中 ---
pip install -r requirements.txt
pip install pyinstaller

REM 既存のビルドをクリーン
echo.
echo --- 既存のビルドをクリーン中 ---
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

REM PyInstaller でビルド
echo.
echo --- PyInstaller でビルド中 ---
pyinstaller --onedir --windowed --name %APP_NAME% --collect-data customtkinter %SCRIPT%

REM exe の存在確認
if not exist "dist\%APP_NAME%\%APP_NAME%.exe" (
    echo エラー: dist\%APP_NAME%\%APP_NAME%.exe が生成されませんでした。
    exit /b 1
)

echo.
echo --- exe の作成完了 ---
echo   dist\%APP_NAME%\%APP_NAME%.exe

REM Inno Setup でインストーラーを作成
echo.
echo --- Inno Setup でインストーラーを作成中 ---

where iscc >nul 2>nul
if %errorlevel% equ 0 (
    iscc installer\setup.iss
) else (
    echo.
    echo 注意: iscc コマンドが見つかりません。
    echo Inno Setup をインストールして PATH に追加するか、
    echo Inno Setup の GUI で installer\setup.iss を開いてコンパイルしてください。
    echo.
    echo ダウンロード: https://jrsoftware.org/isdl.php
    echo.
    echo ==========================================
    echo exe のビルドは完了しています。
    echo   dist\%APP_NAME%\%APP_NAME%.exe
    echo ==========================================
    goto :end
)

REM 配布用 ZIP の作成
echo.
echo --- 配布用 ZIP を作成中 ---

set ZIP_NAME=EmailBulkSender_win.zip
if exist "%ZIP_NAME%" del "%ZIP_NAME%"

REM 一時ディレクトリに配布物をまとめる
set DIST_DIR=dist\EmailBulkSender_win
if exist "%DIST_DIR%" rmdir /s /q "%DIST_DIR%"
mkdir "%DIST_DIR%"

copy installer\Output\EmailBulkSender_Setup.exe "%DIST_DIR%\"
xcopy examples "%DIST_DIR%\examples\" /s /e /i
copy README.md "%DIST_DIR%\"

REM PowerShell で ZIP を作成
powershell -Command "Compress-Archive -Path '%DIST_DIR%\*' -DestinationPath '%ZIP_NAME%' -Force"

echo.
echo ==========================================
echo ビルド完了!
echo   exe: dist\%APP_NAME%\%APP_NAME%.exe
echo   インストーラー: installer\Output\EmailBulkSender_Setup.exe
echo   配布用ZIP: %ZIP_NAME%
echo.
echo   ZIP の内容:
echo     - EmailBulkSender_Setup.exe （インストーラー）
echo     - examples/ （サンプルファイル）
echo     - README.md
echo ==========================================

:end
endlocal
