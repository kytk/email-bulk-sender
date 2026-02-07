#!/bin/bash
# ==========================================
# macOS ビルドスクリプト
# メール一括送信ツール (Email Bulk Sender)
#
# 使い方:
#   cd email-bulk-sender
#   bash installer/build_mac.sh
#
# Intel Mac と Apple Silicon Mac の両方で使用可能。
# 実行した Mac のアーキテクチャに合わせた DMG を生成します。
# ==========================================

set -e

APP_NAME="メール一括送信ツール"
SCRIPT="email_bulk_sender_gui.py"

# アーキテクチャ検出
ARCH=$(uname -m)
if [ "$ARCH" = "arm64" ]; then
    ARCH_LABEL="arm64"
    echo "=== Apple Silicon (arm64) 用のビルドを開始します ==="
else
    ARCH_LABEL="intel"
    echo "=== Intel (x86_64) 用のビルドを開始します ==="
fi

OUTPUT_DIR="installer/Output"
DMG_NAME="EmailBulkSender_mac_${ARCH_LABEL}.dmg"

# venv 名をアーキテクチャに応じて設定
if [ "$ARCH_LABEL" = "arm64" ]; then
    VENV_NAME="pybulk_maca"
else
    VENV_NAME="pybulk_maci"
fi

# スクリプトの存在確認
if [ ! -f "$SCRIPT" ]; then
    echo "エラー: $SCRIPT が見つかりません。プロジェクトのルートディレクトリで実行してください。"
    exit 1
fi

# Python 仮想環境の作成・有効化
echo ""
echo "--- Python 仮想環境 (${VENV_NAME}) を準備中 ---"
if [ ! -d "${VENV_NAME}" ]; then
    /usr/bin/python3 -m venv "${VENV_NAME}"
    echo "仮想環境を作成しました: ${VENV_NAME}"
fi
source "${VENV_NAME}/bin/activate"
echo "仮想環境を有効化しました: $(which python3)"

# 依存パッケージのインストール確認
echo ""
echo "--- 依存パッケージを確認中 ---"
pip install -r requirements.txt
pip install pyinstaller

# 出力ディレクトリの作成
mkdir -p "${OUTPUT_DIR}"

# 既存のビルドをクリーン
echo ""
echo "--- 既存のビルドをクリーン中 ---"
rm -rf build dist "${OUTPUT_DIR}/${DMG_NAME}"

# PyInstaller でビルド
echo ""
echo "--- PyInstaller でビルド中 ---"
pyinstaller --onedir --windowed \
    --name "${APP_NAME}" \
    --collect-data customtkinter \
    "${SCRIPT}"

# .app の存在確認
APP_PATH="dist/${APP_NAME}.app"
if [ ! -d "$APP_PATH" ]; then
    echo "エラー: ${APP_PATH} が生成されませんでした。"
    exit 1
fi

echo ""
echo "--- .app バンドルの作成完了 ---"

# DMG の作成
echo ""
echo "--- DMG を作成中 ---"

if ! command -v create-dmg &> /dev/null; then
    echo "create-dmg が見つかりません。インストールします..."
    brew install create-dmg
fi

create-dmg \
    --volname "${APP_NAME}" \
    --window-size 600 400 \
    --icon-size 128 \
    --app-drop-link 400 200 \
    --icon "${APP_NAME}.app" 200 200 \
    "${OUTPUT_DIR}/${DMG_NAME}" \
    "${APP_PATH}"

# 配布用 ZIP の作成
echo ""
echo "--- 配布用 ZIP を作成中 ---"

ZIP_NAME="EmailBulkSender_mac_${ARCH_LABEL}.zip"
rm -f "${OUTPUT_DIR}/${ZIP_NAME}"

# 一時ディレクトリに配布物をまとめる
DIST_DIR="dist/EmailBulkSender_mac_${ARCH_LABEL}"
rm -rf "${DIST_DIR}"
mkdir -p "${DIST_DIR}"

cp "${OUTPUT_DIR}/${DMG_NAME}" "${DIST_DIR}/"
cp -r examples "${DIST_DIR}/examples"
cp README.md "${DIST_DIR}/"

# ZIP を作成
(cd dist && zip -r "../${OUTPUT_DIR}/${ZIP_NAME}" "EmailBulkSender_mac_${ARCH_LABEL}")

# 完了
echo ""
echo "=========================================="
echo "ビルド完了!"
echo "  DMG: ${OUTPUT_DIR}/${DMG_NAME}"
echo "  配布用ZIP: ${OUTPUT_DIR}/${ZIP_NAME}"
echo "  アーキテクチャ: ${ARCH} (${ARCH_LABEL})"
echo ""
echo "  ZIP の内容:"
echo "    - ${DMG_NAME} （インストーラー）"
echo "    - examples/ （サンプルファイル）"
echo "    - README.md"
echo "=========================================="
