#!/bin/bash
# ==========================================
# ソースコード配布用パッケージ作成スクリプト
# メール一括送信ツール (Email Bulk Sender)
#
# 使い方:
#   cd email-bulk-sender
#   bash installer/build_source.sh
#
# Python 環境があれば OS を問わず利用できる
# ソースコード配布用の ZIP を生成します。
# ==========================================

set -e

# 必要ファイルの存在確認
for f in email_bulk_sender.py email_bulk_sender_gui.py requirements.txt README_source.md; do
    if [ ! -f "$f" ]; then
        echo "エラー: $f が見つかりません。プロジェクトのルートディレクトリで実行してください。"
        exit 1
    fi
done

if [ ! -d "examples" ]; then
    echo "エラー: examples/ ディレクトリが見つかりません。"
    exit 1
fi

# markdown モジュールの確認
if ! python3 -c "import markdown" 2>/dev/null; then
    echo "markdown モジュールをインストールします..."
    pip install markdown
fi

OUTPUT_DIR="installer/Output"
ZIP_NAME="EmailBulkSender_source.zip"
DIST_DIR="dist/EmailBulkSender_source"

# 出力ディレクトリの作成
mkdir -p "${OUTPUT_DIR}"

# 一時ディレクトリの準備
rm -rf "${DIST_DIR}"
mkdir -p "${DIST_DIR}"

# ファイルのコピー
echo "--- ファイルをコピー中 ---"
cp email_bulk_sender.py "${DIST_DIR}/"
cp email_bulk_sender_gui.py "${DIST_DIR}/"
cp requirements.txt "${DIST_DIR}/"
cp -r examples "${DIST_DIR}/examples"

# README_source.md を HTML に変換
echo "--- README.html を生成中 ---"
python3 -c "
import markdown
with open('README_source.md', encoding='utf-8') as f:
    md = f.read()
html = markdown.markdown(md, extensions=['tables', 'fenced_code'])
with open('${DIST_DIR}/README.html', 'w', encoding='utf-8') as f:
    f.write('<!DOCTYPE html><html><head><meta charset=\"utf-8\"><title>メール一括送信ツール</title></head><body>\n')
    f.write(html)
    f.write('\n</body></html>')
"

# ZIP を作成
echo "--- ZIP を作成中 ---"
rm -f "${OUTPUT_DIR}/${ZIP_NAME}"
(cd dist && zip -r "../${OUTPUT_DIR}/${ZIP_NAME}" "EmailBulkSender_source")

# クリーンアップ
rm -rf "${DIST_DIR}"

# 完了
echo ""
echo "=========================================="
echo "パッケージ作成完了!"
echo "  配布用ZIP: ${OUTPUT_DIR}/${ZIP_NAME}"
echo ""
echo "  ZIP の内容:"
echo "    - email_bulk_sender.py （CLI版）"
echo "    - email_bulk_sender_gui.py （GUI版）"
echo "    - requirements.txt"
echo "    - examples/ （サンプルファイル）"
echo "    - README.html"
echo "=========================================="
