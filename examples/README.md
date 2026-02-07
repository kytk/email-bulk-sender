# サンプルファイル / Sample Files

このディレクトリには、メール一括送信ツールを試すためのサンプルファイルが含まれています。

This directory contains sample files for testing the Email Bulk Sender tool.

## ファイル一覧 / File List

### list.csv.sample
サンプル受信者リスト

Sample recipient list

**使い方 / Usage:**
```bash
cp examples/list.csv.sample list.csv
# 必要に応じて list.csv を編集 / Edit list.csv as needed
```

**フォーマット / Format:**
```csv
所属,氏名,メールアドレス
Affiliation Name,Person Name,email@example.com
```

### body.txt.sample
日本語メールテンプレートのサンプル

Sample email template in Japanese

**使い方 / Usage:**
```bash
cp examples/body.txt.sample body.txt
# 必要に応じて body.txt を編集 / Edit body.txt as needed
```

**フォーマット / Format:**
```
件名（1行目） / Subject (Line 1)
[空行] / [Empty line]
本文（3行目以降） / Body (Line 3 onwards)
```

**プレースホルダー / Placeholders:**
- `{所属}` - 所属名に置換されます / Replaced with affiliation name
- `{氏名}` - 氏名に置換されます / Replaced with person name

### body_en.txt.sample
英語メールテンプレートのサンプル

Sample email template in English

**使い方 / Usage:**
```bash
cp examples/body_en.txt.sample body.txt
# 必要に応じて body.txt を編集 / Edit body.txt as needed
```

## 注意事項 / Important Notes

⚠️ これらはサンプルファイルです。実際に使用する前に、以下の点にご注意ください：

⚠️ These are sample files. Please note the following before actual use:

1. **メールアドレス / Email Addresses**
   - サンプルには `@example.com` ドメインを使用しています
   - 実際のメールアドレスに置き換えてください
   - Sample uses `@example.com` domain
   - Replace with actual email addresses

2. **テンプレートの編集 / Edit Templates**
   - 件名と本文を実際の用途に合わせて編集してください
   - Edit subject and body according to your actual use case

3. **テスト送信 / Test Sending**
   - 本番前に必ず自分宛てにテスト送信してください
   - Always send a test email to yourself before production use

4. **個人情報の保護 / Protect Personal Information**
   - 実際のリストファイルはGitにコミットしないでください
   - `.gitignore` により、`list.csv` と `body.txt` は除外できます（必要に応じてコメント解除）
   - Do not commit actual list files to Git
   - `.gitignore` can exclude `list.csv` and `body.txt` (uncomment if needed)

## クイックスタート / Quick Start

```bash
# 1. サンプルファイルをコピー / Copy sample files
cp examples/list.csv.sample list.csv
cp examples/body.txt.sample body.txt

# 2. ファイルを編集 / Edit files
# list.csv - 受信者リストを編集 / Edit recipient list
# body.txt - メールテンプレートを編集 / Edit email template

# 3. CLI版を実行（日本語） / Run CLI version (Japanese)
python email_bulk_sender.py --lang ja
# または / or
python gmail_bulk_sender.py --lang ja

# 4. GUI版を実行 / Run GUI version
python email_bulk_sender_gui.py
# または / or
python gmail_bulk_sender_gui.py
```

## テスト用の設定例 / Example Test Configuration

自分宛てにテスト送信する場合：

For sending a test email to yourself:

**list.csv:**
```csv
所属,氏名,メールアドレス
テスト株式会社,テスト太郎,your-email@example.com
```

**body.txt:**
```
テスト送信 - {所属} {氏名}様

{所属}
{氏名}様

これはテスト送信です。
プレースホルダーが正しく置換されているか確認してください。

所属名: {所属}
氏名: {氏名}
```

このテンプレートを使用すると、プレースホルダーが正しく動作しているか確認できます。

This template allows you to verify that placeholders are working correctly.
