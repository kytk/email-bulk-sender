# メール一括送信ツール

CSVやExcelでのリストに基づいた個別対応のメールを一斉送信するツールです。コマンドライン版（CLI）とGUI版の両方を提供しています。

## ツールの種類

### CLI版（コマンドライン）

1. **email_bulk_sender.py** - 汎用版
   - 任意のSMTPサーバーに対応（Gmail、Outlook、独自メールサーバーなど）
   - SMTPサーバー、ポート、認証情報を自由に設定可能
   - SSL/TLS両方に対応
   - 送信元表示名のカスタマイズが可能
   - 設定ファイルでデフォルト値を事前設定できる

2. **gmail_bulk_sender.py** - Gmail専用版
   - Gmail SMTPサーバーに特化
   - Gmailアプリパスワードを使用
   - 設定が簡単で、Gmailユーザーにおすすめ

### GUI版（グラフィカルユーザーインターフェース）

1. **email_bulk_sender_gui.py** - 汎用版GUI
   - 視覚的に操作できるGUIアプリケーション
   - ファイル選択ダイアログで簡単にファイルを指定
   - リアルタイムの進捗表示とログ表示
   - 送信前のプレビュー機能
   - コマンドラインに不慣れな方におすすめ

2. **gmail_bulk_sender_gui.py** - Gmail専用版GUI
   - Gmail専用のGUIアプリケーション
   - 設定項目が少なく、より簡単に使用可能

### スタンドアロン実行ファイル（.exe / バイナリ）

プラットフォーム別に以下の実行ファイルを作成できます：

- **Windows**: `EmailBulkSender_win.exe`, `GmailBulkSender_win.exe`
- **Linux**: `EmailBulkSender_lnx`, `GmailBulkSender_lnx`
- **macOS (ARM64)**: `EmailBulkSender_mac_arm64`, `GmailBulkSender_mac_arm64`
- **macOS (AMD64)**: `EmailBulkSender_mac_amd64`, `GmailBulkSender_mac_amd64`

Pythonがインストールされていない環境でも実行可能です。詳細は[BUILD_GUIDE.md](BUILD_GUIDE.md)を参照してください。

**どれを使うべきか：**
- **初めて使う / GUIが好き** → GUI版（`email_bulk_sender_gui.py` または `gmail_bulk_sender_gui.py`）
- **コマンドラインに慣れている** → CLI版（`email_bulk_sender.py` または `gmail_bulk_sender.py`）
- **Pythonがない環境で使いたい** → スタンドアロン実行ファイル（dist/内のプラットフォーム別実行ファイル）
- **Gmailのみ使用** → Gmail専用版
- **Gmail以外のメールサーバー使用** → 汎用版

以下の説明は主にCLI版に関するものです。GUI版の使い方は[GUI版の使い方](#gui版の使い方)をご覧ください。

## 目次

- [ツールの種類](#ツールの種類)
- [機能](#機能)
- [必要な環境](#必要な環境)
- [セットアップ](#セットアップ)
- [ファイルの準備](#ファイルの準備)
- [使い方（CLI版）](#使い方cli版)
- [設定ファイルの使用](#設定ファイルの使用)
- [トラブルシューティング](#トラブルシューティング)
- [注意事項](#注意事項)
- [よくある質問](#よくある質問)
- [汎用版(email_bulk_sender.py)](#汎用版email_bulk_senderpy)
- [GUI版の使い方](#gui版の使い方)
- [スタンドアロン実行ファイルの作成](#スタンドアロン実行ファイルの作成)
- [ライセンス](#ライセンス)
- [サポート](#サポート)

## 機能

### 共通機能
- **CSVおよびExcelファイル（.xlsx）対応** - 受信者リスト（企業名、氏名、メールアドレス）を読み込み
- **改行コードの自動処理** - Windows/Unix/Mac間でのファイル交換に対応
- テンプレート内の`{企業}`と`{氏名}`を各受信者の情報に自動置換
- 件名と本文を1つのテンプレートファイルで管理
- CC、BCC、Reply-To の設定に対応
- ファイルの添付に対応
- 文字コード自動検出（UTF-8、Shift_JIS、EUC-JPなど）
- 送信制限対策（送信間隔の自動調整）
- 送信結果のリアルタイム表示
- **多言語対応**（日本語・英語の切り替え可能）
- **設定ファイル対応**（JSON形式で設定を保存・読み込み可能）

### GUI版の追加機能
- 視覚的で直感的な操作インターフェース
- ファイル選択ダイアログ
- 送信前のプレビュー機能
- リアルタイムプログレスバー
- 詳細な送信ログ表示
- エラー時のわかりやすいメッセージ表示

## 必要な環境

- Python 3.6以上
- Gmailアカウント（2段階認証が有効であること）
- インターネット接続

### 必要なPythonライブラリ

#### CLI版（基本機能）
標準ライブラリのみ使用するため、追加インストールは不要です。

#### Excelファイル対応（オプション）
Excelファイル（.xlsx）を使用する場合は、以下をインストール：
```bash
pip install openpyxl
```

#### GUI版
GUI版を使用する場合は、以下をインストール：
```bash
pip install customtkinter chardet openpyxl
```

## セットアップ

### 1. スクリプトのダウンロード

```bash
# リポジトリをクローン、またはファイルをダウンロード
git clone https://github.com/kytk/email-bulk-sender.git
cd email-bulk-sender

# または、bulk_email_sender.py を直接ダウンロード
```

### 2. 仮想環境のセットアップ（推奨）

仮想環境を使用することで、システム全体に影響を与えずにパッケージを管理できます。

#### 仮想環境の作成と有効化

**Unix/Mac/Linux:**
```bash
# 仮想環境を作成
python3 -m venv venv_email

# 仮想環境を有効化
source venv_email/bin/activate
```

**Windows:**
```bash
# 仮想環境を作成
python -m venv venv_email

# 仮想環境を有効化（コマンドプロンプト）
venv_email\Scripts\activate

# 仮想環境を有効化（PowerShell）
# PowerShellで初めて実行する場合、実行ポリシーを設定する必要があります
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
venv_email\Scripts\Activate.ps1
```

#### 必要なパッケージのインストール

仮想環境を有効化した後、必要なパッケージをインストール：

**GUI版を使用する場合:**
```bash
pip install customtkinter chardet openpyxl
```

**Excelファイル対応のみ（CLI版）:**
```bash
pip install openpyxl
```

**注:** CLI版でCSVファイルのみ使用する場合は、追加パッケージのインストールは不要です。

#### 仮想環境の終了

作業が終わったら、仮想環境を終了：
```bash
deactivate
```

### 3. Gmailアプリパスワードの取得

**重要**: 通常のGmailパスワードではなく、「アプリパスワード」を使用する必要があります。

#### 手順：

1. **Googleアカウントにアクセス**
   - [https://myaccount.google.com/](https://myaccount.google.com/) にアクセス
   - Gmailアカウントでログイン

2. **2段階認証を有効化**（まだの場合）
   - 左メニューから「セキュリティ」を選択
   - 「Google へのログイン」セクションで「2段階認証プロセス」をクリック
   - 画面の指示に従って設定

3. **アプリパスワードを生成**
   - 「セキュリティ」ページに戻る
   - 上の検索窓に「アプリ パスワード」と入力
   - 「アプリ固有のパスワードを新規作成するには、下にアプリ名を入力してください」に任意のアプリ名を入力（例：「一斉送信メール」）
   - 「作成」をクリック

4. **パスワードをコピー**
   - 表示された16文字のパスワードをコピー
   - **このパスワードは再表示されないため、安全な場所に保存してください**

   例：`abcd efgh ijkl mnop`（スペースは入力時に不要）

### 3. ファイル構成の確認

プロジェクトフォルダに以下のファイルを配置します：

```
email-bulk-sender/
├── bulk_email_sender.py  # メインスクリプト
├── list.csv              # 受信者リスト
├── body.txt              # メールテンプレート
└── README.md             # このファイル
```

## ファイルの準備

**サンプルファイルの利用:**

`examples/` ディレクトリにサンプルファイルが用意されています。以下のコマンドでコピーしてすぐに試すことができます：

**Unix/Mac/Linux:**
```bash
cp examples/list_sample.csv list.csv
cp examples/body_sample.txt body.txt
# 必要に応じてファイルを編集してください
```

**Windows（コマンドプロンプト）:**
```bash
copy examples\list_sample.csv list.csv
copy examples\body_sample_win.txt body.txt
# 必要に応じてファイルを編集してください
```

**利用可能なサンプルファイル:**
- `list_sample.csv` - CSV形式の受信者リスト
- `list_sample.xlsx` - Excel形式の受信者リスト
- `body_sample.txt` - UTF-8、Unix改行のメールテンプレート
- `body_sample_win.txt` - Shift JIS、Windows改行（CR+LF）のメールテンプレート
- `body_en_sample.txt` - 英語版メールテンプレート

詳細は [examples/README.md](examples/README.md) を参照してください。

### 受信者リスト（list.csv または list.xlsx）

受信者の情報をCSVファイルまたはExcelファイル（.xlsx）で記載します。

#### CSVファイル形式（list.csv）

**フォーマット：**
```csv
企業,氏名,メールアドレス
株式会社ABC,山田太郎,yamada@example.com
XYZ商事株式会社,佐藤花子,sato@example.com
テクノロジー株式会社,鈴木一郎,suzuki@example.com
サンプル株式会社,田中美咲,tanaka@example.com
```

**注意点：**
- 1行目は必ずヘッダー行（`企業,氏名,メールアドレス`）
- 文字コードは UTF-8 で保存
- Excelで編集する場合、保存時に「CSV UTF-8（コンマ区切り）」を選択
- Windows/Unix/Mac間でのファイル交換も自動で対応

#### Excelファイル形式（list.xlsx）

**特徴：**
- Excelで直接編集・保存が可能（CSV変換不要）
- 改行コードの問題が発生しない
- 文字コードを気にせず使用可能

**フォーマット：**
| 企業 | 氏名 | メールアドレス |
|------|------|----------------|
| 株式会社ABC | 山田太郎 | yamada@example.com |
| XYZ商事株式会社 | 佐藤花子 | sato@example.com |

**注意点：**
- 1行目は必ずヘッダー行（`企業`、`氏名`、`メールアドレス`）
- アクティブシート（最初のシート）が読み込まれます
- openpyxlライブラリが必要です（`pip install openpyxl`）

### body.txt（メールテンプレート）

メールの件名と本文を記載します。

**フォーマット：**
```
【重要】新サービスのご案内 - {企業} {氏名}様

{企業}
{氏名}様

いつもお世話になっております。
株式会社サンプルの営業部です。

{企業}様におかれましては、ますますご清栄のこととお慶び申し上げます。

この度、新サービス「○○○」をリリースいたしましたので、
ご案内させていただきます。

■ サービス概要
- 特徴1：○○○
- 特徴2：△△△
- 特徴3：□□□

詳細につきましては、下記URLをご参照ください。
https://example.com/service

ご不明な点がございましたら、お気軽にお問い合わせください。

今後とも何卒よろしくお願い申し上げます。

────────────────────────
株式会社サンプル
営業部 山田太郎
Email: info@example.com
Tel: 03-1234-5678
────────────────────────
```

**ルール：**
- **1行目**: 件名（`{企業}`で企業名、`{氏名}`で受信者名を挿入可能）
- **2行目**: 空行（必須）
- **3行目以降**: 本文（`{企業}`で企業名、`{氏名}`で受信者名を挿入可能）

## 使い方

### 言語設定

すべてのプログラムは日本語・英語の両方に対応しています。

**CLI版（コマンドライン版）:**
```bash
# 日本語で実行（デフォルト、システム言語を自動検出）
python email_bulk_sender.py

# 英語で実行
python email_bulk_sender.py --lang en

# 日本語を明示的に指定
python email_bulk_sender.py --lang ja
```

**GUI版:**
- アプリケーション内の「言語設定」タブで言語を選択
- 選択した言語は設定ファイルに保存され、次回起動時に反映
- 言語を変更した場合は、アプリケーションを再起動してください

### 1. スクリプトの実行

```bash
python bulk_email_sender.py
```

### 2. 対話形式での入力

実行すると以下の項目を順番に入力します：

```
=== Gmail一斉送信ツール ===

送信元Gmailアドレス: your.email@gmail.com
Gmailアプリパスワード: ****************
受信者リストCSVファイル (デフォルト: list.csv): 
メールテンプレートファイル (デフォルト: body.txt): 
CC (複数の場合はカンマ区切り、不要ならEnter): cc@example.com
BCC (複数の場合はカンマ区切り、不要ならEnter): 
Reply-To (不要ならEnter): reply@example.com
```

**入力のポイント：**
- **Gmailアプリパスワード**: 取得した16文字のパスワード（スペースなしで入力）
- **ファイル名**: Enterキーで デフォルト値（list.csv、body.txt）を使用
- **CC/BCC**: 複数指定する場合は `email1@example.com,email2@example.com` のようにカンマ区切り
- **不要な項目**: そのままEnterキーを押す

### 3. 送信内容の確認

入力後、送信内容が表示されます：

```
=== 送信内容確認 ===
件名: 【重要】新サービスのご案内 - {企業} {氏名}様
送信先: 4件
送信元: your.email@gmail.com
CC: cc@example.com
Reply-To: reply@example.com

送信を開始しますか？ (yes/no):
```

### 4. 送信開始

`yes` と入力すると送信が開始されます：

```
[1/4] 送信成功: 株式会社ABC 山田太郎 (yamada@example.com)
[2/4] 送信成功: XYZ商事株式会社 佐藤花子 (sato@example.com)
[3/4] 送信成功: テクノロジー株式会社 鈴木一郎 (suzuki@example.com)
[4/4] 送信成功: サンプル株式会社 田中美咲 (tanaka@example.com)

送信完了: 成功 4件, 失敗 0件
```

## 設定ファイルの使用

### 概要

バージョン2.1から、設定をJSON形式のファイルに保存・読み込みできるようになりました。これにより、スクリプトを編集せずに設定を管理できます。

**主な特徴：**
- スクリプトの`DEFAULT_*`変数を編集する必要がなくなります
- GUI/CLIの両方から設定を保存・読み込み可能
- テキストエディタでJSONファイルを直接編集することも可能
- パスワードは保存されません（セキュリティのため）

### 設定ファイルの場所

- **汎用版**: `~/.email_bulk_sender/config.json`
- **Gmail版**: `~/.gmail_bulk_sender/config.json`

（`~`はユーザーのホームディレクトリを表します）

### CLI版での使用方法

#### 設定の保存

```bash
# 設定を入力して保存
python email_bulk_sender.py --save-config

# Gmail版の場合
python gmail_bulk_sender.py --save-config
```

初回実行時に設定を入力すると、設定ファイルに保存されます。

#### 設定の読み込み

CLI版では、`--load-config`フラグを指定した場合のみ設定ファイルを読み込みます：

```bash
# デフォルト動作：スクリプト内のDEFAULT_*変数を使用
python email_bulk_sender.py

# 設定ファイルから読み込む場合
python email_bulk_sender.py --load-config

# Gmail版の場合
python gmail_bulk_sender.py --load-config
```

**動作の違い：**
- **フラグなし**：スクリプト内の`DEFAULT_*`変数が使用されます。設定ファイルがあっても無視されます。
- **--load-config**：設定ファイルから読み込みます。設定ファイルに保存されていない項目（例：パスワード）のみ入力を求められます。

**注意**：GUI版は起動時に自動的に設定ファイルを読み込みます。

### GUI版での使用方法

GUI版では、基本設定タブに「設定管理」セクションがあります：

1. **設定を保存**
   - 現在入力されている設定を設定ファイルに保存
   - 「設定を保存」ボタンをクリック
   - パスワードは保存されません

2. **設定を読み込み**
   - 保存された設定を読み込む
   - 「設定を読み込み」ボタンをクリック
   - すべてのフィールドに設定が反映されます

### 設定ファイルの形式

設定ファイルはJSON形式です。テキストエディタで直接編集することもできます：

```json
{
  "version": "2.0",
  "smtp": {
    "server": "smtp.gmail.com",
    "port": 587
  },
  "sender": {
    "email_address": "your.email@gmail.com",
    "display_name": "株式会社サンプル 営業部"
  },
  "files": {
    "csv_file": "list.csv",
    "template_file": "body.txt",
    "attachments": ["file1.pdf", "file2.docx"]
  },
  "email_options": {
    "cc": "cc@example.com",
    "bcc": "",
    "reply_to": "reply@example.com",
    "send_delay": 5
  },
  "ui": {
    "language": "ja"
  }
}
```

**注意**: パスワードはセキュリティ上の理由から設定ファイルに保存されません。毎回入力する必要があります。

### 複数の設定を使い分ける

設定ファイルをコピーして、異なる用途で使い分けることができます：

```bash
# 設定ファイルをバックアップ
cp ~/.email_bulk_sender/config.json ~/.email_bulk_sender/config_work.json
cp ~/.email_bulk_sender/config.json ~/.email_bulk_sender/config_personal.json

# 必要に応じて設定を入れ替え
cp ~/.email_bulk_sender/config_work.json ~/.email_bulk_sender/config.json
```

## トラブルシューティング

### エラー: 認証失敗

```
SMTP接続エラー: (535, b'5.7.8 Username and Password not accepted...')
```

**原因と対処法：**
1. アプリパスワードが間違っている → 再度確認・再生成
2. 2段階認証が無効 → Googleアカウントで有効化
3. 通常のパスワードを使用している → アプリパスワードを使用

### エラー: ファイルが見つからない

```
FileNotFoundError: [Errno 2] No such file or directory: 'list.csv'
```

**対処法：**
- ファイル名とパスを確認
- スクリプトと同じフォルダにファイルがあるか確認
- ファイル名の大文字小文字を確認

### エラー: 送信制限

```
SMTPDataError: (550, b'5.4.5 Daily sending quota exceeded')
```

**対処法：**
- Gmail無料版: 1日500通まで
- Google Workspace: 1日2,000通まで
- 翌日まで待つか、複数のアカウントで分散送信

### 文字化け

**対処法：**
- CSVファイルをUTF-8で保存
- body.txtをUTF-8で保存
- Windowsの場合、メモ帳で開いて「UTF-8」を選択して保存

## 注意事項

### 送信制限

- **無料Gmail**: 1日あたり500通まで
- **Google Workspace**: 1日あたり2,000通まで
- 1時間あたりの制限もあるため、大量送信時は注意

### セキュリティ

- アプリパスワードをコードに直接書かない
- アプリパスワードをGitにコミットしない
- 使用後は安全に保管
- 不要になったアプリパスワードは削除

### ベストプラクティス

1. **テスト送信**: 本番前に自分宛てにテスト送信
2. **少数から開始**: 最初は5-10件程度で動作確認
3. **送信リストの確認**: 送信前に必ず内容を確認
4. **エラーログ**: 失敗した送信先は記録して再送
5. **配信停止対応**: オプトアウトのリンクを含める

### 法的注意事項

- 受信者の同意なく営業メールを送信しない
- 特定電子メール法を遵守
- オプトアウト（配信停止）の手段を提供
- 個人情報を適切に管理

## よくある質問

### Q: HTMLメールを送信できますか？

A: 現在のバージョンはテキストメールのみ対応。HTMLメールを送りたい場合は、`MIMEText`の第2引数を`'html'`に変更し、body.txtにHTMLを記述してください。

### Q: 添付ファイルを送れますか？

A: はい、対応しています。GUI版では「ファイル選択タブ」で添付ファイルを追加できます。CLI版でも添付ファイル機能が実装されています。

### Q: 送信間隔を変更できますか？

A: `send_bulk_emails`メソッドの`delay`パラメータ（デフォルト1秒）を変更できます。

### Q: 複数の変数（企業・氏名以外）を使えますか？

A: 現在、`{企業}`と`{氏名}`の2つのプレースホルダーが使用可能です。さらに追加したい場合は、CSVに列を追加し、スクリプトを修正すれば可能です。例えば`{役職}`なども使用できます。

## 汎用版(email_bulk_sender.py)

### 概要

`email_bulk_sender.py`は任意のSMTPサーバーに対応した汎用的なメール一斉送信スクリプトです。Gmail以外のメールサービス（Outlook、Yahoo、独自メールサーバーなど）や、送信元表示名のカスタマイズが必要な場合に使用します。

### Gmail専用版との主な違い

1. **対応SMTPサーバー**
   - Gmail専用版: Gmail SMTPのみ
   - 汎用版: 任意のSMTPサーバー

2. **設定方法**
   - Gmail専用版: 実行時にGmailアドレスとアプリパスワードを入力
   - 汎用版: スクリプト内の設定セクションで事前設定可能（実行時入力も可能）

3. **送信元表示名**
   - Gmail専用版: メールアドレスのみ
   - 汎用版: 任意の表示名を設定可能（例: "株式会社サンプル 営業部"）

4. **SSL/TLS対応**
   - Gmail専用版: TLS固定
   - 汎用版: ポート番号により自動判別（465=SSL、587=TLS）

### 使い方

#### 1. 設定ファイルの編集（オプション）

`email_bulk_sender.py`を開き、冒頭の設定セクションを編集します：

```python
# SMTPサーバー（空文字列の場合は実行時に入力を求めます）
DEFAULT_SMTP_SERVER = "smtp.example.com"  # 例: "smtp.gmail.com", "smtp.office365.com"

# SMTPポート番号（空文字列の場合は実行時に入力を求めます）
DEFAULT_SMTP_PORT = "587"  # 587 (TLS), 465 (SSL), 25 (非暗号化)

# メールアドレス（空文字列の場合は実行時に入力を求めます）
DEFAULT_EMAIL_ADDRESS = "your.email@example.com"

# メールパスワード（セキュリティ上、空文字列のまま実行時入力を推奨）
DEFAULT_EMAIL_PASSWORD = ""

# 送信元表示名（空文字列の場合はメールアドレスのみ表示）
SENDER_DISPLAY_NAME = "株式会社サンプル 営業部"

# 受信者リストCSVファイル
DEFAULT_CSV_FILE = "list.csv"

# メールテンプレートファイル
DEFAULT_TEMPLATE_FILE = "body.txt"

# CC, BCC, Reply-To（オプション）
DEFAULT_CC = ""
DEFAULT_BCC = ""
DEFAULT_REPLY_TO = ""
```

**設定のポイント：**
- 空文字列 `""` のままにした項目は実行時に入力を求められます
- パスワードは**セキュリティ上、空文字列のまま**にすることを推奨
- 設定した項目は実行時に「(設定済み)」と表示され、入力をスキップします

#### 2. ファイルの準備

`list.csv`と`body.txt`は`bulk_email_sender.py`と同じ形式です（[ファイルの準備](#ファイルの準備)参照）。

#### 3. スクリプトの実行

```bash
python email_bulk_sender.py
```

実行すると、設定していない項目について入力を求められます：

```
=== メール一斉送信ツール ===

SMTPサーバー (例: smtp.gmail.com): smtp.example.com
SMTPポート (デフォルト: 587): 587
送信元メールアドレス: your.email@example.com
メールパスワード: ********
送信元表示名 (不要ならEnter): 株式会社サンプル 営業部
受信者リストCSVファイル (デフォルト: list.csv):
メールテンプレートファイル (デフォルト: body.txt):
CC (複数の場合はカンマ区切り、不要ならEnter):
BCC (複数の場合はカンマ区切り、不要ならEnter):
Reply-To (不要ならEnter):
```

### 主要メールサービスのSMTP設定

| サービス | SMTPサーバー | ポート | セキュリティ |
|---------|-------------|-------|-------------|
| Gmail | smtp.gmail.com | 587 | TLS |
| Outlook/Office365 | smtp.office365.com | 587 | TLS |
| Yahoo Mail | smtp.mail.yahoo.com | 587 | TLS |
| iCloud | smtp.mail.me.com | 587 | TLS |

**注意：** Gmailを使用する場合、アプリパスワードの取得が必要です（[Gmailアプリパスワードの取得](#2-gmailアプリパスワードの取得)参照）。

### トラブルシューティング（汎用版特有）

#### ポート番号エラー

```
ConnectionRefusedError: [Errno 111] Connection refused
```

**対処法：**
- 正しいポート番号を確認（一般的には587または465）
- ファイアウォールがポートをブロックしていないか確認
- SMTPサーバーアドレスが正しいか確認

#### SSL/TLS エラー

```
ssl.SSLError: [SSL: WRONG_VERSION_NUMBER]
```

**対処法：**
- ポート465の場合はSSL、587の場合はTLSが使用されます
- ポート番号とサーバー設定の組み合わせを確認

## GUI版の使い方

### 必要なパッケージ

GUI版を使用する場合、追加のパッケージが必要です：

```bash
pip install customtkinter chardet openpyxl
```

注: `openpyxl`はExcelファイル（.xlsx）のサポートに必要です。CSVファイルのみ使用する場合は不要ですが、インストールを推奨します。

### 起動方法

#### 汎用版GUI
```bash
python email_bulk_sender_gui.py
```

#### Gmail専用版GUI
```bash
python gmail_bulk_sender_gui.py
```

### 操作手順

1. **基本設定タブ**
   - SMTPサーバー情報（汎用版）またはGmailアカウント情報（Gmail版）を入力
   - メールアドレスとパスワードを入力
   - 送信元表示名を入力（オプション）
   - **設定管理**: 「設定を保存」または「設定を読み込み」ボタンで設定を管理できます

2. **ファイル選択タブ**
   - 「ファイル選択」ボタンをクリックして、受信者リストCSVファイルを選択
   - 「ファイル選択」ボタンをクリックして、メールテンプレートファイルを選択
   - 添付ファイルがある場合は「ファイル追加」ボタンで追加

3. **オプション設定タブ**
   - CC、BCC、Reply-Toを入力（オプション）
   - 送信間隔（秒）を設定

4. **送信実行タブ**
   - 「送信内容をプレビュー」ボタンをクリックして内容を確認
   - 「メール送信開始」ボタンをクリックして送信開始
   - プログレスバーとログで進捗を確認

### GUI版の特徴

- **ファイル選択ダイアログ**: ファイルパスを手動で入力する必要がありません
- **プレビュー機能**: 送信前に内容を確認できます
- **リアルタイム進捗**: プログレスバーで送信状況を視覚的に確認
- **詳細ログ**: 送信成功/失敗の詳細をリアルタイムで表示
- **エラーハンドリング**: わかりやすいエラーメッセージを表示

## スタンドアロン実行ファイルの作成

Pythonがインストールされていない環境でも実行できるスタンドアロン実行ファイル（Windows: .exe、Linux/Mac: バイナリ）を作成できます。

ビルドスクリプトは自動的に以下を作成します：
- 実行ファイル単体
- **配布用zipパッケージ**（実行ファイル + サンプルファイル + README + LICENSE）

### 自動ビルド（推奨）

#### Windows
```cmd
build_executables.bat
```

作成されるファイル：
- `dist/EmailBulkSender_win.exe` - 汎用版実行ファイル
- `dist/EmailBulkSender_win.zip` - 汎用版配布パッケージ ⭐
- `dist/GmailBulkSender_win.exe` - Gmail版実行ファイル
- `dist/GmailBulkSender_win.zip` - Gmail版配布パッケージ ⭐

#### Linux/Mac
```bash
chmod +x build_executables.sh
./build_executables.sh
```

作成されるファイル（例：Linux）：
- `dist/EmailBulkSender_lnx` - 汎用版実行ファイル
- `dist/EmailBulkSender_lnx.zip` - 汎用版配布パッケージ ⭐
- `dist/GmailBulkSender_lnx` - Gmail版実行ファイル
- `dist/GmailBulkSender_lnx.zip` - Gmail版配布パッケージ ⭐

### 配布パッケージの内容

zipファイルには以下が含まれます：
```
EmailBulkSender_xxx.zip
├── EmailBulkSender_xxx[.exe]  # 実行ファイル
├── examples/                   # サンプルファイル
│   ├── list_sample.csv
│   ├── list_sample.xlsx
│   ├── body_sample.txt
│   ├── body_sample_win.txt    # Windows用（Shift JIS、CR+LF）
│   ├── body_en_sample.txt
│   └── README.md
├── README.md                   # 使い方ガイド
└── LICENSE                     # ライセンス
```

**配布方法：**
zipファイルを配布すれば、受け取った人は解凍してすぐに使い始められます：
1. zipファイルを解凍
2. サンプルファイルをコピー：
   - **Windows:** `examples\list_sample.csv` を `list.csv` にコピー、`examples\body_sample_win.txt` を `body.txt` にコピー
   - **Unix/Mac:** `examples/list_sample.csv` を `list.csv` にコピー、`examples/body_sample.txt` を `body.txt` にコピー
3. ファイルを編集して実行ファイルをダブルクリック

### 手動ビルド

1. PyInstallerをインストール：
   ```bash
   pip install pyinstaller customtkinter chardet openpyxl
   ```

2. 実行ファイルを作成：

   **Windows:**
   ```bash
   # 汎用版GUI
   pyinstaller --onefile --windowed --name="EmailBulkSender_win" email_bulk_sender_gui.py

   # Gmail専用版GUI
   pyinstaller --onefile --windowed --name="GmailBulkSender_win" gmail_bulk_sender_gui.py
   ```

   **Linux:**
   ```bash
   # 汎用版GUI
   pyinstaller --onefile --windowed --name="EmailBulkSender_lnx" email_bulk_sender_gui.py

   # Gmail専用版GUI
   pyinstaller --onefile --windowed --name="GmailBulkSender_lnx" gmail_bulk_sender_gui.py
   ```

   **macOS:**
   ```bash
   # ARM64版 (M1/M2/M3 Mac)
   pyinstaller --onefile --windowed --name="EmailBulkSender_mac_arm64" email_bulk_sender_gui.py
   pyinstaller --onefile --windowed --name="GmailBulkSender_mac_arm64" gmail_bulk_sender_gui.py

   # AMD64版 (Intel Mac)
   pyinstaller --onefile --windowed --name="EmailBulkSender_mac_amd64" email_bulk_sender_gui.py
   pyinstaller --onefile --windowed --name="GmailBulkSender_mac_amd64" gmail_bulk_sender_gui.py
   ```

3. 実行ファイルは `dist/` ディレクトリに作成されます

詳細は [BUILD_GUIDE.md](BUILD_GUIDE.md) を参照してください。

### 実行ファイルの配布

作成された実行ファイル（例：`dist/EmailBulkSender_win.exe`、`dist/EmailBulkSender_lnx`など）を配布すれば、Pythonがインストールされていない環境でも実行できます。

**注意事項：**
- Windows用の.exeファイルを作成するには、Windows環境でビルドする必要があります
- Linux用のバイナリを作成するには、Linux環境でビルドする必要があります
- Mac用のバイナリを作成するには、Mac環境でビルドする必要があります

## ライセンス

このプロジェクトはMITライセンスの下でライセンスされています。詳細は[LICENSE](LICENSE)ファイルを参照してください。

## サポート

問題や質問がある場合は、Issueを作成してください。

---

**作成日**: 2025年9月
**更新日**: 2025年11月
**バージョン**: 2.3（設定読み込みロジック改善）

---

# Email Bulk Sender Tool

A tool for sending personalized bulk emails based on CSV or Excel recipient lists. Provides both CLI (command line) and GUI versions.

## Tool Types

### CLI Version (Command Line)

1. **email_bulk_sender.py** - Generic version
   - Supports any SMTP server (Gmail, Outlook, custom mail servers, etc.)
   - Flexible SMTP server, port, and authentication settings
   - Supports both SSL/TLS
   - Customizable sender display name
   - Pre-configurable default values in settings section

2. **gmail_bulk_sender.py** - Gmail-specific version
   - Specialized for Gmail SMTP server
   - Uses Gmail App Password
   - Easy to set up, recommended for Gmail users

### GUI Version (Graphical User Interface)

1. **email_bulk_sender_gui.py** - Generic GUI version
   - Visual and intuitive GUI application
   - Easy file selection with file dialogs
   - Real-time progress display and logs
   - Email preview before sending
   - Recommended for users unfamiliar with command line

2. **gmail_bulk_sender_gui.py** - Gmail-specific GUI version
   - Gmail-dedicated GUI application
   - Fewer configuration items, easier to use

### Standalone Executables (.exe / Binary)

Platform-specific executables can be created:

- **Windows**: `EmailBulkSender_win.exe`, `GmailBulkSender_win.exe`
- **Linux**: `EmailBulkSender_lnx`, `GmailBulkSender_lnx`
- **macOS (ARM64)**: `EmailBulkSender_mac_arm64`, `GmailBulkSender_mac_arm64`
- **macOS (AMD64)**: `EmailBulkSender_mac_amd64`, `GmailBulkSender_mac_amd64`

Can run on systems without Python installed. See [BUILD_GUIDE.md](BUILD_GUIDE.md) for details.

**Which one to use:**
- **First time / Prefer GUI** → GUI version (`email_bulk_sender_gui.py` or `gmail_bulk_sender_gui.py`)
- **Comfortable with command line** → CLI version (`email_bulk_sender.py` or `gmail_bulk_sender.py`)
- **No Python environment** → Standalone executables (platform-specific files in dist/)
- **Gmail only** → Gmail-specific version
- **Non-Gmail mail servers** → Generic version

The following documentation primarily covers the CLI version. For GUI version usage, see [GUI Version Usage](#gui-version-usage).

## Table of Contents

- [Tool Types](#tool-types)
- [Features](#features)
- [Requirements](#requirements)
- [Setup](#setup)
- [File Preparation](#file-preparation)
- [Usage (CLI Version)](#usage-cli-version)
- [Using Configuration Files](#using-configuration-files)
- [GUI Version Usage](#gui-version-usage)
- [Standalone Executable Creation](#standalone-executable-creation)
- [Troubleshooting](#troubleshooting)
- [Important Notes](#important-notes)
- [FAQ](#faq)
- [Generic Version (email_bulk_sender.py)](#generic-version-email_bulk_senderpy)

## Features

### Common Features
- **CSV and Excel file (.xlsx) support** - Load recipient list (company name, name, email address)
- **Automatic newline handling** - Compatible with file exchange between Windows/Unix/Mac
- Automatically replace `{企業}` (company) and `{氏名}` (name) placeholders with individual recipient information
- Manage subject and body in a single template file
- Support for CC, BCC, and Reply-To
- Support for file attachments
- Automatic character encoding detection (UTF-8, Shift_JIS, EUC-JP, etc.)
- Rate limiting protection (automatic sending interval adjustment)
- Real-time sending status display
- **Multi-language support** (Japanese/English switchable)
- **Configuration file support** (Save and load settings in JSON format)

### Additional GUI Features
- Visual and intuitive interface
- File selection dialogs
- Email preview before sending
- Real-time progress bar
- Detailed sending log display
- Clear error messages

## Requirements

- Python 3.6 or higher
- Gmail account with 2-Step Verification enabled
- Internet connection

### Required Python Libraries

#### CLI Version (Basic Features)
Only standard libraries are used, so no additional installation is required.

#### Excel File Support (Optional)
To use Excel files (.xlsx), install the following:
```bash
pip install openpyxl
```

#### GUI Version
To use the GUI version, install the following:
```bash
pip install customtkinter chardet openpyxl
```

## Setup

### 1. Download the Script

```bash
# Clone the repository or download the files
git clone <repository-url>
cd email-bulk-sender

# Or download bulk_email_sender.py directly
```

### 2. Virtual Environment Setup (Recommended)

Using a virtual environment allows you to manage packages without affecting your system-wide Python installation.

#### Creating and Activating Virtual Environment

**Unix/Mac/Linux:**
```bash
# Create virtual environment
python3 -m venv venv_email

# Activate virtual environment
source venv_email/bin/activate
```

**Windows:**
```bash
# Create virtual environment
python -m venv venv_email

# Activate virtual environment (Command Prompt)
venv_email\Scripts\activate

# Activate virtual environment (PowerShell)
# If running PowerShell for the first time, you need to set the execution policy
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
venv_email\Scripts\Activate.ps1
```

#### Installing Required Packages

After activating the virtual environment, install required packages:

**For GUI version:**
```bash
pip install customtkinter chardet openpyxl
```

**For Excel file support only (CLI version):**
```bash
pip install openpyxl
```

**Note:** No additional packages are required for CLI version with CSV files only.

#### Deactivating Virtual Environment

When you're done, deactivate the virtual environment:
```bash
deactivate
```

### 3. Obtain Gmail App Password

**Important**: You must use an "App Password", not your regular Gmail password.

#### Steps:

1. **Access Google Account**
   - Go to [https://myaccount.google.com/](https://myaccount.google.com/)
   - Sign in with your Gmail account

2. **Enable 2-Step Verification** (if not already enabled)
   - Select "Security" from the left menu
   - Click "2-Step Verification" under "Signing in to Google"
   - Follow the on-screen instructions

3. **Generate App Password**
   - Return to the "Security" page
   - Enter "App passwords" in the search bar at the top
   - In "To create a new app-specific password, enter the app name below", enter any app name (e.g., "Bulk Email Sender")
   - Click "Create"

4. **Copy the Password**
   - Copy the 16-character password displayed
   - **This password will not be shown again, so save it securely**

   Example: `abcd efgh ijkl mnop` (spaces are not needed when entering)

### 3. Verify File Structure

Place the following files in your project folder:

```
email-bulk-sender/
├── bulk_email_sender.py  # Main script
├── list.csv              # Recipient list
├── body.txt              # Email template
└── README.md             # This file
```

## File Preparation

**Using Sample Files:**

Sample files are provided in the `examples/` directory. You can copy and start using them immediately:

**Unix/Mac/Linux:**
```bash
cp examples/list_sample.csv list.csv
cp examples/body_sample.txt body.txt
# Edit the files as needed
```

**Windows (Command Prompt):**
```bash
copy examples\list_sample.csv list.csv
copy examples\body_sample_win.txt body.txt
# Edit the files as needed
```

**Available Sample Files:**
- `list_sample.csv` - Recipient list in CSV format
- `list_sample.xlsx` - Recipient list in Excel format
- `body_sample.txt` - Email template (UTF-8, Unix line endings)
- `body_sample_win.txt` - Email template for Windows (Shift JIS, CR+LF)
- `body_en_sample.txt` - English email template

See [examples/README.md](examples/README.md) for details.

### Recipient List (list.csv or list.xlsx)

Enter recipient information in CSV or Excel file (.xlsx).

#### CSV File Format (list.csv)

**Format:**
```csv
企業,氏名,メールアドレス
株式会社ABC,山田太郎,yamada@example.com
XYZ商事株式会社,佐藤花子,sato@example.com
テクノロジー株式会社,鈴木一郎,suzuki@example.com
サンプル株式会社,田中美咲,tanaka@example.com
```

**Notes:**
- First line must be the header row (`企業,氏名,メールアドレス`)
- Save with UTF-8 encoding
- If editing in Excel, select "CSV UTF-8 (Comma delimited)" when saving
- Automatic newline handling for file exchange between Windows/Unix/Mac

#### Excel File Format (list.xlsx)

**Features:**
- Direct editing and saving in Excel (no CSV conversion needed)
- No newline code issues
- No character encoding concerns

**Format:**
| 企業 | 氏名 | メールアドレス |
|------|------|----------------|
| 株式会社ABC | 山田太郎 | yamada@example.com |
| XYZ商事株式会社 | 佐藤花子 | sato@example.com |

**Notes:**
- First line must be the header row (`企業`, `氏名`, `メールアドレス`)
- The active sheet (first sheet) will be read
- Requires openpyxl library (`pip install openpyxl`)

### body.txt (Email Template)

Write the email subject and body.

**Format:**
```
[Important] New Service Announcement - {企業} {氏名}様

{企業}
{氏名}様

Thank you for your continued support.
This is the Sales Department of Sample Corporation.

We hope {企業} is prospering.

We are pleased to announce the release of our new service "XXX".

■ Service Overview
- Feature 1: XXX
- Feature 2: YYY
- Feature 3: ZZZ

For more details, please visit:
https://example.com/service

If you have any questions, please feel free to contact us.

We look forward to your continued support.

────────────────────────
Sample Corporation
Sales Department - Taro Yamada
Email: info@example.com
Tel: 03-1234-5678
────────────────────────
```

**Rules:**
- **Line 1**: Subject (can use `{企業}` for company name and `{氏名}` for recipient name)
- **Line 2**: Empty line (required)
- **Line 3 onwards**: Body (can use `{企業}` for company name and `{氏名}` for recipient name)

## Usage (CLI Version)

### Language Settings

All programs support both Japanese and English.

**CLI Version (Command Line):**
```bash
# Run in Japanese (default, auto-detects system language)
python email_bulk_sender.py

# Run in English
python email_bulk_sender.py --lang en

# Explicitly specify Japanese
python email_bulk_sender.py --lang ja
```

**GUI Version:**
- Select language in the "Language" tab within the application
- Selected language is saved to a configuration file and applied on next startup
- Please restart the application after changing the language

### 1. Run the Script

```bash
python bulk_email_sender.py
```

### 2. Interactive Input

You will be prompted to enter the following information:

```
=== Gmail一斉送信ツール ===

送信元Gmailアドレス: your.email@gmail.com
Gmailアプリパスワード: ****************
受信者リストCSVファイル (デフォルト: list.csv): 
メールテンプレートファイル (デフォルト: body.txt): 
CC (複数の場合はカンマ区切り、不要ならEnter): cc@example.com
BCC (複数の場合はカンマ区切り、不要ならEnter): 
Reply-To (不要ならEnter): reply@example.com
```

**Input Tips:**
- **Gmail App Password**: Enter the 16-character password (without spaces)
- **File names**: Press Enter to use default values (list.csv, body.txt)
- **CC/BCC**: For multiple addresses, separate with commas like `email1@example.com,email2@example.com`
- **Optional fields**: Just press Enter to skip

### 3. Confirm Sending Details

After input, the sending details will be displayed:

```
=== 送信内容確認 ===
件名: 【重要】新サービスのご案内 - {企業} {氏名}様
送信先: 4件
送信元: your.email@gmail.com
CC: cc@example.com
Reply-To: reply@example.com

送信を開始しますか？ (yes/no):
```

### 4. Start Sending

Type `yes` to begin sending:

```
[1/4] 送信成功: 株式会社ABC 山田太郎 (yamada@example.com)
[2/4] 送信成功: XYZ商事株式会社 佐藤花子 (sato@example.com)
[3/4] 送信成功: テクノロジー株式会社 鈴木一郎 (suzuki@example.com)
[4/4] 送信成功: サンプル株式会社 田中美咲 (tanaka@example.com)

送信完了: 成功 4件, 失敗 0件
```

## Using Configuration Files

### Overview

Starting from version 2.1, you can save and load settings in JSON format. This allows you to manage settings without editing scripts.

**Key Features:**
- No need to edit `DEFAULT_*` variables in scripts
- Save and load settings from both GUI and CLI
- Directly edit JSON files with a text editor
- Passwords are not saved (for security)

### Configuration File Locations

- **Generic version**: `~/.email_bulk_sender/config.json`
- **Gmail version**: `~/.gmail_bulk_sender/config.json`

(`~` represents the user's home directory)

### CLI Version Usage

#### Saving Settings

```bash
# Enter and save settings
python email_bulk_sender.py --save-config

# For Gmail version
python gmail_bulk_sender.py --save-config
```

When you run for the first time and enter settings, they will be saved to the configuration file.

#### Loading Settings

In CLI version, configuration files are loaded only when `--load-config` flag is specified:

```bash
# Default behavior: Uses DEFAULT_* variables in the script
python email_bulk_sender.py

# Load from configuration file
python email_bulk_sender.py --load-config

# For Gmail version
python gmail_bulk_sender.py --load-config
```

**Behavior Differences:**
- **Without flag**: Uses `DEFAULT_*` variables in the script. Configuration files are ignored even if they exist.
- **With --load-config**: Loads from configuration file. You will only be prompted for items not saved in the file (e.g., password).

**Note**: GUI version automatically loads configuration files on startup.

### GUI Version Usage

In the GUI version, there is a "Configuration Management" section in the Basic Settings tab:

1. **Save Settings**
   - Saves currently entered settings to the configuration file
   - Click the "Save Settings" button
   - Passwords are not saved

2. **Load Settings**
   - Loads saved settings
   - Click the "Load Settings" button
   - All fields will be populated with the saved settings

### Configuration File Format

The configuration file is in JSON format. You can edit it directly with a text editor:

```json
{
  "version": "2.0",
  "smtp": {
    "server": "smtp.gmail.com",
    "port": 587
  },
  "sender": {
    "email_address": "your.email@gmail.com",
    "display_name": "Sample Corp Sales Dept"
  },
  "files": {
    "csv_file": "list.csv",
    "template_file": "body.txt",
    "attachments": ["file1.pdf", "file2.docx"]
  },
  "email_options": {
    "cc": "cc@example.com",
    "bcc": "",
    "reply_to": "reply@example.com",
    "send_delay": 5
  },
  "ui": {
    "language": "en"
  }
}
```

**Note**: Passwords are not saved in the configuration file for security reasons. You need to enter them each time.

### Using Multiple Configurations

You can copy the configuration file to use different settings for different purposes:

```bash
# Backup configuration files
cp ~/.email_bulk_sender/config.json ~/.email_bulk_sender/config_work.json
cp ~/.email_bulk_sender/config.json ~/.email_bulk_sender/config_personal.json

# Switch configurations as needed
cp ~/.email_bulk_sender/config_work.json ~/.email_bulk_sender/config.json
```

## Troubleshooting

### Error: Authentication Failed

```
SMTP接続エラー: (535, b'5.7.8 Username and Password not accepted...')
```

**Causes and Solutions:**
1. Incorrect app password → Verify or regenerate
2. 2-Step Verification disabled → Enable in Google Account
3. Using regular password → Use app password instead

### Error: File Not Found

```
FileNotFoundError: [Errno 2] No such file or directory: 'list.csv'
```

**Solutions:**
- Verify file name and path
- Confirm file is in the same folder as the script
- Check case sensitivity of file name

### Error: Sending Limit Exceeded

```
SMTPDataError: (550, b'5.4.5 Daily sending quota exceeded')
```

**Solutions:**
- Free Gmail: 500 messages per day limit
- Google Workspace: 2,000 messages per day limit
- Wait until the next day or distribute across multiple accounts

### Character Encoding Issues

**Solutions:**
- Save CSV file as UTF-8
- Save body.txt as UTF-8
- On Windows, open in Notepad and select "UTF-8" when saving

## Important Notes

### Sending Limits

- **Free Gmail**: Up to 500 messages per day
- **Google Workspace**: Up to 2,000 messages per day
- Be cautious with hourly limits for large volumes

### Security

- Do not hardcode app passwords in the script
- Do not commit app passwords to Git
- Store securely after use
- Delete unused app passwords

### Best Practices

1. **Test Sending**: Send to yourself before production use
2. **Start Small**: Begin with 5-10 recipients to verify
3. **Verify List**: Always check content before sending
4. **Error Logging**: Record failed recipients for resending
5. **Unsubscribe Option**: Include opt-out links

### Legal Considerations

- Do not send marketing emails without consent
- Comply with anti-spam laws (e.g., CAN-SPAM Act)
- Provide unsubscribe mechanisms
- Properly manage personal information

## FAQ

### Q: Can I send HTML emails?

A: The current version only supports plain text. For HTML emails, change the second parameter of `MIMEText` to `'html'` and write HTML in body.txt.

### Q: Can I send attachments?

A: Yes, it is supported. In the GUI version, you can add attachments in the "File Selection Tab". The CLI version also has attachment functionality implemented.

### Q: Can I change the sending interval?

A: Yes, modify the `delay` parameter (default 1 second) in the `send_bulk_emails` method.

### Q: Can I use multiple variables (besides company and name)?

A: Currently, two placeholders `{企業}` (company) and `{氏名}` (name) are available. To add more, you can add columns to the CSV and modify the script. For example, you can add `{役職}` (title).

## Generic Version (email_bulk_sender.py)

### Overview

`email_bulk_sender.py` is a generic bulk email sender that supports any SMTP server. Use this when you need to send emails from non-Gmail services (Outlook, Yahoo, custom mail servers, etc.) or when you need to customize the sender display name.

### Key Differences from Gmail-Specific Version

1. **SMTP Server Support**
   - Gmail-specific: Gmail SMTP only
   - Generic: Any SMTP server

2. **Configuration Method**
   - Gmail-specific: Enter Gmail address and app password at runtime
   - Generic: Pre-configure in settings section (runtime input also available)

3. **Sender Display Name**
   - Gmail-specific: Email address only
   - Generic: Customizable display name (e.g., "Sample Corp Sales Department")

4. **SSL/TLS Support**
   - Gmail-specific: TLS only
   - Generic: Auto-detects based on port (465=SSL, 587=TLS)

### Usage

#### 1. Edit Configuration (Optional)

Open `email_bulk_sender.py` and edit the settings section at the top:

```python
# SMTP server (leave empty to prompt at runtime)
DEFAULT_SMTP_SERVER = "smtp.example.com"  # e.g., "smtp.gmail.com", "smtp.office365.com"

# SMTP port (leave empty to prompt at runtime)
DEFAULT_SMTP_PORT = "587"  # 587 (TLS), 465 (SSL), 25 (unencrypted)

# Email address (leave empty to prompt at runtime)
DEFAULT_EMAIL_ADDRESS = "your.email@example.com"

# Email password (recommended to leave empty for security)
DEFAULT_EMAIL_PASSWORD = ""

# Sender display name (leave empty to show email address only)
SENDER_DISPLAY_NAME = "Sample Corp Sales Department"

# Recipient list CSV file
DEFAULT_CSV_FILE = "list.csv"

# Email template file
DEFAULT_TEMPLATE_FILE = "body.txt"

# CC, BCC, Reply-To (optional)
DEFAULT_CC = ""
DEFAULT_BCC = ""
DEFAULT_REPLY_TO = ""
```

**Configuration Tips:**
- Items left as empty strings `""` will prompt for input at runtime
- **For security, keep passwords empty** and enter at runtime
- Configured items will show "(configured)" at runtime and skip prompts

#### 2. File Preparation

Use the same format for `list.csv` and `body.txt` as `bulk_email_sender.py` (see [File Preparation](#file-preparation)).

#### 3. Run the Script

```bash
python email_bulk_sender.py
```

You'll be prompted for unconfigured items:

```
=== Email Bulk Sender Tool ===

SMTP server (e.g., smtp.gmail.com): smtp.example.com
SMTP port (default: 587): 587
Sender email address: your.email@example.com
Email password: ********
Sender display name (press Enter to skip): Sample Corp Sales Department
Recipient list CSV file (default: list.csv):
Email template file (default: body.txt):
CC (comma-separated for multiple, press Enter to skip):
BCC (comma-separated for multiple, press Enter to skip):
Reply-To (press Enter to skip):
```

### SMTP Settings for Major Email Services

| Service | SMTP Server | Port | Security |
|---------|------------|------|----------|
| Gmail | smtp.gmail.com | 587 | TLS |
| Outlook/Office365 | smtp.office365.com | 587 | TLS |
| Yahoo Mail | smtp.mail.yahoo.com | 587 | TLS |
| iCloud | smtp.mail.me.com | 587 | TLS |

**Note:** When using Gmail, you need to obtain an App Password (see [Obtain Gmail App Password](#2-obtain-gmail-app-password)).

### Troubleshooting (Generic Version Specific)

#### Port Number Error

```
ConnectionRefusedError: [Errno 111] Connection refused
```

**Solutions:**
- Verify correct port number (typically 587 or 465)
- Check if firewall is blocking the port
- Verify SMTP server address is correct

#### SSL/TLS Error

```
ssl.SSLError: [SSL: WRONG_VERSION_NUMBER]
```

**Solutions:**
- Port 465 uses SSL, port 587 uses TLS
- Verify port number and server configuration combination

## GUI Version Usage

### Required Packages

For GUI version, additional packages are required:

```bash
pip install customtkinter chardet openpyxl
```

Note: `openpyxl` is required for Excel file (.xlsx) support. If you only use CSV files, it's not required, but installation is recommended.

### How to Launch

#### Generic GUI Version
```bash
python email_bulk_sender_gui.py
```

#### Gmail-specific GUI Version
```bash
python gmail_bulk_sender_gui.py
```

### Operating Instructions

1. **Basic Settings Tab**
   - Enter SMTP server information (generic version) or Gmail account information (Gmail version)
   - Enter email address and password
   - Enter sender display name (optional)
   - **Configuration Management**: Use "Load Settings" or "Save Settings" buttons to manage your settings

2. **File Selection Tab**
   - Click "File Selection" button to select recipient list CSV file
   - Click "File Selection" button to select email template file
   - If you have attachments, click "Add File" button to add them

3. **Options Tab**
   - Enter CC, BCC, Reply-To (optional)
   - Set sending interval (seconds)

4. **Send Tab**
   - Click "Preview Email Content" button to review the content
   - Click "Start Sending" button to begin sending
   - Monitor progress with progress bar and logs

### GUI Version Features

- **File Selection Dialogs**: No need to manually enter file paths
- **Preview Function**: Review content before sending
- **Real-time Progress**: Visual progress bar for sending status
- **Detailed Logs**: Real-time display of sending success/failure details
- **Error Handling**: Clear and understandable error messages

## Standalone Executable Creation

You can create standalone executables (Windows: .exe, Linux/Mac: binary) that can run on systems without Python installed.

The build script automatically creates:
- Standalone executables
- **Distribution zip packages** (executable + sample files + README + LICENSE)

### Automatic Build (Recommended)

#### Windows
```cmd
build_executables.bat
```

Created files:
- `dist/EmailBulkSender_win.exe` - Generic version executable
- `dist/EmailBulkSender_win.zip` - Generic version distribution package ⭐
- `dist/GmailBulkSender_win.exe` - Gmail version executable
- `dist/GmailBulkSender_win.zip` - Gmail version distribution package ⭐

#### Linux/Mac
```bash
chmod +x build_executables.sh
./build_executables.sh
```

Created files (example: Linux):
- `dist/EmailBulkSender_lnx` - Generic version executable
- `dist/EmailBulkSender_lnx.zip` - Generic version distribution package ⭐
- `dist/GmailBulkSender_lnx` - Gmail version executable
- `dist/GmailBulkSender_lnx.zip` - Gmail version distribution package ⭐

### Distribution Package Contents

The zip file includes:
```
EmailBulkSender_xxx.zip
├── EmailBulkSender_xxx[.exe]  # Executable file
├── examples/                   # Sample files
│   ├── list.csv.sample
│   ├── body.txt.sample
│   ├── body_en.txt.sample
│   └── README.md
├── README.md                   # User guide
└── LICENSE                     # License
```

**How to Distribute:**
Distribute the zip file, and recipients can start using it immediately:
1. Extract the zip file
2. Copy `examples/list.csv.sample` to `list.csv`
3. Copy `examples/body.txt.sample` to `body.txt`
4. Edit the files and double-click the executable

### Manual Build

1. Install PyInstaller:
   ```bash
   pip install pyinstaller customtkinter chardet openpyxl
   ```

2. Create executables:

   **Windows:**
   ```bash
   # Generic GUI version
   pyinstaller --onefile --windowed --name="EmailBulkSender_win" email_bulk_sender_gui.py

   # Gmail-specific GUI version
   pyinstaller --onefile --windowed --name="GmailBulkSender_win" gmail_bulk_sender_gui.py
   ```

   **Linux:**
   ```bash
   # Generic GUI version
   pyinstaller --onefile --windowed --name="EmailBulkSender_lnx" email_bulk_sender_gui.py

   # Gmail-specific GUI version
   pyinstaller --onefile --windowed --name="GmailBulkSender_lnx" gmail_bulk_sender_gui.py
   ```

   **macOS:**
   ```bash
   # ARM64 (M1/M2/M3 Mac)
   pyinstaller --onefile --windowed --name="EmailBulkSender_mac_arm64" email_bulk_sender_gui.py
   pyinstaller --onefile --windowed --name="GmailBulkSender_mac_arm64" gmail_bulk_sender_gui.py

   # AMD64 (Intel Mac)
   pyinstaller --onefile --windowed --name="EmailBulkSender_mac_amd64" email_bulk_sender_gui.py
   pyinstaller --onefile --windowed --name="GmailBulkSender_mac_amd64" gmail_bulk_sender_gui.py
   ```

3. Executables will be created in the `dist/` directory

See [BUILD_GUIDE.md](BUILD_GUIDE.md) for details.

### Distributing Executables

You can distribute the created executables (e.g., `dist/EmailBulkSender_win.exe`, `dist/EmailBulkSender_lnx`, etc.), which can run on systems without Python installed.

**Important Notes:**
- To create Windows .exe files, you must build on a Windows environment
- To create Linux binaries, you must build on a Linux environment
- To create Mac binaries, you must build on a Mac environment

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

For issues or questions, please create an Issue.

---

**Created**: 23 September 2025
**Updated**: 26 November 2025
**Version**: 2.3 (Configuration loading logic improvements)
