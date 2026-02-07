#!/usr/bin/env python3
"""
メール一括送信ツール GUI版
Email Bulk Sender - GUI Application
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.ttk as ttk
import threading
import time
import os
import smtplib
from typing import Dict, Any, Optional, List

# CLI版からビジネスロジックを再利用
from email_bulk_sender import EmailBulkSender, I18n, InternalConfigManager

# keyring サポート（パスワード保存用）
try:
    import keyring
    KEYRING_AVAILABLE = True
except ImportError:
    KEYRING_AVAILABLE = False

KEYRING_SERVICE = "email_bulk_sender"


# ==================== I18n 拡張 ====================

class GuiI18n(I18n):
    """GUI用のI18n拡張クラス"""

    GUI_TEXTS = {
        'ja': {
            # ウィンドウ
            'app_title': 'メール一括送信ツール',

            # ツールバー
            'menu_save_config': '設定を保存',
            'menu_load_config': '設定を読み込み',
            'menu_language': '言語',

            # SMTP設定セクション
            'section_smtp': 'SMTP設定',
            'label_smtp_server': 'SMTPサーバー',
            'label_smtp_port': 'SMTPポート',
            'placeholder_smtp_server': '例: smtp.gmail.com',
            'placeholder_smtp_port': '587',

            # 送信元設定セクション
            'section_sender': '送信元設定',
            'label_email': 'メールアドレス',
            'label_password': 'パスワード',
            'label_display_name': '表示名',
            'placeholder_email': '例: user@example.com',
            'placeholder_display_name': '例: 株式会社サンプル 営業部',

            # ファイル設定セクション
            'section_files': 'ファイル設定',
            'label_recipient_file': '受信者リスト',
            'label_template_file': 'テンプレート',
            'label_attachments': '添付ファイル',
            'button_browse': '参照...',
            'placeholder_recipient_file': 'CSV または Excel ファイル',
            'placeholder_template_file': 'メールテンプレートファイル',
            'placeholder_attachments': 'カンマ区切りで複数指定可',

            # メールオプションセクション
            'section_options': 'メールオプション',
            'label_cc': 'CC',
            'label_bcc': 'BCC',
            'label_reply_to': 'Reply-To',
            'label_delay': '送信間隔（秒）',
            'placeholder_cc': 'カンマ区切りで複数指定可',
            'placeholder_bcc': 'カンマ区切りで複数指定可',
            'placeholder_reply_to': '返信先アドレス',

            # ボタン
            'button_recipient_list': '受信者リスト確認',
            'button_preview': 'プレビュー',
            'button_test_send': 'テスト送信',
            'button_send': '送信開始',
            'button_cancel': 'キャンセル',
            'button_close': '閉じる',

            # ログ
            'section_log': '送信ログ',

            # ステータス
            'status_ready': '準備完了',
            'status_sending': '送信中... ({0}/{1})',
            'status_complete': '送信完了: 成功 {0}件, 失敗 {1}件',
            'status_cancelled': '送信がキャンセルされました',
            'status_config_saved': '設定を保存しました',
            'status_config_loaded': '設定を読み込みました',
            'status_config_not_found': '設定ファイルが見つかりません',
            'status_keyring_not_available': '注意: keyringが利用できないため、パスワードは保存されません',

            # 受信者リストダイアログ
            'dialog_recipient_title': '受信者リスト一覧',
            'dialog_recipient_count': '受信者数: {0}件',
            'dialog_recipient_col_affiliation': '所属',
            'dialog_recipient_col_name': '氏名',
            'dialog_recipient_col_email': 'メールアドレス',

            # プレビューダイアログ
            'dialog_preview_title': '送信前プレビュー',
            'dialog_preview_subject': '件名',
            'dialog_preview_from': '送信元',
            'dialog_preview_to': '送信先（1件目）',
            'dialog_preview_cc': 'CC',
            'dialog_preview_bcc': 'BCC',
            'dialog_preview_reply_to': 'Reply-To',
            'dialog_preview_attachments': '添付ファイル',
            'dialog_preview_body': '本文',

            # テスト送信ダイアログ
            'dialog_test_title': 'テスト送信',
            'dialog_test_address': '送信先メールアドレス',
            'dialog_test_send': 'テスト送信実行',
            'dialog_test_sending': 'テスト送信中...',
            'dialog_test_success': 'テスト送信が成功しました',
            'dialog_test_failed': 'テスト送信に失敗しました: {0}',

            # エラー
            'error_no_recipient_file': '受信者リストファイルを指定してください',
            'error_no_template_file': 'テンプレートファイルを指定してください',
            'error_file_not_found': 'ファイルが見つかりません: {0}',
            'error_no_smtp_server': 'SMTPサーバーを入力してください',
            'error_no_email': 'メールアドレスを入力してください',
            'error_no_password': 'パスワードを入力してください',
            'error_invalid_port': '無効なポート番号です',
            'error_read_failed': 'ファイルの読み込みに失敗しました: {0}',
            'error_send_failed': '送信エラー: {0}',
            'error_no_test_address': 'テスト送信先アドレスを入力してください',

            # 確認
            'confirm_send_title': '送信確認',
            'confirm_send_message': '{0}件の宛先にメールを送信します。よろしいですか？',

            # ファイルダイアログ
            'filedialog_recipient': '受信者リストを選択',
            'filedialog_template': 'テンプレートファイルを選択',
            'filedialog_attachment': '添付ファイルを選択',
        },
        'en': {
            'app_title': 'Email Bulk Sender',
            'menu_save_config': 'Save Settings',
            'menu_load_config': 'Load Settings',
            'menu_language': 'Language',
            'section_smtp': 'SMTP Settings',
            'label_smtp_server': 'SMTP Server',
            'label_smtp_port': 'SMTP Port',
            'placeholder_smtp_server': 'e.g., smtp.gmail.com',
            'placeholder_smtp_port': '587',
            'section_sender': 'Sender Settings',
            'label_email': 'Email Address',
            'label_password': 'Password',
            'label_display_name': 'Display Name',
            'placeholder_email': 'e.g., user@example.com',
            'placeholder_display_name': 'e.g., Sales Department',
            'section_files': 'File Settings',
            'label_recipient_file': 'Recipient List',
            'label_template_file': 'Template',
            'label_attachments': 'Attachments',
            'button_browse': 'Browse...',
            'placeholder_recipient_file': 'CSV or Excel file',
            'placeholder_template_file': 'Email template file',
            'placeholder_attachments': 'Comma-separated for multiple',
            'section_options': 'Email Options',
            'label_cc': 'CC',
            'label_bcc': 'BCC',
            'label_reply_to': 'Reply-To',
            'label_delay': 'Send Delay (sec)',
            'placeholder_cc': 'Comma-separated for multiple',
            'placeholder_bcc': 'Comma-separated for multiple',
            'placeholder_reply_to': 'Reply-to address',
            'button_recipient_list': 'View Recipients',
            'button_preview': 'Preview',
            'button_test_send': 'Test Send',
            'button_send': 'Start Sending',
            'button_cancel': 'Cancel',
            'button_close': 'Close',
            'section_log': 'Send Log',
            'status_ready': 'Ready',
            'status_sending': 'Sending... ({0}/{1})',
            'status_complete': 'Complete: {0} succeeded, {1} failed',
            'status_cancelled': 'Sending cancelled',
            'status_config_saved': 'Settings saved',
            'status_config_loaded': 'Settings loaded',
            'status_config_not_found': 'Config file not found',
            'status_keyring_not_available': 'Note: keyring not available, password will not be saved',
            'dialog_recipient_title': 'Recipient List',
            'dialog_recipient_count': 'Recipients: {0}',
            'dialog_recipient_col_affiliation': 'Affiliation',
            'dialog_recipient_col_name': 'Name',
            'dialog_recipient_col_email': 'Email',
            'dialog_preview_title': 'Send Preview',
            'dialog_preview_subject': 'Subject',
            'dialog_preview_from': 'From',
            'dialog_preview_to': 'To (1st recipient)',
            'dialog_preview_cc': 'CC',
            'dialog_preview_bcc': 'BCC',
            'dialog_preview_reply_to': 'Reply-To',
            'dialog_preview_attachments': 'Attachments',
            'dialog_preview_body': 'Body',
            'dialog_test_title': 'Test Send',
            'dialog_test_address': 'Send to address',
            'dialog_test_send': 'Send Test',
            'dialog_test_sending': 'Sending test...',
            'dialog_test_success': 'Test email sent successfully',
            'dialog_test_failed': 'Test send failed: {0}',
            'error_no_recipient_file': 'Please specify a recipient list file',
            'error_no_template_file': 'Please specify a template file',
            'error_file_not_found': 'File not found: {0}',
            'error_no_smtp_server': 'Please enter SMTP server',
            'error_no_email': 'Please enter email address',
            'error_no_password': 'Please enter password',
            'error_invalid_port': 'Invalid port number',
            'error_read_failed': 'Failed to read file: {0}',
            'error_send_failed': 'Send error: {0}',
            'error_no_test_address': 'Please enter test recipient address',
            'confirm_send_title': 'Confirm Send',
            'confirm_send_message': 'Send email to {0} recipients. Proceed?',
            'filedialog_recipient': 'Select Recipient List',
            'filedialog_template': 'Select Template File',
            'filedialog_attachment': 'Select Attachments',
        }
    }

    def __init__(self, lang='ja'):
        super().__init__(lang)

    def get(self, key, *args):
        """GUI用テキストを優先的に検索し、なければ親クラスのテキストを返す"""
        text = self.GUI_TEXTS.get(self.lang, {}).get(key)
        if text is None:
            text = self.TEXTS.get(self.lang, {}).get(key, key)
        if args:
            return text.format(*args)
        return text


# ==================== 設定管理（keyring連携） ====================

class GuiConfigManager(InternalConfigManager):
    """GUI用設定管理クラス（keyring連携）"""

    def save_password(self, email_address: str, password: str) -> bool:
        """keyringにパスワードを保存"""
        if not KEYRING_AVAILABLE:
            return False
        try:
            keyring.set_password(KEYRING_SERVICE, email_address, password)
            return True
        except Exception:
            return False

    def load_password(self, email_address: str) -> Optional[str]:
        """keyringからパスワードを読み込み"""
        if not KEYRING_AVAILABLE:
            return None
        try:
            return keyring.get_password(KEYRING_SERVICE, email_address)
        except Exception:
            return None


# ==================== ダイアログ ====================

class RecipientListDialog(ctk.CTkToplevel):
    """受信者リスト一覧ダイアログ"""

    def __init__(self, parent, recipients: List[Dict], i18n: GuiI18n):
        super().__init__(parent)
        self.i18n = i18n

        self.title(i18n.get('dialog_recipient_title'))
        self.geometry("650x450")
        self.transient(parent)
        self.grab_set()

        # 件数ラベル
        ctk.CTkLabel(
            self,
            text=i18n.get('dialog_recipient_count', len(recipients)),
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(padx=10, pady=(10, 5))

        # Treeview用フレーム
        tree_frame = tk.Frame(self)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=5)

        style = ttk.Style()
        style.configure("Recipient.Treeview", rowheight=28)

        columns = ('num', 'affiliation', 'name', 'email')
        self.tree = ttk.Treeview(
            tree_frame, columns=columns, show='headings',
            style="Recipient.Treeview"
        )

        self.tree.heading('num', text='#')
        self.tree.heading('affiliation', text=i18n.get('dialog_recipient_col_affiliation'))
        self.tree.heading('name', text=i18n.get('dialog_recipient_col_name'))
        self.tree.heading('email', text=i18n.get('dialog_recipient_col_email'))

        self.tree.column('num', width=40, anchor='center')
        self.tree.column('affiliation', width=180)
        self.tree.column('name', width=120)
        self.tree.column('email', width=260)

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for i, r in enumerate(recipients, 1):
            self.tree.insert('', 'end', values=(i, r['affiliation'], r['name'], r['email']))

        # 閉じるボタン
        ctk.CTkButton(
            self, text=i18n.get('button_close'), command=self.destroy
        ).pack(padx=10, pady=10)


class PreviewDialog(ctk.CTkToplevel):
    """送信前プレビューダイアログ"""

    def __init__(self, parent, preview_data: Dict, i18n: GuiI18n):
        super().__init__(parent)
        self.i18n = i18n

        self.title(i18n.get('dialog_preview_title'))
        self.geometry("650x550")
        self.transient(parent)
        self.grab_set()

        # スクロール可能フレーム
        scroll_frame = ctk.CTkScrollableFrame(self)
        scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # ヘッダー情報
        headers = [
            ('dialog_preview_from', preview_data.get('from', '')),
            ('dialog_preview_to', preview_data.get('to', '')),
            ('dialog_preview_subject', preview_data.get('subject', '')),
        ]
        if preview_data.get('cc'):
            headers.append(('dialog_preview_cc', preview_data['cc']))
        if preview_data.get('bcc'):
            headers.append(('dialog_preview_bcc', preview_data['bcc']))
        if preview_data.get('reply_to'):
            headers.append(('dialog_preview_reply_to', preview_data['reply_to']))
        if preview_data.get('attachments'):
            headers.append(('dialog_preview_attachments', preview_data['attachments']))

        for key, value in headers:
            row = ctk.CTkFrame(scroll_frame, fg_color="transparent")
            row.pack(fill="x", pady=2)
            ctk.CTkLabel(
                row, text=f"{i18n.get(key)}:",
                font=ctk.CTkFont(weight="bold"), width=140, anchor="e"
            ).pack(side="left", padx=(0, 5))
            ctk.CTkLabel(
                row, text=str(value), anchor="w", wraplength=450
            ).pack(side="left", fill="x", expand=True)

        # 区切り線
        ctk.CTkFrame(scroll_frame, height=2, fg_color="gray50").pack(fill="x", pady=10)

        # 本文ラベル
        ctk.CTkLabel(
            scroll_frame, text=i18n.get('dialog_preview_body'),
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", pady=(0, 5))

        # 本文表示
        body_text = ctk.CTkTextbox(scroll_frame, height=280)
        body_text.pack(fill="both", expand=True)
        body_text.insert("1.0", preview_data.get('body', ''))
        body_text.configure(state="disabled")

        # 閉じるボタン
        ctk.CTkButton(
            self, text=i18n.get('button_close'), command=self.destroy
        ).pack(padx=10, pady=10)


class TestSendDialog(ctk.CTkToplevel):
    """テスト送信ダイアログ"""

    def __init__(self, parent, app: 'EmailBulkSenderApp', i18n: GuiI18n):
        super().__init__(parent)
        self.app = app
        self.i18n = i18n

        self.title(i18n.get('dialog_test_title'))
        self.geometry("480x230")
        self.transient(parent)
        self.grab_set()

        # 説明
        ctk.CTkLabel(
            self, text=i18n.get('dialog_test_address'),
            font=ctk.CTkFont(size=14)
        ).pack(padx=20, pady=(20, 5))

        # メールアドレス入力
        self.address_entry = ctk.CTkEntry(self, width=380, placeholder_text="test@example.com")
        self.address_entry.pack(padx=20, pady=5)

        # 結果表示ラベル
        self.result_label = ctk.CTkLabel(self, text="", wraplength=430)
        self.result_label.pack(padx=20, pady=10)

        # ボタンフレーム
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=10)

        self.send_btn = ctk.CTkButton(
            btn_frame, text=i18n.get('dialog_test_send'),
            command=self._send_test
        )
        self.send_btn.pack(side="left", padx=5)

        ctk.CTkButton(
            btn_frame, text=i18n.get('button_close'),
            command=self.destroy
        ).pack(side="left", padx=5)

    def _send_test(self):
        """テスト送信実行"""
        address = self.address_entry.get().strip()
        if not address:
            self.result_label.configure(
                text=self.i18n.get('error_no_test_address'), text_color="red"
            )
            return

        self.send_btn.configure(state="disabled")
        self.result_label.configure(
            text=self.i18n.get('dialog_test_sending'), text_color="white"
        )

        thread = threading.Thread(target=self._do_test_send, args=(address,), daemon=True)
        thread.start()

    def _do_test_send(self, address: str):
        """テスト送信の実行（バックグラウンド）"""
        try:
            sender = self.app._create_sender()
            recipients = self.app._read_recipients()
            subject_template, body_template = self.app._read_template()

            # 受信者リストの1件目のデータでプレースホルダーを置換
            if recipients:
                recipient = recipients[0]
            else:
                recipient = {'affiliation': 'Test', 'name': 'Test', 'email': address}

            cc = self.app.cc_entry.get().strip() or None
            bcc = self.app.bcc_entry.get().strip() or None
            reply_to = self.app.reply_to_entry.get().strip() or None
            attachments = self.app._get_attachments()

            msg = sender.create_message(
                to_email=address,
                to_name=recipient['name'],
                to_affiliation=recipient['affiliation'],
                subject_template=subject_template,
                body_template=body_template,
                cc=cc, bcc=bcc, reply_to=reply_to,
                attachments=attachments
            )

            # SMTP送信
            if sender.smtp_port == 465:
                server = smtplib.SMTP_SSL(sender.smtp_server, sender.smtp_port)
            else:
                server = smtplib.SMTP(sender.smtp_server, sender.smtp_port)
                server.starttls()
            server.login(sender.email_address, sender.email_password)
            server.send_message(msg)
            server.quit()

            self.after(0, lambda: self._on_result(True, ""))
        except Exception as e:
            self.after(0, lambda err=str(e): self._on_result(False, err))

    def _on_result(self, success: bool, error: str):
        """テスト送信結果の表示"""
        self.send_btn.configure(state="normal")
        if success:
            self.result_label.configure(
                text=self.i18n.get('dialog_test_success'), text_color="green"
            )
        else:
            self.result_label.configure(
                text=self.i18n.get('dialog_test_failed', error), text_color="red"
            )


# ==================== メインアプリケーション ====================

class EmailBulkSenderApp(ctk.CTk):
    """メール一括送信ツール GUI メインウィンドウ"""

    def __init__(self):
        super().__init__()

        self.i18n = GuiI18n()
        self.config_manager = GuiConfigManager("email")

        # 送信制御用フラグ
        self._sending = False
        self._cancel_requested = False

        # ウィンドウ設定
        self.title(self.i18n.get('app_title'))
        self.geometry("750x880")
        self.minsize(680, 700)

        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        self._create_widgets()
        self._load_config()

    # ==================== UI構築 ====================

    def _create_widgets(self):
        """全ウィジェットを構築"""

        # --- ツールバー ---
        toolbar = ctk.CTkFrame(self, fg_color="transparent")
        toolbar.pack(fill="x", padx=10, pady=(5, 0))

        self.save_btn = ctk.CTkButton(
            toolbar, text=self.i18n.get('menu_save_config'),
            command=self._save_config, width=130
        )
        self.save_btn.pack(side="left", padx=2)

        self.load_btn = ctk.CTkButton(
            toolbar, text=self.i18n.get('menu_load_config'),
            command=self._load_config, width=130
        )
        self.load_btn.pack(side="left", padx=2)

        # 言語切り替え
        self.lang_var = ctk.StringVar(value=self.i18n.get_language().upper())
        self.lang_menu = ctk.CTkSegmentedButton(
            toolbar, values=["JA", "EN"],
            variable=self.lang_var,
            command=self._change_language
        )
        self.lang_menu.pack(side="right", padx=2)
        self.lang_label = ctk.CTkLabel(toolbar, text=self.i18n.get('menu_language'))
        self.lang_label.pack(side="right", padx=(0, 5))

        # --- メインスクロールエリア ---
        self.main_frame = ctk.CTkScrollableFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self._create_smtp_section()
        self._create_sender_section()
        self._create_files_section()
        self._create_options_section()
        self._create_buttons_section()
        self._create_log_section()

        # --- ステータスバー ---
        self.status_label = ctk.CTkLabel(self, text=self.i18n.get('status_ready'), anchor="w")
        self.status_label.pack(fill="x", padx=10, pady=(0, 5))

    def _create_smtp_section(self):
        """SMTP設定セクション"""
        frame = ctk.CTkFrame(self.main_frame)
        frame.pack(fill="x", pady=(0, 5))

        ctk.CTkLabel(
            frame, text=self.i18n.get('section_smtp'),
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=10, pady=(5, 0))

        grid = ctk.CTkFrame(frame, fg_color="transparent")
        grid.pack(fill="x", padx=10, pady=5)
        grid.columnconfigure(1, weight=1)

        ctk.CTkLabel(grid, text=self.i18n.get('label_smtp_server')).grid(
            row=0, column=0, sticky="e", padx=(0, 8), pady=2
        )
        self.smtp_server_entry = ctk.CTkEntry(
            grid, placeholder_text=self.i18n.get('placeholder_smtp_server')
        )
        self.smtp_server_entry.grid(row=0, column=1, sticky="ew", pady=2)

        ctk.CTkLabel(grid, text=self.i18n.get('label_smtp_port')).grid(
            row=1, column=0, sticky="e", padx=(0, 8), pady=2
        )
        self.port_entry = ctk.CTkEntry(
            grid, placeholder_text=self.i18n.get('placeholder_smtp_port'), width=100
        )
        self.port_entry.grid(row=1, column=1, sticky="w", pady=2)

    def _create_sender_section(self):
        """送信元設定セクション"""
        frame = ctk.CTkFrame(self.main_frame)
        frame.pack(fill="x", pady=5)

        ctk.CTkLabel(
            frame, text=self.i18n.get('section_sender'),
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=10, pady=(5, 0))

        grid = ctk.CTkFrame(frame, fg_color="transparent")
        grid.pack(fill="x", padx=10, pady=5)
        grid.columnconfigure(1, weight=1)

        ctk.CTkLabel(grid, text=self.i18n.get('label_email')).grid(
            row=0, column=0, sticky="e", padx=(0, 8), pady=2
        )
        self.email_entry = ctk.CTkEntry(
            grid, placeholder_text=self.i18n.get('placeholder_email')
        )
        self.email_entry.grid(row=0, column=1, sticky="ew", pady=2)

        ctk.CTkLabel(grid, text=self.i18n.get('label_password')).grid(
            row=1, column=0, sticky="e", padx=(0, 8), pady=2
        )
        self.password_entry = ctk.CTkEntry(grid, show="*")
        self.password_entry.grid(row=1, column=1, sticky="ew", pady=2)

        ctk.CTkLabel(grid, text=self.i18n.get('label_display_name')).grid(
            row=2, column=0, sticky="e", padx=(0, 8), pady=2
        )
        self.display_name_entry = ctk.CTkEntry(
            grid, placeholder_text=self.i18n.get('placeholder_display_name')
        )
        self.display_name_entry.grid(row=2, column=1, sticky="ew", pady=2)

    def _create_files_section(self):
        """ファイル設定セクション"""
        frame = ctk.CTkFrame(self.main_frame)
        frame.pack(fill="x", pady=5)

        ctk.CTkLabel(
            frame, text=self.i18n.get('section_files'),
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=10, pady=(5, 0))

        grid = ctk.CTkFrame(frame, fg_color="transparent")
        grid.pack(fill="x", padx=10, pady=5)
        grid.columnconfigure(1, weight=1)

        # 受信者リスト
        ctk.CTkLabel(grid, text=self.i18n.get('label_recipient_file')).grid(
            row=0, column=0, sticky="e", padx=(0, 8), pady=2
        )
        self.recipient_entry = ctk.CTkEntry(
            grid, placeholder_text=self.i18n.get('placeholder_recipient_file')
        )
        self.recipient_entry.grid(row=0, column=1, sticky="ew", pady=2)
        ctk.CTkButton(
            grid, text=self.i18n.get('button_browse'), width=80,
            command=self._browse_recipient
        ).grid(row=0, column=2, padx=(5, 0), pady=2)

        # テンプレート
        ctk.CTkLabel(grid, text=self.i18n.get('label_template_file')).grid(
            row=1, column=0, sticky="e", padx=(0, 8), pady=2
        )
        self.template_entry = ctk.CTkEntry(
            grid, placeholder_text=self.i18n.get('placeholder_template_file')
        )
        self.template_entry.grid(row=1, column=1, sticky="ew", pady=2)
        ctk.CTkButton(
            grid, text=self.i18n.get('button_browse'), width=80,
            command=self._browse_template
        ).grid(row=1, column=2, padx=(5, 0), pady=2)

        # 添付ファイル
        ctk.CTkLabel(grid, text=self.i18n.get('label_attachments')).grid(
            row=2, column=0, sticky="e", padx=(0, 8), pady=2
        )
        self.attachments_entry = ctk.CTkEntry(
            grid, placeholder_text=self.i18n.get('placeholder_attachments')
        )
        self.attachments_entry.grid(row=2, column=1, sticky="ew", pady=2)
        ctk.CTkButton(
            grid, text=self.i18n.get('button_browse'), width=80,
            command=self._browse_attachments
        ).grid(row=2, column=2, padx=(5, 0), pady=2)

    def _create_options_section(self):
        """メールオプションセクション"""
        frame = ctk.CTkFrame(self.main_frame)
        frame.pack(fill="x", pady=5)

        ctk.CTkLabel(
            frame, text=self.i18n.get('section_options'),
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=10, pady=(5, 0))

        grid = ctk.CTkFrame(frame, fg_color="transparent")
        grid.pack(fill="x", padx=10, pady=5)
        grid.columnconfigure(1, weight=1)

        ctk.CTkLabel(grid, text=self.i18n.get('label_cc')).grid(
            row=0, column=0, sticky="e", padx=(0, 8), pady=2
        )
        self.cc_entry = ctk.CTkEntry(
            grid, placeholder_text=self.i18n.get('placeholder_cc')
        )
        self.cc_entry.grid(row=0, column=1, sticky="ew", pady=2)

        ctk.CTkLabel(grid, text=self.i18n.get('label_bcc')).grid(
            row=1, column=0, sticky="e", padx=(0, 8), pady=2
        )
        self.bcc_entry = ctk.CTkEntry(
            grid, placeholder_text=self.i18n.get('placeholder_bcc')
        )
        self.bcc_entry.grid(row=1, column=1, sticky="ew", pady=2)

        ctk.CTkLabel(grid, text=self.i18n.get('label_reply_to')).grid(
            row=2, column=0, sticky="e", padx=(0, 8), pady=2
        )
        self.reply_to_entry = ctk.CTkEntry(
            grid, placeholder_text=self.i18n.get('placeholder_reply_to')
        )
        self.reply_to_entry.grid(row=2, column=1, sticky="ew", pady=2)

        ctk.CTkLabel(grid, text=self.i18n.get('label_delay')).grid(
            row=3, column=0, sticky="e", padx=(0, 8), pady=2
        )
        self.delay_entry = ctk.CTkEntry(grid, width=100)
        self.delay_entry.grid(row=3, column=1, sticky="w", pady=2)
        self.delay_entry.insert(0, "5")

    def _create_buttons_section(self):
        """操作ボタンセクション"""
        button_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        button_frame.pack(fill="x", pady=10)

        ctk.CTkButton(
            button_frame, text=self.i18n.get('button_recipient_list'),
            command=self._show_recipient_list
        ).pack(side="left", padx=3)

        ctk.CTkButton(
            button_frame, text=self.i18n.get('button_preview'),
            command=self._show_preview
        ).pack(side="left", padx=3)

        ctk.CTkButton(
            button_frame, text=self.i18n.get('button_test_send'),
            command=self._show_test_send
        ).pack(side="left", padx=3)

        self.send_btn = ctk.CTkButton(
            button_frame, text=self.i18n.get('button_send'),
            fg_color="#28a745", hover_color="#218838",
            command=self._start_sending
        )
        self.send_btn.pack(side="right", padx=3)

        self.cancel_btn = ctk.CTkButton(
            button_frame, text=self.i18n.get('button_cancel'),
            fg_color="#dc3545", hover_color="#c82333",
            command=self._cancel_sending, state="disabled"
        )
        self.cancel_btn.pack(side="right", padx=3)

    def _create_log_section(self):
        """ログ表示セクション"""
        # 進捗バー
        self.progress_bar = ctk.CTkProgressBar(self.main_frame)
        self.progress_bar.pack(fill="x", pady=(0, 5))
        self.progress_bar.set(0)

        # ログラベル
        ctk.CTkLabel(
            self.main_frame, text=self.i18n.get('section_log'),
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w")

        # ログテキストエリア
        self.log_text = ctk.CTkTextbox(self.main_frame, height=160)
        self.log_text.pack(fill="both", expand=True, pady=(0, 5))
        self.log_text.configure(state="disabled")

    # ==================== ファイル選択 ====================

    def _browse_recipient(self):
        """受信者リストファイルを選択"""
        filetypes = [
            ("CSV/Excel", "*.csv *.xlsx"),
            ("CSV", "*.csv"),
            ("Excel", "*.xlsx"),
            ("All", "*.*"),
        ]
        path = filedialog.askopenfilename(
            title=self.i18n.get('filedialog_recipient'), filetypes=filetypes
        )
        if path:
            self._set_entry(self.recipient_entry, path)

    def _browse_template(self):
        """テンプレートファイルを選択"""
        filetypes = [("Text", "*.txt"), ("All", "*.*")]
        path = filedialog.askopenfilename(
            title=self.i18n.get('filedialog_template'), filetypes=filetypes
        )
        if path:
            self._set_entry(self.template_entry, path)

    def _browse_attachments(self):
        """添付ファイルを選択（複数可）"""
        paths = filedialog.askopenfilenames(
            title=self.i18n.get('filedialog_attachment')
        )
        if paths:
            self._set_entry(self.attachments_entry, ", ".join(paths))

    # ==================== ヘルパーメソッド ====================

    def _set_entry(self, entry: ctk.CTkEntry, value: str):
        """エントリーの値を設定"""
        entry.delete(0, "end")
        if value:
            entry.insert(0, value)

    def _log(self, message: str):
        """ログエリアにメッセージを追加"""
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _get_form_state(self) -> Dict[str, str]:
        """現在のフォーム入力値を取得"""
        return {
            'smtp_server': self.smtp_server_entry.get(),
            'port': self.port_entry.get(),
            'email': self.email_entry.get(),
            'password': self.password_entry.get(),
            'display_name': self.display_name_entry.get(),
            'recipient_file': self.recipient_entry.get(),
            'template_file': self.template_entry.get(),
            'attachments': self.attachments_entry.get(),
            'cc': self.cc_entry.get(),
            'bcc': self.bcc_entry.get(),
            'reply_to': self.reply_to_entry.get(),
            'delay': self.delay_entry.get(),
        }

    def _set_form_state(self, state: Dict[str, str]):
        """フォーム入力値を復元"""
        self._set_entry(self.smtp_server_entry, state.get('smtp_server', ''))
        self._set_entry(self.port_entry, state.get('port', ''))
        self._set_entry(self.email_entry, state.get('email', ''))
        self._set_entry(self.password_entry, state.get('password', ''))
        self._set_entry(self.display_name_entry, state.get('display_name', ''))
        self._set_entry(self.recipient_entry, state.get('recipient_file', ''))
        self._set_entry(self.template_entry, state.get('template_file', ''))
        self._set_entry(self.attachments_entry, state.get('attachments', ''))
        self._set_entry(self.cc_entry, state.get('cc', ''))
        self._set_entry(self.bcc_entry, state.get('bcc', ''))
        self._set_entry(self.reply_to_entry, state.get('reply_to', ''))
        self._set_entry(self.delay_entry, state.get('delay', '5'))

    def _validate_inputs(self) -> bool:
        """入力値のバリデーション"""
        if not self.smtp_server_entry.get().strip():
            messagebox.showerror("Error", self.i18n.get('error_no_smtp_server'))
            return False

        try:
            int(self.port_entry.get().strip() or "587")
        except ValueError:
            messagebox.showerror("Error", self.i18n.get('error_invalid_port'))
            return False

        if not self.email_entry.get().strip():
            messagebox.showerror("Error", self.i18n.get('error_no_email'))
            return False

        if not self.password_entry.get():
            messagebox.showerror("Error", self.i18n.get('error_no_password'))
            return False

        recipient_file = self.recipient_entry.get().strip()
        if not recipient_file:
            messagebox.showerror("Error", self.i18n.get('error_no_recipient_file'))
            return False
        if not os.path.exists(recipient_file):
            messagebox.showerror("Error", self.i18n.get('error_file_not_found', recipient_file))
            return False

        template_file = self.template_entry.get().strip()
        if not template_file:
            messagebox.showerror("Error", self.i18n.get('error_no_template_file'))
            return False
        if not os.path.exists(template_file):
            messagebox.showerror("Error", self.i18n.get('error_file_not_found', template_file))
            return False

        return True

    def _create_sender(self) -> EmailBulkSender:
        """EmailBulkSenderインスタンスを作成"""
        return EmailBulkSender(
            email_address=self.email_entry.get().strip(),
            email_password=self.password_entry.get(),
            smtp_server=self.smtp_server_entry.get().strip(),
            smtp_port=int(self.port_entry.get().strip() or "587"),
            sender_display_name=self.display_name_entry.get().strip(),
        )

    def _read_recipients(self) -> List[Dict]:
        """受信者リストを読み込み"""
        sender = self._create_sender()
        return sender.read_recipients(self.recipient_entry.get().strip())

    def _read_template(self):
        """テンプレートを読み込み"""
        sender = self._create_sender()
        return sender.read_email_template(self.template_entry.get().strip())

    def _get_attachments(self) -> Optional[List[str]]:
        """添付ファイルのリストを取得"""
        text = self.attachments_entry.get().strip()
        if not text:
            return None
        return [f.strip() for f in text.split(',') if f.strip()]

    # ==================== ダイアログ表示 ====================

    def _show_recipient_list(self):
        """受信者リスト一覧ダイアログを表示"""
        if not self.recipient_entry.get().strip():
            messagebox.showerror("Error", self.i18n.get('error_no_recipient_file'))
            return

        recipient_file = self.recipient_entry.get().strip()
        if not os.path.exists(recipient_file):
            messagebox.showerror("Error", self.i18n.get('error_file_not_found', recipient_file))
            return

        try:
            recipients = self._read_recipients()
            RecipientListDialog(self, recipients, self.i18n)
        except Exception as e:
            messagebox.showerror("Error", self.i18n.get('error_read_failed', str(e)))

    def _show_preview(self):
        """送信前プレビューダイアログを表示"""
        if not self._validate_inputs():
            return

        try:
            recipients = self._read_recipients()
            subject_template, body_template = self._read_template()

            if not recipients:
                messagebox.showerror("Error", self.i18n.get('error_read_failed', "No recipients"))
                return

            r = recipients[0]
            subject = subject_template.replace('{所属}', r['affiliation']).replace('{氏名}', r['name'])
            body = body_template.replace('{所属}', r['affiliation']).replace('{氏名}', r['name'])

            display_name = self.display_name_entry.get().strip()
            email = self.email_entry.get().strip()
            from_str = f"{display_name} <{email}>" if display_name else email

            preview_data = {
                'from': from_str,
                'to': f"{r['affiliation']} {r['name']} <{r['email']}>",
                'subject': subject,
                'body': body,
                'cc': self.cc_entry.get().strip() or None,
                'bcc': self.bcc_entry.get().strip() or None,
                'reply_to': self.reply_to_entry.get().strip() or None,
                'attachments': self.attachments_entry.get().strip() or None,
            }

            PreviewDialog(self, preview_data, self.i18n)
        except Exception as e:
            messagebox.showerror("Error", self.i18n.get('error_read_failed', str(e)))

    def _show_test_send(self):
        """テスト送信ダイアログを表示"""
        if not self._validate_inputs():
            return
        TestSendDialog(self, self, self.i18n)

    # ==================== 送信処理 ====================

    def _start_sending(self):
        """一括送信を開始"""
        if not self._validate_inputs():
            return

        try:
            recipients = self._read_recipients()
        except Exception as e:
            messagebox.showerror("Error", self.i18n.get('error_read_failed', str(e)))
            return

        if not recipients:
            messagebox.showerror("Error", self.i18n.get('error_read_failed', "No recipients"))
            return

        # 確認ダイアログ
        if not messagebox.askyesno(
            self.i18n.get('confirm_send_title'),
            self.i18n.get('confirm_send_message', len(recipients))
        ):
            return

        # UI状態を送信中に変更
        self._sending = True
        self._cancel_requested = False
        self.send_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")
        self.progress_bar.set(0)

        # ログをクリア
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        # バックグラウンドスレッドで送信
        thread = threading.Thread(
            target=self._do_send, args=(recipients,), daemon=True
        )
        thread.start()

    def _do_send(self, recipients: List[Dict]):
        """送信処理（バックグラウンドスレッド）"""
        try:
            sender = self._create_sender()
            subject_template, body_template = self._read_template()

            cc = self.cc_entry.get().strip() or None
            bcc = self.bcc_entry.get().strip() or None
            reply_to = self.reply_to_entry.get().strip() or None
            attachments = self._get_attachments()

            try:
                delay = float(self.delay_entry.get().strip() or "5")
            except ValueError:
                delay = 5.0

            # SMTP接続
            if sender.smtp_port == 465:
                server = smtplib.SMTP_SSL(sender.smtp_server, sender.smtp_port)
            else:
                server = smtplib.SMTP(sender.smtp_server, sender.smtp_port)
                server.starttls()
            server.login(sender.email_address, sender.email_password)

            success_count = 0
            fail_count = 0
            total = len(recipients)

            for i, recipient in enumerate(recipients, 1):
                if self._cancel_requested:
                    self.after(0, lambda: self._log(self.i18n.get('status_cancelled')))
                    break

                try:
                    msg = sender.create_message(
                        to_email=recipient['email'],
                        to_name=recipient['name'],
                        to_affiliation=recipient['affiliation'],
                        subject_template=subject_template,
                        body_template=body_template,
                        cc=cc, bcc=bcc, reply_to=reply_to,
                        attachments=attachments,
                    )
                    server.send_message(msg)
                    success_count += 1

                    log_msg = self.i18n.get(
                        'send_success', i, total,
                        recipient['affiliation'], recipient['name'], recipient['email']
                    )
                    self.after(0, lambda m=log_msg: self._log(m))

                except Exception as e:
                    fail_count += 1
                    log_msg = self.i18n.get(
                        'send_failed', i, total,
                        recipient['affiliation'], recipient['name'],
                        recipient['email'], str(e)
                    )
                    self.after(0, lambda m=log_msg: self._log(m))

                # 進捗更新
                progress = i / total
                status = self.i18n.get('status_sending', i, total)
                self.after(0, lambda p=progress, s=status: self._update_progress(p, s))

                # 送信間隔
                if i < total and not self._cancel_requested:
                    time.sleep(delay)

            server.quit()

            # 完了メッセージ
            complete_msg = self.i18n.get('status_complete', success_count, fail_count)
            self.after(0, lambda m=complete_msg: self._on_send_complete(m))

        except Exception as e:
            error_msg = self.i18n.get('error_send_failed', str(e))
            self.after(0, lambda m=error_msg: self._on_send_error(m))

    def _update_progress(self, value: float, status: str):
        """進捗を更新"""
        self.progress_bar.set(value)
        self.status_label.configure(text=status)

    def _on_send_complete(self, message: str):
        """送信完了時の処理"""
        self._sending = False
        self.send_btn.configure(state="normal")
        self.cancel_btn.configure(state="disabled")
        self.progress_bar.set(1)
        self.status_label.configure(text=message)
        self._log(message)

    def _on_send_error(self, message: str):
        """送信エラー時の処理"""
        self._sending = False
        self.send_btn.configure(state="normal")
        self.cancel_btn.configure(state="disabled")
        self.status_label.configure(text=message)
        self._log(message)
        messagebox.showerror("Error", message)

    def _cancel_sending(self):
        """送信をキャンセル"""
        self._cancel_requested = True
        self.cancel_btn.configure(state="disabled")

    # ==================== 設定の保存/読み込み ====================

    def _save_config(self):
        """設定をファイルとkeyringに保存"""
        try:
            port = int(self.port_entry.get().strip() or "587")
        except ValueError:
            port = 587

        try:
            delay = float(self.delay_entry.get().strip() or "5")
        except ValueError:
            delay = 5.0

        config = {
            "version": self.config_manager.CONFIG_VERSION,
            "smtp": {
                "server": self.smtp_server_entry.get().strip(),
                "port": port,
            },
            "sender": {
                "email_address": self.email_entry.get().strip(),
                "display_name": self.display_name_entry.get().strip(),
            },
            "files": {
                "csv_file": self.recipient_entry.get().strip(),
                "template_file": self.template_entry.get().strip(),
                "attachments": self._get_attachments() or [],
            },
            "email_options": {
                "cc": self.cc_entry.get().strip(),
                "bcc": self.bcc_entry.get().strip(),
                "reply_to": self.reply_to_entry.get().strip(),
                "send_delay": delay,
            },
            "ui": {
                "language": self.i18n.get_language(),
            },
        }

        if self.config_manager.save_config(config):
            # パスワードをkeyringに保存
            email = self.email_entry.get().strip()
            password = self.password_entry.get()
            if email and password:
                if self.config_manager.save_password(email, password):
                    pass  # 成功
                elif KEYRING_AVAILABLE:
                    self._log(self.i18n.get('status_keyring_not_available'))

            self.status_label.configure(text=self.i18n.get('status_config_saved'))
            self._log(self.i18n.get('status_config_saved'))

    def _load_config(self):
        """設定ファイルとkeyringから読み込み"""
        config = self.config_manager.load_config()
        if not config:
            return

        # SMTP設定
        smtp = config.get('smtp', {})
        self._set_entry(self.smtp_server_entry, smtp.get('server', ''))
        port = smtp.get('port', '')
        self._set_entry(self.port_entry, str(port) if port else '')

        # 送信元設定
        sender_cfg = config.get('sender', {})
        email = sender_cfg.get('email_address', '')
        self._set_entry(self.email_entry, email)
        self._set_entry(self.display_name_entry, sender_cfg.get('display_name', ''))

        # パスワードをkeyringから読み込み
        if email:
            password = self.config_manager.load_password(email)
            if password:
                self._set_entry(self.password_entry, password)

        # ファイル設定
        files = config.get('files', {})
        self._set_entry(self.recipient_entry, files.get('csv_file', ''))
        self._set_entry(self.template_entry, files.get('template_file', ''))
        attachments = files.get('attachments', [])
        if isinstance(attachments, list) and attachments:
            self._set_entry(self.attachments_entry, ', '.join(attachments))

        # オプション設定
        options = config.get('email_options', {})
        self._set_entry(self.cc_entry, options.get('cc', ''))
        self._set_entry(self.bcc_entry, options.get('bcc', ''))
        self._set_entry(self.reply_to_entry, options.get('reply_to', ''))
        delay = options.get('send_delay', 5)
        self._set_entry(self.delay_entry, str(delay) if delay else '5')

        # 言語設定（次回の _change_language で反映）
        ui = config.get('ui', {})
        lang = ui.get('language', '')
        if lang in ['ja', 'en']:
            current_lang = self.i18n.get_language()
            if lang != current_lang:
                self.i18n = GuiI18n(lang)
                self.lang_var.set(lang.upper())
                self._rebuild_ui()

        self.status_label.configure(text=self.i18n.get('status_config_loaded'))

    # ==================== 言語切り替え ====================

    def _change_language(self, lang_code: str):
        """言語を切り替え（UI再構築）"""
        lang = lang_code.lower()
        if lang == self.i18n.get_language():
            return

        self.i18n = GuiI18n(lang)
        self._rebuild_ui()

    def _rebuild_ui(self):
        """現在のフォーム状態を保持したままUIを再構築"""
        state = self._get_form_state()

        # 全ウィジェットを削除
        for widget in self.winfo_children():
            widget.destroy()

        # ウィンドウタイトルを更新
        self.title(self.i18n.get('app_title'))

        # ウィジェットを再作成
        self._create_widgets()

        # フォーム状態を復元
        self._set_form_state(state)


# ==================== エントリーポイント ====================

def main():
    app = EmailBulkSenderApp()
    app.mainloop()


if __name__ == "__main__":
    main()
