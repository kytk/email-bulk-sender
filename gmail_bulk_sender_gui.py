import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import sys
import json
import os
from gmail_bulk_sender import GmailBulkSender
from i18n import get_i18n
from config import ConfigManager

# CustomTkinterの外観設定
ctk.set_appearance_mode("System")  # "System", "Dark", "Light"
ctk.set_default_color_theme("blue")


class GmailBulkSenderGUI(ctk.CTk):
    def __init__(self):
        super().__init__()

        # 設定マネージャーの初期化
        self.config_manager = ConfigManager("gmail")

        # i18n初期化（設定ファイルから言語を読み込む）
        self.config_file = os.path.join(os.path.expanduser("~"), ".gmail_bulk_sender_config.json")
        saved_lang = self.load_language_config()
        self.i18n = get_i18n(saved_lang)

        # ウィンドウ設定
        self.title(self.i18n.get('app_title_gmail'))
        self.geometry("900x700")

        # 変数の初期化
        self.gmail_address = ctk.StringVar()
        self.gmail_password = ctk.StringVar()
        self.sender_display_name = ctk.StringVar()

        self.csv_file = ctk.StringVar()
        self.template_file = ctk.StringVar()
        self.attachment_files = []

        self.cc = ctk.StringVar()
        self.bcc = ctk.StringVar()
        self.reply_to = ctk.StringVar()
        self.send_delay = ctk.StringVar(value="5")

        # タブビューの作成
        self.tabview = ctk.CTkTabview(self, width=850, height=600)
        self.tabview.pack(pady=20, padx=20, fill="both", expand=True)

        # タブの追加
        self.tabview.add(self.i18n.get('tab_basic'))
        self.tabview.add(self.i18n.get('tab_files'))
        self.tabview.add(self.i18n.get('tab_options'))
        self.tabview.add(self.i18n.get('tab_send'))
        self.tabview.add(self.i18n.get('tab_language'))

        # 各タブの作成
        self.create_basic_settings_tab()
        self.create_file_selection_tab()
        self.create_options_tab()
        self.create_send_tab()
        self.create_language_tab()

    def load_language_config(self):
        """設定ファイルから言語設定を読み込む"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    return config.get('language')
        except:
            pass
        return None

    def save_language_config(self, lang):
        """言語設定を設定ファイルに保存"""
        try:
            config = {'language': lang}
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except:
            pass

    def load_config_settings(self):
        """設定ファイルから全ての設定を読み込む"""
        config = self.config_manager.load_config()
        if not config:
            if self.i18n.get_language() == 'ja':
                messagebox.showinfo("情報", "設定ファイルが見つかりません。")
            else:
                messagebox.showinfo("Info", "No configuration file found.")
            return

        # 送信者情報
        sender_config = config.get('sender', {})
        if sender_config.get('email_address'):
            self.gmail_address.set(sender_config['email_address'])
        if sender_config.get('display_name'):
            self.sender_display_name.set(sender_config['display_name'])

        # ファイル設定
        files_config = config.get('files', {})
        if files_config.get('csv_file'):
            self.csv_file.set(files_config['csv_file'])
        if files_config.get('template_file'):
            self.template_file.set(files_config['template_file'])
        if files_config.get('attachments'):
            self.attachment_files = files_config['attachments']
            self.update_attachment_listbox()

        # メールオプション
        email_options = config.get('email_options', {})
        if email_options.get('cc') is not None:
            self.cc.set(email_options['cc'])
        if email_options.get('bcc') is not None:
            self.bcc.set(email_options['bcc'])
        if email_options.get('reply_to') is not None:
            self.reply_to.set(email_options['reply_to'])
        if email_options.get('send_delay') is not None:
            self.send_delay.set(str(email_options['send_delay']))

        if self.i18n.get_language() == 'ja':
            messagebox.showinfo("成功", f"設定を読み込みました\n{self.config_manager.get_config_path()}")
        else:
            messagebox.showinfo("Success", f"Configuration loaded from\n{self.config_manager.get_config_path()}")

    def save_config_settings(self):
        """現在の設定を設定ファイルに保存（パスワードは除く）"""
        config_to_save = {
            "version": self.config_manager.CONFIG_VERSION,
            "sender": {
                "email_address": self.gmail_address.get(),
                "display_name": self.sender_display_name.get()
            },
            "files": {
                "csv_file": self.csv_file.get(),
                "template_file": self.template_file.get(),
                "attachments": self.attachment_files
            },
            "email_options": {
                "cc": self.cc.get(),
                "bcc": self.bcc.get(),
                "reply_to": self.reply_to.get(),
                "send_delay": float(self.send_delay.get()) if self.send_delay.get() else 5
            },
            "ui": {
                "language": self.i18n.get_language()
            }
        }

        if self.config_manager.save_config(config_to_save):
            if self.i18n.get_language() == 'ja':
                messagebox.showinfo("成功", f"設定を保存しました\n{self.config_manager.get_config_path()}")
            else:
                messagebox.showinfo("Success", f"Configuration saved to\n{self.config_manager.get_config_path()}")
        else:
            if self.i18n.get_language() == 'ja':
                messagebox.showerror("エラー", "設定ファイルの保存に失敗しました。")
            else:
                messagebox.showerror("Error", "Failed to save configuration file.")

    def create_basic_settings_tab(self):
        """基本設定タブの作成"""
        tab = self.tabview.tab(self.i18n.get('tab_basic'))

        # スクロール可能なフレーム
        scroll_frame = ctk.CTkScrollableFrame(tab, width=800, height=500)
        scroll_frame.pack(pady=10, padx=10, fill="both", expand=True)

        # Gmail設定
        ctk.CTkLabel(scroll_frame, text=self.i18n.get('smtp_settings'), font=("", 16, "bold")).pack(pady=(10, 5), anchor="w")

        # Gmailアドレス
        ctk.CTkLabel(scroll_frame, text=self.i18n.get('gmail_address')).pack(pady=(10, 0), anchor="w")
        ctk.CTkEntry(scroll_frame, textvariable=self.gmail_address, width=400).pack(pady=5, anchor="w")

        # アプリパスワード
        ctk.CTkLabel(scroll_frame, text=self.i18n.get('gmail_app_password')).pack(pady=(10, 0), anchor="w")
        ctk.CTkEntry(scroll_frame, textvariable=self.gmail_password, width=400, show="*").pack(pady=5, anchor="w")

        # アプリパスワードの説明
        info_text = """
アプリパスワードの取得方法:
1. Googleアカウントの「セキュリティ」設定を開く
2. 「2段階認証プロセス」を有効にする
3. 「アプリパスワード」を選択
4. アプリを選択（メール）、デバイスを選択
5. 生成された16桁のパスワードをここに入力
        """
        ctk.CTkLabel(scroll_frame, text=info_text, text_color="gray", justify="left").pack(pady=(0, 10), anchor="w")

        # 送信元表示名
        ctk.CTkLabel(scroll_frame, text=self.i18n.get('sender_display_name')).pack(pady=(10, 0), anchor="w")
        ctk.CTkEntry(scroll_frame, textvariable=self.sender_display_name, width=400).pack(pady=5, anchor="w")
        ctk.CTkLabel(scroll_frame, text=self.i18n.get('sender_display_name_example'),
                    text_color="gray").pack(pady=(0, 10), anchor="w")

        # 設定の保存・読み込みボタン
        ctk.CTkLabel(scroll_frame, text="設定管理" if self.i18n.get_language() == 'ja' else "Configuration Management",
                    font=("", 16, "bold")).pack(pady=(20, 5), anchor="w")

        config_button_frame = ctk.CTkFrame(scroll_frame)
        config_button_frame.pack(pady=10, fill="x")

        ctk.CTkButton(config_button_frame,
                     text="設定を読み込み" if self.i18n.get_language() == 'ja' else "Load Settings",
                     command=self.load_config_settings,
                     width=200).pack(side="left", padx=5)

        ctk.CTkButton(config_button_frame,
                     text="設定を保存" if self.i18n.get_language() == 'ja' else "Save Settings",
                     command=self.save_config_settings,
                     width=200).pack(side="left", padx=5)

        ctk.CTkLabel(scroll_frame,
                    text="※パスワードは保存されません（セキュリティのため）" if self.i18n.get_language() == 'ja' else "* Password is not saved (for security)",
                    text_color="gray").pack(pady=(5, 10), anchor="w")

    def create_file_selection_tab(self):
        """ファイル選択タブの作成"""
        tab = self.tabview.tab(self.i18n.get('tab_files'))

        # スクロール可能なフレーム
        scroll_frame = ctk.CTkScrollableFrame(tab, width=800, height=500)
        scroll_frame.pack(pady=10, padx=10, fill="both", expand=True)

        # CSVファイル
        ctk.CTkLabel(scroll_frame, text=self.i18n.get('recipient_list'), font=("", 16, "bold")).pack(pady=(10, 5), anchor="w")

        csv_frame = ctk.CTkFrame(scroll_frame)
        csv_frame.pack(pady=10, fill="x")

        ctk.CTkEntry(csv_frame, textvariable=self.csv_file, width=500).pack(side="left", padx=5)
        ctk.CTkButton(csv_frame, text=self.i18n.get('select_file'), command=self.select_csv_file, width=100).pack(side="left", padx=5)

        ctk.CTkLabel(scroll_frame, text=self.i18n.get('csv_description'),
                    text_color="gray").pack(pady=(0, 10), anchor="w")

        # テンプレートファイル
        ctk.CTkLabel(scroll_frame, text=self.i18n.get('email_template'), font=("", 16, "bold")).pack(pady=(20, 5), anchor="w")

        template_frame = ctk.CTkFrame(scroll_frame)
        template_frame.pack(pady=10, fill="x")

        ctk.CTkEntry(template_frame, textvariable=self.template_file, width=500).pack(side="left", padx=5)
        ctk.CTkButton(template_frame, text=self.i18n.get('select_file'), command=self.select_template_file, width=100).pack(side="left", padx=5)

        ctk.CTkLabel(scroll_frame, text=self.i18n.get('template_description'),
                    text_color="gray").pack(pady=(0, 10), anchor="w")

        # 添付ファイル
        ctk.CTkLabel(scroll_frame, text=self.i18n.get('attachments'), font=("", 16, "bold")).pack(pady=(20, 5), anchor="w")

        attachment_btn_frame = ctk.CTkFrame(scroll_frame)
        attachment_btn_frame.pack(pady=10, fill="x")

        ctk.CTkButton(attachment_btn_frame, text=self.i18n.get('add_file'), command=self.add_attachment_files, width=100).pack(side="left", padx=5)
        ctk.CTkButton(attachment_btn_frame, text=self.i18n.get('remove_file'), command=self.clear_attachment_files, width=100).pack(side="left", padx=5)

        # 添付ファイルリスト表示
        self.attachment_listbox = ctk.CTkTextbox(scroll_frame, width=700, height=100)
        self.attachment_listbox.pack(pady=10, fill="x")
        self.attachment_listbox.configure(state="disabled")

    def create_options_tab(self):
        """オプション設定タブの作成"""
        tab = self.tabview.tab(self.i18n.get('tab_options'))

        # スクロール可能なフレーム
        scroll_frame = ctk.CTkScrollableFrame(tab, width=800, height=500)
        scroll_frame.pack(pady=10, padx=10, fill="both", expand=True)

        # CC
        ctk.CTkLabel(scroll_frame, text=self.i18n.get('cc')).pack(pady=(10, 0), anchor="w")
        ctk.CTkEntry(scroll_frame, textvariable=self.cc, width=400).pack(pady=5, anchor="w")

        # BCC
        ctk.CTkLabel(scroll_frame, text=self.i18n.get('bcc')).pack(pady=(10, 0), anchor="w")
        ctk.CTkEntry(scroll_frame, textvariable=self.bcc, width=400).pack(pady=5, anchor="w")

        # Reply-To
        ctk.CTkLabel(scroll_frame, text=self.i18n.get('reply_to')).pack(pady=(10, 0), anchor="w")
        ctk.CTkEntry(scroll_frame, textvariable=self.reply_to, width=400).pack(pady=5, anchor="w")

        # 送信間隔
        ctk.CTkLabel(scroll_frame, text=self.i18n.get('send_delay')).pack(pady=(20, 0), anchor="w")
        ctk.CTkEntry(scroll_frame, textvariable=self.send_delay, width=400).pack(pady=5, anchor="w")
        ctk.CTkLabel(scroll_frame, text=self.i18n.get('send_delay_hint'),
                    text_color="gray").pack(pady=(0, 10), anchor="w")

    def create_send_tab(self):
        """送信実行タブの作成"""
        tab = self.tabview.tab(self.i18n.get('tab_send'))

        # プレビューボタン
        preview_frame = ctk.CTkFrame(tab)
        preview_frame.pack(pady=10, padx=10, fill="x")

        ctk.CTkButton(preview_frame, text=self.i18n.get('preview_button'), command=self.preview_content,
                     width=200, height=40, font=("", 14, "bold")).pack(pady=10)

        # 送信ボタン
        send_frame = ctk.CTkFrame(tab)
        send_frame.pack(pady=10, padx=10, fill="x")

        self.send_button = ctk.CTkButton(send_frame, text=self.i18n.get('send_button'), command=self.start_sending,
                                         width=200, height=40, font=("", 14, "bold"), fg_color="green")
        self.send_button.pack(pady=10)

        # プログレスバー
        self.progress_bar = ctk.CTkProgressBar(tab, width=800)
        self.progress_bar.pack(pady=10, padx=10)
        self.progress_bar.set(0)

        # ログ表示
        ctk.CTkLabel(tab, text=self.i18n.get('log') + ":", font=("", 14, "bold")).pack(pady=(10, 5), anchor="w", padx=10)

        self.log_textbox = ctk.CTkTextbox(tab, width=850, height=300)
        self.log_textbox.pack(pady=10, padx=10, fill="both", expand=True)

    def create_language_tab(self):
        """言語設定タブの作成"""
        tab = self.tabview.tab(self.i18n.get('tab_language'))

        # スクロール可能なフレーム
        scroll_frame = ctk.CTkScrollableFrame(tab, width=800, height=500)
        scroll_frame.pack(pady=10, padx=10, fill="both", expand=True)

        # 言語設定
        ctk.CTkLabel(scroll_frame, text=self.i18n.get('language_setting'), font=("", 16, "bold")).pack(pady=(10, 5), anchor="w")

        # 言語選択
        ctk.CTkLabel(scroll_frame, text=self.i18n.get('language')).pack(pady=(10, 0), anchor="w")

        self.language_var = ctk.StringVar(value=self.i18n.get_language())
        language_options = ["ja", "en"]
        language_menu = ctk.CTkOptionMenu(scroll_frame, variable=self.language_var,
                                          values=language_options,
                                          command=self.on_language_changed,
                                          width=200)
        language_menu.pack(pady=5, anchor="w")

        # 注意メッセージ
        ctk.CTkLabel(scroll_frame, text=self.i18n.get('language_note'),
                    text_color="gray", wraplength=700).pack(pady=(20, 10), anchor="w")

    def on_language_changed(self, choice):
        """言語変更時のハンドラ"""
        self.save_language_config(choice)
        messagebox.showinfo(
            self.i18n.get('info'),
            self.i18n.get('language_note')
        )

    def select_csv_file(self):
        """CSVまたはExcelファイル選択"""
        filename = filedialog.askopenfilename(
            title=self.i18n.get('select_csv'),
            filetypes=[
                ("CSV/Excel files", "*.csv;*.xlsx"),
                ("CSV files", "*.csv"),
                ("Excel files", "*.xlsx"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.csv_file.set(filename)

    def select_template_file(self):
        """テンプレートファイル選択"""
        filename = filedialog.askopenfilename(
            title=self.i18n.get('select_template'),
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if filename:
            self.template_file.set(filename)

    def add_attachment_files(self):
        """添付ファイル追加"""
        filenames = filedialog.askopenfilenames(
            title=self.i18n.get('select_attachment'),
            filetypes=[("All files", "*.*")]
        )
        if filenames:
            self.attachment_files.extend(filenames)
            self.update_attachment_listbox()

    def clear_attachment_files(self):
        """添付ファイルクリア"""
        self.attachment_files.clear()
        self.update_attachment_listbox()

    def update_attachment_listbox(self):
        """添付ファイルリスト更新"""
        self.attachment_listbox.configure(state="normal")
        self.attachment_listbox.delete("1.0", "end")
        if self.attachment_files:
            for i, file in enumerate(self.attachment_files, 1):
                self.attachment_listbox.insert("end", f"{i}. {file}\n")
        else:
            self.attachment_listbox.insert("end", self.i18n.get('no_attachments'))
        self.attachment_listbox.configure(state="disabled")

    def log(self, message):
        """ログに追加"""
        self.log_textbox.insert("end", message + "\n")
        self.log_textbox.see("end")
        self.update()

    def validate_inputs(self):
        """入力検証"""
        if not self.gmail_address.get():
            messagebox.showerror(self.i18n.get('error'), self.i18n.get('error_missing_email'))
            return False
        if not self.gmail_password.get():
            messagebox.showerror(self.i18n.get('error'), self.i18n.get('error_missing_password'))
            return False
        if not self.csv_file.get():
            messagebox.showerror(self.i18n.get('error'), self.i18n.get('error_missing_csv'))
            return False
        if not self.template_file.get():
            messagebox.showerror(self.i18n.get('error'), self.i18n.get('error_missing_template'))
            return False

        try:
            float(self.send_delay.get())
        except ValueError:
            messagebox.showerror(self.i18n.get('error'), self.i18n.get('error_invalid_delay'))
            return False

        return True

    def preview_content(self):
        """送信内容のプレビュー"""
        if not self.validate_inputs():
            return

        try:
            sender = GmailBulkSender(
                self.gmail_address.get(),
                self.gmail_password.get(),
                self.sender_display_name.get()
            )

            # 受信者リストを読み込み
            recipients = sender.read_recipients(self.csv_file.get())
            subject_template, body_template = sender.read_email_template(self.template_file.get())

            # プレビュー内容を作成
            preview_text = f"""
=== 送信内容プレビュー ===

【基本情報】
送信元: {self.sender_display_name.get() + ' <' + self.gmail_address.get() + '>' if self.sender_display_name.get() else self.gmail_address.get()}
SMTPサーバー: smtp.gmail.com:587
送信先: {len(recipients)}件

【件名】
{subject_template}

【本文サンプル】（1件目の受信者で展開）
{body_template.replace('{企業}', recipients[0]['company']).replace('{氏名}', recipients[0]['name'])}

【オプション】
CC: {self.cc.get() if self.cc.get() else 'なし'}
BCC: {self.bcc.get() if self.bcc.get() else 'なし'}
Reply-To: {self.reply_to.get() if self.reply_to.get() else 'なし'}
添付ファイル: {len(self.attachment_files)}件
送信間隔: {self.send_delay.get()}秒

【送信先リスト】
"""
            for i, recipient in enumerate(recipients, 1):
                preview_text += f"{i}. {recipient['company']} {recipient['name']} ({recipient['email']})\n"

            # プレビューウィンドウを表示
            preview_window = ctk.CTkToplevel(self)
            preview_window.title(self.i18n.get('preview_title'))
            preview_window.geometry("800x600")

            preview_textbox = ctk.CTkTextbox(preview_window, width=780, height=550)
            preview_textbox.pack(pady=10, padx=10, fill="both", expand=True)
            preview_textbox.insert("1.0", preview_text)
            preview_textbox.configure(state="disabled")

        except Exception as e:
            messagebox.showerror(self.i18n.get('error'), f"{self.i18n.get('error_send_failed')}:\n{str(e)}")

    def start_sending(self):
        """送信開始"""
        if not self.validate_inputs():
            return

        # 確認ダイアログ
        result = messagebox.askyesno(
            self.i18n.get('confirm'),
            self.i18n.get('confirm_send')
        )
        if not result:
            return

        # ボタンを無効化
        self.send_button.configure(state="disabled")

        # ログクリア
        self.log_textbox.delete("1.0", "end")
        self.progress_bar.set(0)

        # 別スレッドで送信処理を実行
        thread = threading.Thread(target=self.send_emails, daemon=True)
        thread.start()

    def send_emails(self):
        """メール送信処理（別スレッドで実行）"""
        try:
            self.log(f"=== {self.i18n.get('sending')} ===")

            sender = GmailBulkSender(
                self.gmail_address.get(),
                self.gmail_password.get(),
                self.sender_display_name.get()
            )

            # 受信者リストとテンプレートを読み込み
            recipients = sender.read_recipients(self.csv_file.get())
            subject_template, body_template = sender.read_email_template(self.template_file.get())

            self.log(self.i18n.get('preview_recipients', len(recipients)))
            self.log(self.i18n.get('preview_subject', subject_template))

            # オプション設定
            cc = self.cc.get() if self.cc.get() else None
            bcc = self.bcc.get() if self.bcc.get() else None
            reply_to = self.reply_to.get() if self.reply_to.get() else None
            attachments = self.attachment_files if self.attachment_files else None
            delay = float(self.send_delay.get())

            # SMTP接続して送信
            import smtplib
            import time

            try:
                # SMTP接続（Gmail）
                server = smtplib.SMTP(sender.smtp_server, sender.smtp_port)
                server.starttls()
                server.login(self.gmail_address.get(), self.gmail_password.get())
                self.log("Gmail SMTP接続成功")

                success_count = 0
                fail_count = 0

                for i, recipient in enumerate(recipients, 1):
                    try:
                        msg = sender.create_message(
                            to_email=recipient['email'],
                            to_name=recipient['name'],
                            to_company=recipient['company'],
                            subject_template=subject_template,
                            body_template=body_template,
                            cc=cc,
                            bcc=bcc,
                            reply_to=reply_to,
                            attachments=attachments
                        )

                        server.send_message(msg)
                        success_count += 1
                        self.log(self.i18n.get('send_success', i, len(recipients), recipient['company'], recipient['name'], recipient['email']))

                        # プログレスバー更新
                        progress = i / len(recipients)
                        self.progress_bar.set(progress)

                        # 送信間隔
                        if i < len(recipients):
                            time.sleep(delay)

                    except Exception as e:
                        fail_count += 1
                        self.log(self.i18n.get('send_failed', i, len(recipients), recipient['company'], recipient['name'], recipient['email'], str(e)))

                server.quit()
                self.log(f"\n=== {self.i18n.get('send_complete', success_count, fail_count)} ===")

                messagebox.showinfo(self.i18n.get('success'), self.i18n.get('send_complete', success_count, fail_count))

            except Exception as e:
                self.log(f"Gmail SMTP {self.i18n.get('error')}: {e}")
                messagebox.showerror(self.i18n.get('error'), f"Gmail SMTP {self.i18n.get('error')}:\n{str(e)}\n\nGmailアプリパスワードを確認してください。")

        except Exception as e:
            self.log(f"{self.i18n.get('error')}: {e}")
            messagebox.showerror(self.i18n.get('error'), f"{self.i18n.get('error_send_failed')}:\n{str(e)}")

        finally:
            # ボタンを有効化
            self.send_button.configure(state="normal")


def main():
    app = GmailBulkSenderGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
