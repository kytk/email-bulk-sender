; Inno Setup Script for メール一括送信ツール (Email Bulk Sender)
;
; 使用方法:
;   1. PyInstaller でビルド:
;      pyinstaller --onedir --windowed --name EmailBulkSender email_bulk_sender_gui.py
;
;   2. Inno Setup でこのスクリプトをコンパイル:
;      iscc installer\setup.iss
;
;   ※ 出力先: installer\Output\EmailBulkSender_Setup.exe

#define MyAppName "メール一括送信ツール"
#define MyAppNameEn "Email Bulk Sender"
#define MyAppVersion "3.0"
#define MyAppPublisher "kytk"
#define MyAppURL "https://github.com/kytk/email-bulk-sender"
#define MyAppExeName "EmailBulkSender.exe"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppNameEn}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=Output
OutputBaseFilename=EmailBulkSender_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; PyInstaller の出力ディレクトリからコピー
Source: "..\dist\EmailBulkSender\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
