; Inno Setup 脚本 —— 图片表格转 Excel 助手（安装版）
; 版本号通过命令行注入：iscc /DMyAppVersion=1.2.0 installer.iss
; 输出目录通过 /O 指定；资产名保持 ASCII，中文仅用于显示名

#ifndef MyAppVersion
  #define MyAppVersion "0.0.0"
#endif

#define MyAppName "图片表格转 Excel 助手"
#define MyAppPublisher "jianRY"
#define MyAppURL "https://github.com/jianRY"
#define MyAppExeName "图片表格转Excel助手.exe"

[Setup]
; AppId 全局唯一，勿与其他项目重复（否则会被当成同一软件的升级）
AppId={{7C4E9B21-8D53-4A6F-9E18-3B5F2C7D0A64}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
DefaultDirName={autopf}\Img2ExcelTool
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=dist\installer
OutputBaseFilename=Img2ExcelTool_v{#MyAppVersion}_setup
SetupIconFile=assets\app.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName} {#MyAppVersion}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible
DisableDirPage=no

[Languages]
Name: "chinese"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式(&D)"; GroupDescription: "附加任务："

[Files]
; 安装包负载 = onedir 目录版产物（启动比单文件版快得多）
Source: "dist\onedir\图片表格转Excel助手\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\卸载 {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "立即运行 {#MyAppName}"; Flags: nowait postinstall skipifsilent
