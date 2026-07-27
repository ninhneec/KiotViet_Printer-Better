[Setup]
AppId={{8E1A7C3D-2D91-4E37-8B54-123456789ABC}
AppName=KiotViet Printer Better
AppVersion=0.1.0-beta.36
AppPublisher=ninhneec

DefaultDirName={localappdata}\Programs\KiotViet Printer Better
DefaultGroupName=KiotViet Printer Better
DisableProgramGroupPage=yes
PrivilegesRequired=lowest

OutputDir=installer_output
OutputBaseFilename=KiotViet-Printer-Better-Setup

Compression=lzma
SolidCompression=yes
WizardStyle=modern

; Cho phép cài đè / update
AppMutex=KiotVietLabelPrinterProV2Mutex
CloseApplications=force
RestartApplications=no

; Gỡ bản cũ theo cùng AppId trước khi cài bản mới
UninstallDisplayIcon={app}\KiotViet Label Printer Pro V2.exe

[Files]
Source: "publish\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\KiotViet Printer Better"; Filename: "{app}\KiotViet Label Printer Pro V2.exe"; WorkingDir: "{app}"
Name: "{autodesktop}\KiotViet Printer Better"; Filename: "{app}\KiotViet Label Printer Pro V2.exe"; WorkingDir: "{app}"

[Run]
Filename: "{app}\KiotViet Label Printer Pro V2.exe"; Description: "Mở KiotViet Printer Better"; Flags: nowait postinstall skipifsilent
