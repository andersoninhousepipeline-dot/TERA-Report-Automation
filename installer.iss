; ============================================================
; TERA Report Generator - Inno Setup Installer Script
; Requires Inno Setup 6: https://jrsoftware.org/isdl.php
; Build: ISCC.exe installer.iss
; ============================================================

#define AppName      "TERA Report Generator"
#define AppVersion   "1.0"
#define AppPublisher "Anderson Diagnostics & Labs"
#define AppExeName   "TERA Report.exe"
#define AppDir       "dist\TERA Report"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL=https://www.andersondiagnostics.com
DefaultDirName={autopf}\Anderson Diagnostics\TERA Report
DefaultGroupName=Anderson Diagnostics
AllowNoIcons=yes
OutputDir=Output
OutputBaseFilename=TERA_Report_Setup
SetupIconFile=tera_icon.ico
UninstallDisplayIcon={app}\{#AppExeName}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
WizardSmallImageFile=tera_icon.ico
MinVersion=10.0
PrivilegesRequired=admin
DisableProgramGroupPage=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"

[Files]
; Main application (all files from PyInstaller output)
Source: "{#AppDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; Fonts — install to Windows Fonts folder
Source: "fonts\Calibri.ttf";          DestDir: "{fonts}"; FontInstall: "Calibri";           Flags: onlyifdoesntexist uninsneveruninstall
Source: "fonts\Calibri-Bold.ttf";     DestDir: "{fonts}"; FontInstall: "Calibri Bold";       Flags: onlyifdoesntexist uninsneveruninstall
Source: "fonts\SegoeUI.ttf";          DestDir: "{fonts}"; FontInstall: "Segoe UI";            Flags: onlyifdoesntexist uninsneveruninstall
Source: "fonts\SegoeUI-Bold.ttf";     DestDir: "{fonts}"; FontInstall: "Segoe UI Bold";       Flags: onlyifdoesntexist uninsneveruninstall
Source: "fonts\GillSansMT-Bold.ttf";  DestDir: "{fonts}"; FontInstall: "Gill Sans MT Bold";   Flags: onlyifdoesntexist uninsneveruninstall
Source: "fonts\DengXian.ttf";         DestDir: "{fonts}"; FontInstall: "DengXian";            Flags: onlyifdoesntexist uninsneveruninstall
Source: "fonts\DengXian_Bold.ttf";    DestDir: "{fonts}"; FontInstall: "DengXian Bold";       Flags: onlyifdoesntexist uninsneveruninstall

[Icons]
Name: "{group}\{#AppName}";        Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\tera_icon.ico"
Name: "{commondesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\tera_icon.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Launch {#AppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

[Messages]
FinishedLabel=Setup has finished installing {#AppName} on your computer.%nYou can launch the application from the desktop shortcut or Start Menu.
