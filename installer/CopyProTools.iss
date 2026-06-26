#ifndef MyAppVersion
  #define MyAppVersion "0.0.0"
#endif

#define MyAppName "CopyPro Tools"
#define MyAppPublisher "CopyPro"
#define MyAppExeName "CopyPro Tools.exe"

[Setup]
AppId={{4D8C6459-D0A3-4E9D-9E42-5F1A12B66231}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\CopyPro Tools
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=..\Output
OutputBaseFilename=CopyPro-Tools-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
SetupIconFile=..\icon_transparent.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Files]
Source: "..\dist\CopyPro Tools\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autodesktop}\CopyPro Tools"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
Name: "{userprograms}\CopyPro Tools"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Paleisti CopyPro Tools"; Flags: nowait postinstall skipifsilent
