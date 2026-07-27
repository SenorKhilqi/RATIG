[Setup]
; App Information
AppName=RATIG (Reaction Timer Fatigue)
AppVersion=1.0
AppPublisher=HSE Dept
DefaultDirName={autopf}\RATIG
DefaultGroupName=RATIG
OutputDir=release
OutputBaseFilename=RATIG_Setup
SetupIconFile=release\RATIG_Portable\assets\ratig.ico
Compression=lzma
SolidCompression=yes
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=lowest

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Main Executable
Source: "release\RATIG_Portable\RATIG.exe"; DestDir: "{app}"; Flags: ignoreversion

; Folders to include
Source: "release\RATIG_Portable\assets\*"; DestDir: "{app}\assets"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "release\RATIG_Portable\config\*"; DestDir: "{app}\config"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "release\RATIG_Portable\data\*"; DestDir: "{app}\data"; Flags: ignoreversion recursesubdirs createallsubdirs

; Empty Folders to create
[Dirs]
Name: "{app}\backups"
Name: "{app}\logs"
Name: "{app}\reports"

[Icons]
Name: "{group}\RATIG"; Filename: "{app}\RATIG.exe"
Name: "{autodesktop}\RATIG"; Filename: "{app}\RATIG.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\RATIG.exe"; Description: "{cm:LaunchProgram,RATIG}"; Flags: nowait postinstall skipifsilent
