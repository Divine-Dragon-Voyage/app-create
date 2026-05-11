#define MyAppName "App Create Launcher"
#ifndef MyAppVersion
  #define MyAppVersion "1.0.0"
#endif
#ifndef SourceDir
  #define SourceDir GetEnv("APP_CREATE_LAUNCHER_SOURCE")
#endif
#ifndef OutputDir
  #define OutputDir GetEnv("APP_CREATE_LAUNCHER_OUTPUT")
#endif
#ifndef OutputBaseFilename
  #define OutputBaseFilename "AppCreateLauncherSetup"
#endif

#if SourceDir == ""
  #error APP_CREATE_LAUNCHER_SOURCE (or /DSourceDir) is required.
#endif

#if OutputDir == ""
  #define OutputDir "."
#endif

[Setup]
AppId={{D4E312E6-2CE2-4F35-B328-70AE5B175A66}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
DefaultDirName={autopf}\App Create Launcher
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
DisableProgramGroupPage=yes
PrivilegesRequired=admin
Compression=lzma
SolidCompression=yes
WizardStyle=modern
OutputDir={#OutputDir}
OutputBaseFilename={#OutputBaseFilename}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: checkedonce

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\AppCreateLauncher.cmd"; WorkingDir: "{app}"
Name: "{autodesktop}\App Create"; Filename: "{app}\AppCreateLauncher.cmd"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\AppCreateLauncher.cmd"; Description: "Launch App Create"; Flags: nowait postinstall skipifsilent
