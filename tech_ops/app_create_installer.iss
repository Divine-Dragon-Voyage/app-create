#define MyAppName "App Create"
#ifndef MyAppVersion
  #define MyAppVersion "1.0.0"
#endif
#ifndef MyAppPublisher
  #define MyAppPublisher "App Create Team"
#endif
#ifndef MyAppURL
  #define MyAppURL "https://play.google.com/console"
#endif
#ifndef SourceDir
  #define SourceDir GetEnv("APP_CREATE_INSTALLER_SOURCE")
#endif
#ifndef OutputDir
  #define OutputDir GetEnv("APP_CREATE_INSTALLER_OUTPUT")
#endif
#ifndef OutputBaseFilename
  #define OutputBaseFilename "AppCreateSetup"
#endif

#if SourceDir == ""
  #error APP_CREATE_INSTALLER_SOURCE (or /DSourceDir) is required.
#endif

#if OutputDir == ""
  #define OutputDir "."
#endif

[Setup]
AppId={{C78DADAC-1F52-4C5A-9FCA-8F69B7E3E8D1}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\App Create
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
Name: "{group}\{#MyAppName}"; Filename: "{app}\user_ops\run_windows.cmd"; WorkingDir: "{app}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\user_ops\run_windows.cmd"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\user_ops\run_windows.cmd"; Description: "{cm:LaunchProgram,App Create}"; Flags: nowait postinstall skipifsilent
