; Script Inno Setup – compilato dal workflow con:
;   ISCC /DAppVersion=1.2.3 /DSourceDir=<cartella publish> /O<cartella output> CostsManager.iss

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif
#ifndef SourceDir
  #error "SourceDir non definito: passa /DSourceDir=<cartella publish>"
#endif

#define AppName "Gestione Conti"
#define AppExe  "CostsManager.exe"

[Setup]
; GUID univoco dell'app: NON cambiarlo tra una versione e l'altra,
; altrimenti Windows tratterà ogni release come un programma diverso.
AppId={{B7E2C4A1-5D3F-4E8A-9C61-2F0A7D84E3B5}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=Matteo Menichetti
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
OutputBaseFilename=GestioneConti-{#AppVersion}-win-x64-setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
; Installazione per-utente di default (nessun prompt UAC); l'utente può scegliere "per tutti"
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
UninstallDisplayIcon={app}\{#AppExe}

[Languages]
Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExe}"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: nowait postinstall skipifsilent
