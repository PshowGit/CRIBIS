; Installer Inno Setup per CRIBIS Export.
; Prerequisito: l'exe deve essere già stato creato con PyInstaller (dist\cribis_export.exe).
; Compilazione: ISCC.exe installer\cribis_export.iss   (opzionale: /DAppVersion=1.0.2)

#ifndef AppVersion
  #define AppVersion "1.0.2"
#endif
#define AppName "CRIBIS Export"
#define AppExe "cribis_export.exe"

[Setup]
AppId={{0A69DD34-53F9-4ED6-ABE5-59AA901FE595}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=Professional Show
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
; Installazione per utente (nessun permesso di amministratore): l'app salva config, log e
; risultati accanto all'eseguibile, quindi la cartella deve essere scrivibile (no Program Files).
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
UninstallDisplayIcon={app}\{#AppExe}
OutputDir=Output
OutputBaseFilename=CRIBIS_Export_Setup_{#AppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=yes

[Languages]
Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "..\dist\{#AppExe}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{group}\Disinstalla {#AppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExe}"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Rimuove le credenziali e il registro; il file Excel dei risultati viene lasciato.
Type: files; Name: "{app}\config.json"
Type: files; Name: "{app}\cribis_export.log"
