; =========================================================
; Instalador de Control de Participación
; Requiere que exista dist\ControlParticipacion.exe
; =========================================================

#define MyAppName "Control de Participación"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Control de Participación"
#define MyAppExeName "ControlParticipacion.exe"

[Setup]
AppId={{B3D3A1C0-7E4B-4D9B-9C47-5E8E5D44A2F1}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\ControlParticipacion
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=instalador
OutputBaseFilename=Instalador_ControlParticipacion
Compression=lzma
SolidCompression=yes
WizardStyle=modern
Uninstallable=yes

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Files]
Source: "dist\ControlParticipacion.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir {#MyAppName}"; Flags: nowait postinstall skipifsilent
