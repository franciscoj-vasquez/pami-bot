; ============================================================
;  Inno Setup Script — PAMI Bot
;  Requiere Inno Setup 6+: https://jrsoftware.org/isdl.php
; ============================================================

#define MyAppName      "KINETICA"
; Versión inyectada por build.bat vía /DMyAppVersion=X.Y.Z
; Si compilás manualmente sin build.bat, pasá el flag: /DMyAppVersion="1.0.0"
#ifndef MyAppVersion
  #define MyAppVersion "0.0.0"
#endif
#define MyAppPublisher "Francisco Vasquez"
#define MyAppExeName   "KINETICA.exe"
#define DistDir        "..\dist\KINETICA"

[Setup]
AppId={{F3A2C1B0-9D8E-4F7A-B6C5-1234567890AB}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} v{#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DisableProgramGroupPage=yes
LicenseFile=license.rtf
OutputDir=..\dist
OutputBaseFilename=KINETICA_setup
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayName={#MyAppName}
UninstallDisplayIcon={app}\{#MyAppExeName}
MinVersion=10.0

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Messages]
; Textos en español
WelcomeLabel1=Bienvenido al instalador de [name]
WelcomeLabel2=Este asistente instalar[aacute] [name/ver] en su equipo.%n%nSe recomienda cerrar todas las aplicaciones antes de continuar.
FinishedHeadingLabel=Instalaci[oacute]n completada
FinishedLabel=[name] fue instalado correctamente en su equipo.

[Tasks]
Name: "desktopicon"; Description: "Crear acceso directo en el Escritorio"; GroupDescription: "Accesos directos:";

[Files]
Source: "{#DistDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}";  Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Descarga Chromium al finalizar la instalacion (requiere internet, ~300 MB)
Filename: "{app}\bot_runner.exe"; \
  Parameters: "--install-browsers"; \
  WorkingDir: "{app}"; \
  StatusMsg: "Instalando componentes del navegador (requiere internet, puede tardar varios minutos)..."; \
  Flags: runhidden waituntilterminated; \
  Description: "Descargar componentes del navegador (requerido para el funcionamiento del bot)";

; Opcion para iniciar la app al finalizar el instalador
Filename: "{app}\{#MyAppExeName}"; \
  Description: "Iniciar {#MyAppName}"; \
  Flags: nowait postinstall skipifsilent;

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then begin
    SaveStringToFile(ExpandConstant('{app}\version.txt'), '{#MyAppVersion}', False);
    if not DirExists(ExpandConstant('{localappdata}\ms-playwright')) then
      MsgBox(
        'Advertencia: no se pudo descargar el navegador.' + #13#10 +
        'Verificá tu conexión a internet y reinstalá la aplicación.',
        mbError, MB_OK);
  end;
end;

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
