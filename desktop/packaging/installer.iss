; Inno Setup 6 — instalador por usuario de "Conversor PDF a Excel".
; Se compila con:  ISCC.exe packaging\installer.iss   (desde desktop\, después de PyInstaller + prune_qt.py)
; Instala sin UAC en %LOCALAPPDATA%\Programs (el diálogo permite elegir "para todos los usuarios").
; Misma AppId en todas las versiones => actualizar es correr el setup nuevo encima del viejo.

#define AppName "Conversor PDF a Excel"
#define AppVersion "0.1.0"
; ^ la reescribe packaging/stamp_version.py a partir del tag (v1.2.3 -> "1.2.3")
#define AppExe "ConversorPDFaExcel.exe"
#define AppPublisher "Equipo Conversor PDF a Excel"
#define AppURL "https://github.com/juanmaserrate/conversor-pdf-a-excel"
; GUID fijo: NO cambiarlo nunca (identifica la app para actualizar/desinstalar). La doble llave escapa la llave.
#define AppId "{{DC2B38EF-8D6D-4B94-918C-A6BD0B627A32}"

[Setup]
AppId={#AppId}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}/issues
AppUpdatesURL={#AppURL}/releases
DefaultDirName={autopf}\{#AppName}
DisableProgramGroupPage=yes
DisableDirPage=auto
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
UsedUserAreasWarning=no
MinVersion=10.0
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=..\dist\installer
OutputBaseFilename=ConversorPDFaExcel-{#AppVersion}-setup-win64
Compression=lzma2/ultra64
SolidCompression=yes
LZMAUseSeparateProcess=yes
WizardStyle=modern
SetupIconFile=..\resources\icons\app.ico
UninstallDisplayIcon={app}\{#AppExe}
UninstallDisplayName={#AppName}
CloseApplications=yes
CloseApplicationsFilter=*.exe,*.dll,*.pyd
RestartApplications=no
ShowLanguageDialog=no

[Languages]
Name: "es"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear un acceso directo en el Escritorio"; GroupDescription: "Accesos directos:"
Name: "sendto"; Description: "Agregar al menú «Enviar a» del botón derecho (convertir PDFs desde el Explorador)"; GroupDescription: "Integración con Windows:"

[InstallDelete]
; Limpia la carpeta interna de la versión anterior antes de copiar: PyInstaller cambia nombres de .pyd/.dll
; entre versiones y los restos viejos pueden hacer que la app cargue módulos mezclados.
Type: filesandordirs; Name: "{app}\_internal"

[Files]
Source: "..\dist\ConversorPDFaExcel\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\LICENSE"; DestDir: "{app}"; DestName: "LICENSE.txt"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExe}"; Comment: "Convertí PDFs a Excel"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon
Name: "{usersendto}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: sendto

[Run]
Filename: "{app}\{#AppExe}"; Description: "Abrir {#AppName}"; Flags: nowait postinstall skipifsilent
