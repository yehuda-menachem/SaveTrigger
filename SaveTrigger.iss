; ============================================================
;  SaveTrigger — Inno Setup Script
;  Build steps (run before compiling this script):
;
;    dotnet publish -c Release -r win-x64 --self-contained true ^
;      -p:PublishSingleFile=false -o publish\
;
;  Then open this file in Inno Setup Compiler and press Compile.
;
;  NOTE: Convert logo.svg to logo.ico (e.g. with IcoFX or online tools)
;        and place it at the project root before compiling.
; ============================================================

#define AppName      "SaveTrigger"
#define AppVersion   "1.0.0"
#define AppPublisher "Yehuda Menachem"
#define AppExeName   "SaveTrigger.exe"
#define SourceDir    "publish"

[Setup]
AppId={{F3A1B2C4-7E8D-4A9F-B3C1-2D5E6F7A8B9C}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisherURL=
AppPublisher={#AppPublisher}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
OutputDir=installer
OutputBaseFilename=SaveTrigger-Setup-{#AppVersion}
; Uncomment after converting logo.svg to logo.ico:
; SetupIconFile=logo.ico
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayName={#AppName}
; Uncomment after creating logo.ico:
; UninstallDisplayIcon={app}\{#AppExeName}
CloseApplications=yes
CloseApplicationsFilter=*{#AppExeName}
RestartApplications=no

[Languages]
Name: "hebrew"; MessagesFile: "compiler:Languages\Hebrew.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "startup"; Description: "הפעל {#AppName} אוטומטית עם Windows"; GroupDescription: "הגדרות נוספות:"; Flags: unchecked

[Files]
; All published output (self-contained: includes .NET runtime DLLs)
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Overwrite appsettings.json only on first install — preserve user edits on upgrades
Source: "{#SourceDir}\appsettings.json"; DestDir: "{app}"; Flags: ignoreversion onlyifdoesntexist

[Icons]
; Start menu
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{group}\הסר התקנה"; Filename: "{uninstallexe}"

[Registry]
; Run at Windows startup (optional task)
Root: HKCU; Subkey: "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"; \
  ValueType: string; ValueName: "{#AppName}"; \
  ValueData: """{app}\{#AppExeName}"""; \
  Flags: uninsdeletevalue; Tasks: startup

[Run]
; Launch app after install (without waiting)
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,{#AppName}}"; \
  Flags: nowait postinstall skipifsilent

[UninstallRun]
; Kill the tray process gracefully before uninstall
Filename: "taskkill.exe"; Parameters: "/f /im {#AppExeName}"; \
  Flags: runhidden; RunOnceId: "KillSaveTrigger"

[Code]
// Remove the startup registry key if the user unchecked the task
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
  begin
    RegDeleteValue(HKCU,
      'SOFTWARE\Microsoft\Windows\CurrentVersion\Run',
      '{#AppName}');
  end;
end;
