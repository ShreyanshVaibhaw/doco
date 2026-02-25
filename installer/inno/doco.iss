[Setup]
AppId={{9D9AA8FA-4BBF-4CF8-9E03-4440F90581C9}
AppName=Doco
AppVersion=0.1.0
AppPublisher=Doco Project
DefaultDirName={autopf}\Doco
DefaultGroupName=Doco
DisableProgramGroupPage=yes
OutputDir=..\..\dist
OutputBaseFilename=Doco-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64
ArchitecturesAllowed=x64compatible
ChangesAssociations=yes
ChangesEnvironment=yes
UninstallDisplayIcon={app}\doco.exe

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; Flags: unchecked
Name: "addtopath"; Description: "Add Doco install directory to PATH"; Flags: unchecked

[Files]
Source: "..\..\target\release\doco.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\..\pdfium.dll"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "..\..\doco.ini"; DestDir: "{app}"; Flags: onlyifdoesntexist

[Icons]
Name: "{group}\Doco"; Filename: "{app}\doco.exe"
Name: "{commondesktop}\Doco"; Filename: "{app}\doco.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\doco.exe"; Description: "Launch Doco"; Flags: nowait postinstall skipifsilent

[Registry]
Root: HKCR; Subkey: ".docx"; ValueType: string; ValueData: "Doco.Document"; Flags: uninsdeletevalue
Root: HKCR; Subkey: ".pdf"; ValueType: string; ValueData: "Doco.Document"; Flags: uninsdeletevalue
Root: HKCR; Subkey: ".txt"; ValueType: string; ValueData: "Doco.Document"; Flags: uninsdeletevalue
Root: HKCR; Subkey: ".md"; ValueType: string; ValueData: "Doco.Document"; Flags: uninsdeletevalue

Root: HKCR; Subkey: "Doco.Document"; ValueType: string; ValueData: "Doco Document"; Flags: uninsdeletekey
Root: HKCR; Subkey: "Doco.Document\DefaultIcon"; ValueType: string; ValueData: "{app}\doco.exe,0"
Root: HKCR; Subkey: "Doco.Document\shell\open\command"; ValueType: string; ValueData: """{app}\doco.exe"" ""%1"""

Root: HKCR; Subkey: "*\shell\Open with Doco"; ValueType: string; ValueData: "Open with Doco"; Flags: uninsdeletekey
Root: HKCR; Subkey: "*\shell\Open with Doco\command"; ValueType: string; ValueData: """{app}\doco.exe"" ""%1"""
Root: HKLM; Subkey: "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"; \
  ValueType: expandsz; ValueName: "Path"; ValueData: "{olddata};{app}"; \
  Tasks: addtopath; Check: NeedsAddPath(ExpandConstant('{app}'))

[Code]
function NeedsAddPath(PathValue: string): Boolean;
var
  CurrentPath: string;
begin
  if not RegQueryStringValue(HKLM,
     'SYSTEM\CurrentControlSet\Control\Session Manager\Environment',
     'Path',
     CurrentPath) then
  begin
    Result := True;
    exit;
  end;

  Result := Pos(';' + UpperCase(PathValue) + ';', ';' + UpperCase(CurrentPath) + ';') = 0;
end;

