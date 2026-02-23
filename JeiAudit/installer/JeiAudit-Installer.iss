; JeiAudit Installer (Inno Setup)
; Compila con:
;   ISCC.exe JeiAudit-Installer.iss
; Opcional:
;   /DMyConfiguration=Release /DMyRevitYear=2024

#ifndef MyConfiguration
  #define MyConfiguration "Release"
#endif

#ifndef MyRevitYear
  #define MyRevitYear "2024"
#endif

#define MyAppName "JeiAudit"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Jason Rojas Estrada - Coordinador BIM"
#define MyAppURL "https://github.com/rjason381/JeiAudit"
#define MyAppExeName "JeiAudit.dll"
#define MyAddinDir "{userappdata}\Autodesk\Revit\Addins\" + MyRevitYear
#define MyPluginDir "{userappdata}\Autodesk\Revit\Addins\" + MyRevitYear + "\JeiAudit"
#define MyChecksetsDir "{userappdata}\JeiAudit\Checksets"

[Setup]
AppId={{9D5D0A6E-5FEA-4E44-B9FD-B3093708D7B5}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={localappdata}\JeiAudit
DefaultGroupName=JeiAudit
DisableDirPage=yes
DisableProgramGroupPage=yes
OutputDir=..\..\artifacts\installer
OutputBaseFilename=JeiAudit_Setup_R{#MyRevitYear}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\unins000.exe
CloseApplications=yes
ChangesAssociations=no
[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "copychecksets"; Description: "Copiar checksets base (XML) a AppData\Roaming\JeiAudit\Checksets"; Flags: unchecked

[InstallDelete]
Type: filesandordirs; Name: "{#MyPluginDir}"

[Dirs]
Name: "{#MyPluginDir}"
Name: "{#MyChecksetsDir}"; Tasks: copychecksets

[Files]
Source: "..\src\JeiAudit\bin\{#MyConfiguration}\*"; DestDir: "{#MyPluginDir}"; Excludes: "*.pdb"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\..\MCSettings_*_R2024.xml"; DestDir: "{#MyChecksetsDir}"; Flags: ignoreversion; Tasks: copychecksets

[Icons]
Name: "{group}\Desinstalar JeiAudit"; Filename: "{uninstallexe}"

[UninstallDelete]
Type: files; Name: "{#MyAddinDir}\JeiAudit.addin"
Type: filesandordirs; Name: "{#MyPluginDir}"

[Code]
function BuildManifest(): string;
var
  AssemblyPath: string;
begin
  AssemblyPath := ExpandConstant('{#MyPluginDir}\{#MyAppExeName}');
  Result :=
    '<?xml version="1.0" encoding="utf-8" standalone="no"?>' + #13#10 +
    '<RevitAddIns>' + #13#10 +
    '  <AddIn Type="Application">' + #13#10 +
    '    <Name>JeiAudit</Name>' + #13#10 +
    '    <Assembly>' + AssemblyPath + '</Assembly>' + #13#10 +
    '    <AddInId>9D5D0A6E-5FEA-4E44-B9FD-B3093708D7B5</AddInId>' + #13#10 +
    '    <FullClassName>JeiAudit.App</FullClassName>' + #13#10 +
    '    <VendorId>JEIA</VendorId>' + #13#10 +
    '    <VendorDescription>Herramienta de auditoria JeiAudit para Revit</VendorDescription>' + #13#10 +
    '  </AddIn>' + #13#10 +
    '</RevitAddIns>' + #13#10;
end;

procedure WriteAddinManifest();
var
  AddinFilePath: string;
  AddinDirPath: string;
begin
  AddinFilePath := ExpandConstant('{#MyAddinDir}\JeiAudit.addin');
  AddinDirPath := ExtractFileDir(AddinFilePath);
  ForceDirectories(AddinDirPath);
  SaveStringToFile(AddinFilePath, BuildManifest(), False);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    WriteAddinManifest();
  end;
end;
