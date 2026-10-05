#define AppVersion GetEnv("WPS_RELEASE_VERSION")
#define BuildDir GetEnv("WPS_DIST_DIR")

[Setup]
AppId={{7D1B0F71-A6C4-4B61-9E0D-58F4E92A7C31}
AppName=WPS 表格校改
AppVersion={#AppVersion}
AppPublisher=WPS Proofreading Contributors
AppPublisherURL=https://github.com/phongapple-netizen/wps-spreadsheet-proofreading
AppSupportURL=https://github.com/phongapple-netizen/wps-spreadsheet-proofreading/issues
DefaultDirName={localappdata}\Programs\WPSSpreadsheetProofreading
DefaultGroupName=WPS 表格校改
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0
Compression=lzma2
SolidCompression=yes
CloseApplications=yes
RestartApplications=no
LicenseFile=..\LICENSE
OutputDir={#BuildDir}
OutputBaseFilename=WPS-Spreadsheet-Proofreading-{#AppVersion}-Windows-x64-Setup
UninstallDisplayIcon={app}\WPSSpreadsheetProofreadingServer.exe

[Files]
Source: "{#BuildDir}\WPSSpreadsheetProofreadingServer.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "安装说明.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\LICENSE"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\THIRD_PARTY_NOTICES.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\SOURCE_PROVENANCE.md"; DestDir: "{app}"; Flags: ignoreversion

[UninstallRun]
Filename: "{sys}\taskkill.exe"; Parameters: "/F /IM WPSSpreadsheetProofreadingServer.exe"; Flags: runhidden
Filename: "{app}\WPSSpreadsheetProofreadingServer.exe"; Parameters: "--uninstall"; Flags: runhidden skipifdoesntexist

[Code]
var
  InstallFailed: Boolean;

function GetCustomSetupExitCode: Integer;
begin
  Result := 0;
  if InstallFailed then
    Result := 1;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  ExitCode: Integer;
begin
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM WPSSpreadsheetProofreadingServer.exe', '', SW_HIDE, ewWaitUntilTerminated, ExitCode);
  Result := '';
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ExitCode: Integer;
  ServerPath: String;
begin
  if CurStep <> ssPostInstall then
    Exit;
  ServerPath := ExpandConstant('{app}\WPSSpreadsheetProofreadingServer.exe');
  if (not Exec(ServerPath, '--install', '', SW_HIDE, ewWaitUntilTerminated, ExitCode)) or (ExitCode <> 0) then
  begin
    { ssPostInstall exceptions display an error but do not set Setup's exit code. }
    InstallFailed := True;
    RaiseException('安装未完成：本机服务检查、启动或 WPS 注册失败。请查看 %LOCALAPPDATA%\WPSSpreadsheetProofreading\service-error.log。');
  end;
end;
