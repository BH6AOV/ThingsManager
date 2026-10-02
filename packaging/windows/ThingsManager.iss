; ============================================================







#define MyAppName "ThingsManager"
#define MyPublisher "橙子木"
#ifndef MyAppVer
#define MyAppVer "0.11.1"
#endif
#ifndef VerInfoVersion
#define VerInfoVersion "0.11.1.0"
#endif
#ifndef MyArch
#define MyArch "x64"
#endif
#ifndef StageSource

#define StageSource "stage\*"
#endif
#ifndef OutDir

#define OutDir ".."
#endif
#ifndef OutBase



#if MyArch == "arm64"
#define OutBase "ThingsManager_" + MyAppVer + "_windows_ARM64"
#else
#define OutBase "ThingsManager_" + MyAppVer + "_windows_AMD64"
#endif
#endif
#ifndef LangFile


#define LangFile "compiler:Languages\ChineseSimplified.isl"
#endif

[Setup]
AppId={{7E3B6C52-3F8A-4C11-9E7D-2B6A55C4E7F0}
AppName=ThingsManager · 轻量仓库管理
AppVersion={#MyAppVer}
AppVerName=ThingsManager {#MyAppVer}
AppPublisher={#MyPublisher}
AppCopyright=Copyright © 2026 {#MyPublisher}
VersionInfoVersion={#VerInfoVersion}
VersionInfoCompany={#MyPublisher}
VersionInfoDescription=ThingsManager · 轻量仓库管理
VersionInfoCopyright=Copyright © 2026 {#MyPublisher}
DefaultDirName={autopf}\ThingsManager
DefaultGroupName=ThingsManager
DisableProgramGroupPage=yes
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=commandline
#if MyArch == "arm64"

ArchitecturesAllowed=arm64
ArchitecturesInstallIn64BitMode=arm64
#else
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
#endif
OutputDir={#OutDir}
OutputBaseFilename={#OutBase}
SetupIconFile=
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName=ThingsManager {#MyAppVer}（卸载）
UninstallDisplayIcon={app}\ThingsManager.exe
CloseApplications=no


[Languages]
Name: "chinesesimp"; MessagesFile: "{#LangFile}"

[Files]
Source: "{#StageSource}"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs


[Icons]
Name: "{autoprograms}\ThingsManager · 仓库管理"; Filename: "{app}\ThingsManager.exe"; Parameters: "--open"; WorkingDir: "{app}"
Name: "{autodesktop}\ThingsManager · 仓库管理"; Filename: "{app}\ThingsManager.exe"; Parameters: "--open"; WorkingDir: "{app}"; Flags: createonlyiffileexists

[Registry]

Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "ThingsManagerTray"; ValueData: """{app}\ThingsManager.exe"" --tray"; Flags: uninsdeletevalue

[Run]

Filename: "{app}\ThingsManager.exe"; Parameters: "--install-quiet"; Flags: runhidden waituntilterminated; StatusMsg: "正在注册并启动 ThingsManager 服务…"
Filename: "{sys}\netsh.exe"; Parameters: "advfirewall firewall delete rule name=ThingsManager3200"; Flags: runhidden; StatusMsg: "配置防火墙…"
Filename: "{sys}\netsh.exe"; Parameters: "advfirewall firewall add rule name=ThingsManager3200 dir=in action=allow protocol=TCP localport=3200,3443 profile=private,domain,public"; Flags: runhidden; StatusMsg: "配置防火墙…"

Filename: "{app}\ThingsManager.exe"; Parameters: "--tray"; Flags: nowait runascurrentuser; StatusMsg: "正在启动托盘守护…"
Filename: "http://127.0.0.1:3200/"; Flags: postinstall shellexec unchecked; Description: 打开管理面板（http://127.0.0.1:3200）

[UninstallRun]

Filename: "{app}\ThingsManager.exe"; Parameters: "--uninstall-quiet"; Flags: runhidden waituntilterminated; RunOnceId: "StopService"

Filename: "{sys}\taskkill.exe"; Parameters: "/IM ThingsManager.exe /F"; Flags: runhidden; RunOnceId: "KillTray"








[Code]
const
  THMService = 'ThingsManager';
  THMTrayExe = 'ThingsManager.exe';

function THMRunCapture(const CmdLine: String; var OutText: AnsiString): Integer;
var
  TmpFile: String;
  Rc: Integer;
begin
  TmpFile := ExpandConstant('{tmp}\thm_probe.txt');
  DeleteFile(TmpFile);
  Exec(ExpandConstant('{sys}\cmd.exe'), '/C ' + CmdLine + ' > "' + TmpFile + '" 2>&1', '', SW_HIDE, ewWaitUntilTerminated, Rc);
  OutText := '';
  if FileExists(TmpFile) then
    LoadStringFromFile(TmpFile, OutText);
  Result := Rc;
end;

function THMServiceRunning: Boolean;
var
  OutText: AnsiString;
  S: String;
begin
  THMRunCapture('sc query ' + THMService, OutText);
  S := Uppercase(String(OutText));
  Result := (Pos('RUNNING', S) > 0) or (Pos('STOP_PENDING', S) > 0);
end;

function THMTrayRunning: Boolean;
var
  OutText: AnsiString;
  S: String;
begin
  THMRunCapture('tasklist /FI "IMAGENAME eq ' + THMTrayExe + '" /NH', OutText);
  S := Uppercase(String(OutText));
  Result := Pos(Uppercase(THMTrayExe), S) > 0;
end;

function THMAskStop: Boolean;
var
  Form: TSetupForm;
  Text1, Text2: TNewStaticText;
  BtnStop, BtnCancel: TNewButton;
begin
  // Inno 6.6+ 的 CreateCustomForm 需在构造时给定尺寸（之后 ClientWidth/ClientHeight 为只读）
  Form := CreateCustomForm(ScaleX(440), ScaleY(190), False, False);
  try
    Form.Caption := 'ThingsManager 正在后台运行';
    Form.Position := poScreenCenter;

    Text1 := TNewStaticText.Create(Form);
    Text1.Parent := Form;
    Text1.Left := ScaleX(16);
    Text1.Top := ScaleY(16);
    Text1.Width := ScaleX(408);
    Text1.WordWrap := True;
    Text1.Caption := '安装前需要先停止正在运行的 ThingsManager 后台服务与托盘程序。';

    Text2 := TNewStaticText.Create(Form);
    Text2.Parent := Form;
    Text2.Left := ScaleX(16);
    Text2.Top := ScaleY(54);
    Text2.Width := ScaleX(408);
    Text2.WordWrap := True;
    Text2.Caption := '点击「停止运行」后会自动停止后台服务（数据不会受影响），并继续安装；' +
      '安装完成后会自动重新启动后台服务与托盘。';

    BtnStop := TNewButton.Create(Form);
    BtnStop.Parent := Form;
    BtnStop.Caption := '停止运行并继续安装';
    BtnStop.Width := ScaleX(170);
    BtnStop.Height := ScaleY(30);
    BtnStop.Left := ScaleX(62);
    BtnStop.Top := ScaleY(134);
    BtnStop.Default := True;
    BtnStop.ModalResult := mrOk;

    BtnCancel := TNewButton.Create(Form);
    BtnCancel.Parent := Form;
    BtnCancel.Caption := '取消安装';
    BtnCancel.Width := ScaleX(110);
    BtnCancel.Height := ScaleY(30);
    BtnCancel.Left := ScaleX(248);
    BtnCancel.Top := ScaleY(134);
    BtnCancel.Cancel := True;
    BtnCancel.ModalResult := mrCancel;

    Result := (Form.ShowModal = mrOk);
  finally
    Form.Free;
  end;
end;

procedure THMStopServiceAndTray;
var
  Rc: Integer;
  I: Integer;
begin
  Exec(ExpandConstant('{sys}\sc.exe'), 'stop ' + THMService, '', SW_HIDE, ewWaitUntilTerminated, Rc);
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/IM ' + THMTrayExe + ' /F', '', SW_HIDE, ewWaitUntilTerminated, Rc);
  for I := 1 to 24 do
  begin
    if not THMServiceRunning then
      Break;
    Sleep(500);
  end;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := '';
  if not (THMServiceRunning or THMTrayRunning) then
    Exit;
  // 静默安装（/SILENT、/VERYSILENT）时直接停止，不弹窗，避免无人值守时卡住
  if WizardSilent then
  begin
    THMStopServiceAndTray;
    if THMServiceRunning then
      Result := '后台服务未能及时停止，安装已中止。';
    Exit;
  end;
  if not THMAskStop then
  begin
    Result := '安装已取消：ThingsManager 仍在后台运行。' + #13#10 +
      '请先停止后台运行后重新安装；也可在安装向导中点「停止运行」。';
    Exit;
  end;
  THMStopServiceAndTray;
  if THMServiceRunning then
    Result := '后台服务未能及时停止（可能仍在退出中）。' + #13#10 +
      '请在任务管理器 / 服务中确认 ThingsManager 已停止后再运行安装程序。';
end;
