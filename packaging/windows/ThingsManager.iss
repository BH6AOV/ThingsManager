; ============================================================
; ThingsManager · Windows 安装版（非 portable）— Inno Setup
; 作业根目录：dist/windows/（本脚本位于 dist/windows/build/，产物输出到 dist/windows/）
; 安装：注册为 Windows 服务(0.0.0.0:3200) + 托盘 + 开机自启(服务) + 数据在 %ProgramData%\ThingsManager
; 安装前：检测后台是否正在运行（服务 / 托盘）；在运行则弹窗提供「停止运行并继续安装」，停止后继续
; 安装后：[Run] 里的 --install-quiet 重新注册并启动服务、--tray 拉起托盘（即“安装完成后自动重启后端”）
; ============================================================
; 版本 / 架构 / 路径可由打包脚本或 CI 通过 ISCC 的 /D 注入（不传时用下方默认值，等于本目录直接编译）
#define MyAppName "ThingsManager"
#define MyPublisher "橙子木"
#ifndef MyAppVer
#define MyAppVer "0.10.9"
#endif
#ifndef VerInfoVersion
#define VerInfoVersion "0.10.9.0"
#endif
#ifndef MyArch
#define MyArch "x64"
#endif
#ifndef StageSource
; 相对本文件目录：dist/windows/build/stage/*
#define StageSource "stage\*"
#endif
#ifndef OutDir
; 相对本文件目录：dist/windows/
#define OutDir ".."
#endif
#ifndef OutBase
; 产物命名规范：
;   ThingsManager_<Version>_<system>_<cpuplatform>.<filetype>
;   例：ThingsManager_0.10.6_windows_AMD64.exe（cpuplatform: AMD64 / X86 / ARM64）
#if MyArch == "arm64"
#define OutBase "ThingsManager_" + MyAppVer + "_windows_ARM64"
#else
#define OutBase "ThingsManager_" + MyAppVer + "_windows_AMD64"
#endif
#endif
#ifndef LangFile
; 简体中文语言文件：默认用编译器自带的副本（本机安装 Inno Setup 通常自带中文）。
; CI 上 choco 安装的 Inno Setup 常缺中文，build.ps1 会下载一份到 .cache 并经 /DLangFile= 覆盖本值。
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
; arm64 安装包：只允许在 Windows on ARM 上安装（载荷内嵌 arm64 版 node.exe）
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
; 数据默认放 %ProgramData%\ThingsManager（不随卸载删除，保护数据）

[Languages]
Name: "chinesesimp"; MessagesFile: "{#LangFile}"

[Files]
Source: "{#StageSource}"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; 桌面/开始菜单图标双击 = 确保后台服务在运行并自动打开 http://<本机IP>:<端口>（服务未启动会先拉起，修复“关闭程序后再启动端口不起”）
[Icons]
Name: "{autoprograms}\ThingsManager · 仓库管理"; Filename: "{app}\ThingsManager.exe"; Parameters: "--open"; WorkingDir: "{app}"
Name: "{autodesktop}\ThingsManager · 仓库管理"; Filename: "{app}\ThingsManager.exe"; Parameters: "--open"; WorkingDir: "{app}"; Flags: createonlyiffileexists

[Registry]
; 托盘随“当前用户”登录自启（服务本身由 Windows 服务开机自启，二者配合实现常驻）
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "ThingsManagerTray"; ValueData: """{app}\ThingsManager.exe"" --tray"; Flags: uninsdeletevalue

[Run]
; 注册并启动服务（已存在则先停后重装，安装后自动重新启动后端）
Filename: "{app}\ThingsManager.exe"; Parameters: "--install-quiet"; Flags: runhidden waituntilterminated; StatusMsg: "正在注册并启动 ThingsManager 服务…"
Filename: "{sys}\netsh.exe"; Parameters: "advfirewall firewall delete rule name=ThingsManager3200"; Flags: runhidden; StatusMsg: "配置防火墙…"
Filename: "{sys}\netsh.exe"; Parameters: "advfirewall firewall add rule name=ThingsManager3200 dir=in action=allow protocol=TCP localport=3200 profile=private,domain,public"; Flags: runhidden; StatusMsg: "配置防火墙…"
; 立即拉起托盘守护（无需等下次登录；单实例，重复启动自动去重）
Filename: "{app}\ThingsManager.exe"; Parameters: "--tray"; Flags: nowait runascurrentuser; StatusMsg: "正在启动托盘守护…"
Filename: "http://127.0.0.1:3200/"; Flags: postinstall shellexec unchecked; Description: 打开管理面板（http://127.0.0.1:3200）

[UninstallRun]
; 停删后台服务（RunOnceId 保证每次卸载只执行一次，消除编译告警）
Filename: "{app}\ThingsManager.exe"; Parameters: "--uninstall-quiet"; Flags: runhidden waituntilterminated; RunOnceId: "StopService"
; 结束残留托盘进程（防止 exe 被占用导致删除失败）
Filename: "{sys}\taskkill.exe"; Parameters: "/IM ThingsManager.exe /F"; Flags: runhidden; RunOnceId: "KillTray"

; ============================================================
; 安装前检测：程序是否已在后台运行（服务 / 托盘）
;  在运行 → 弹出带「停止运行并继续安装」按钮的窗口；停止服务与托盘后继续安装
;  取消 → 中止安装并给出提示
;  安装完成后：由 [Run] 的 --install-quiet / --tray 重新启动后端与托盘
; 说明：仅使用 sc.exe / taskkill.exe（系统自带，不依赖已安装程序的新旧版本）
; ============================================================
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
