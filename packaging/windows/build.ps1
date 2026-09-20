#Requires -Version 5.1
<#
  ThingsManager · Windows 安装包构建（CI 与本地通用）
  ===========================================================================
  用法（在仓库根执行）：
      powershell -NoProfile -ExecutionPolicy Bypass -File packaging\windows\build.ps1 -Arch x64
      powershell -NoProfile -ExecutionPolicy Bypass -File packaging\windows\build.ps1 -Arch arm64

  可选参数：
      -Version <x.y.z>   指定版本号（默认读 package.json）
      -OutDir  <目录>    产物输出目录（默认 <仓库根>\dist）
      -SkipISCC          只准备载荷，不编安装包（调试用）
      -ShowPlan          只打印将执行的动作并退出

  产物：<OutDir>\ThingsManager_<版本>_windows_<CPU平台>.exe（如 ThingsManager_0.10.6_windows_AMD64.exe；
        命名规范见 packaging/README.md「支持矩阵」）

  步骤：
      1) 内嵌 Node 运行时：下载官方单文件 node.exe（带缓存 .cache\），放进 stage-<arch>\runtime
      2) 组装载荷 stage-<arch>：app 代码 + node_modules（剔除 @napi-rs）+ logo.ico + 编译控制器
      3) Inno Setup 编译安装包（通过 /D 注入版本与架构）

  前置：仓库根已执行 npm ci --omit=dev；Inno Setup 6（CI 里 choco install innosetup -y）。
#>
param(
    # 支持的架构：x64（→ AMD64）/ arm64（→ ARM64）。
    # x86（32 位）不可用：Node.js 官方自 v16 起不再提供 32 位 Windows 二进制，
    # 而本程序依赖 Node ≥ 22.13 的 node:sqlite（下面会在校验里明确报错）。
    [ValidateSet('x64', 'arm64', 'x86')][string]$Arch = 'x64',
    [string]$Version = '',
    [string]$OutDir = '',
    [switch]$SkipISCC,
    [switch]$ShowPlan
)
$ErrorActionPreference = 'Stop'
$OutputEncoding = [Console]::OutputEncoding = [Text.Encoding]::UTF8

# ---------- 路径定位 ----------
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path     # ...\packaging\windows
$Root = $ScriptDir
for ($i = 0; $i -lt 6; $i++) {
    if (Test-Path (Join-Path $Root 'package.json')) { break }
    $Root = Split-Path -Parent $Root
}
if (-not (Test-Path (Join-Path $Root 'server.js'))) { Write-Host ('[X] 未找到仓库根（缺 server.js）：' + $Root) -ForegroundColor Red; exit 1 }

$Stage        = Join-Path $ScriptDir ('stage-' + $Arch)
$StageApp     = Join-Path $Stage 'app'
$StageRuntime = Join-Path $Stage 'runtime'
$IssFile      = Join-Path $ScriptDir 'ThingsManager.iss'
$Controller   = Join-Path $ScriptDir 'controller\ThingsManager.cs'
$Logo         = Join-Path $ScriptDir 'assets\logo.ico'
$CacheDir     = Join-Path $Root '.cache'
if (-not $OutDir) { $OutDir = Join-Path $Root 'dist' }
$OutDir = [IO.Path]::GetFullPath($OutDir)
# 产物命名规范（见 packaging/README.md「支持矩阵」）：ThingsManager_<Version>_<system>_<cpuplatform>.<filetype>
#   $OutBase / $OutSetup 在下面拿到版本号后再计算

function Fail([string]$m) { Write-Host ('[X] ' + $m) -ForegroundColor Red; exit 1 }
function Info([string]$m) { Write-Host ('[.] ' + $m) -ForegroundColor Cyan }
function Ok([string]$m)   { Write-Host ('[OK] ' + $m) -ForegroundColor Green }

# ---------- 版本与 Node 版本（单一来源：package.json / .node-version） ----------
if (-not $Version) {
    $pkgFile = Join-Path $Root 'package.json'
    if (-not (Test-Path $pkgFile)) { Fail ('缺少 package.json：' + $pkgFile) }
    $Version = (Get-Content $pkgFile -Raw | ConvertFrom-Json).version
}
if (-not $Version) { Fail '无法确定版本号（package.json 里没有 version）' }
if ($Arch -eq 'x86') {
    Fail '不支持 32 位 x86：Node.js 官方自 v16 起不再提供 32 位 Windows 二进制，而本程序依赖 Node ≥ 22.13 的 node:sqlite。可选架构：x64 / arm64'
}
# 产物命名：x64 → AMD64，arm64 → ARM64
$CpuPlatform = if ($Arch -eq 'arm64') { 'ARM64' } else { 'AMD64' }
$OutBase = 'ThingsManager_' + $Version + '_windows_' + $CpuPlatform
$OutSetup = Join-Path $OutDir ($OutBase + '.exe')
$NodeVer = '24.20.0'
$verFile = Join-Path $Root '.node-version'
if (Test-Path $verFile) { $v = (Get-Content $verFile -Raw).Trim(); if ($v) { $NodeVer = $v } }

if ($ShowPlan) {
    Write-Host ''
    Write-Host ('仓库根  : ' + $Root)
    Write-Host ('载荷目录: ' + $Stage)
    Write-Host ('产物    : ' + $OutSetup)
    Write-Host ('版本    : V' + $Version + '  内嵌 Node v' + $NodeVer + '  架构: ' + $Arch)
    Write-Host ''
    Write-Host '将要执行：'
    Write-Host ' 1) 准备内嵌 Node 运行时（node.exe）'
    Write-Host ' 2) 同步 app 代码与 node_modules 到载荷（剔除 @napi-rs）'
    Write-Host ' 3) 用系统自带 csc 编译托盘/服务控制器 ThingsManager.exe'
    if (-not $SkipISCC) { Write-Host (' 4) ISCC 编译 ' + $OutBase + '.exe') }
    Write-Host ' 5) 打印产物校验（大小 / SHA256）'
    exit 0
}

Write-Host '=============================================================='
Info ('Windows 安装包构建：V' + $Version + '  架构=' + $Arch + '  内嵌 Node=v' + $NodeVer)
Write-Host '=============================================================='

function Mirror([string]$src, [string]$dst, [string[]]$excludeDirs) {
    if (-not (Test-Path $src)) { Fail ('缺少目录：' + $src) }
    $robocopy = Get-Command robocopy.exe -ErrorAction SilentlyContinue
    if ($robocopy) {
        $rcArgs = @($src, $dst, '/MIR', '/NFL', '/NDL', '/NJH', '/NJS', '/NP', '/R:2', '/W:2')
        if ($excludeDirs) { $rcArgs += '/XD'; $rcArgs += $excludeDirs }
        & robocopy @rcArgs | Out-Null
        if ($LASTEXITCODE -ge 8) { Fail ('robocopy 同步失败（exit=' + $LASTEXITCODE + '）：' + $src) }
    } else {
        if (Test-Path $dst) { Remove-Item $dst -Recurse -Force }
        Copy-Item $src $dst -Recurse -Force
        if ($excludeDirs) { foreach ($e in $excludeDirs) { Remove-Item (Join-Path $dst $e) -Recurse -Force -ErrorAction SilentlyContinue } }
    }
}

# ---------- 1) 内嵌 Node 运行时 ----------
Info '准备内嵌 Node 运行时 ...'
$nodeExe = Join-Path $StageRuntime 'node.exe'
if (-not (Test-Path $nodeExe)) {
    if (-not (Test-Path $StageRuntime)) { New-Item -ItemType Directory -Path $StageRuntime -Force | Out-Null }
    $cached = Join-Path $CacheDir ('node-v' + $NodeVer + '-win-' + $Arch + '.exe')
    if (-not (Test-Path $cached)) {
        $url = 'https://nodejs.org/dist/v' + $NodeVer + '/win-' + $Arch + '/node.exe'
        Info ('下载：' + $url)
        New-Item -ItemType Directory -Path $CacheDir -Force | Out-Null
        $oldProgress = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        try {
            Invoke-WebRequest -Uri $url -OutFile ($cached + '.tmp') -UseBasicParsing
            Move-Item ($cached + '.tmp') $cached -Force
        } catch {
            Fail ('Node 运行时下载失败：' + $_.Exception.Message)
        } finally { $ProgressPreference = $oldProgress }
    }
    Copy-Item $cached $nodeExe -Force
}
Ok ('内嵌 Node 就绪：{0:N1} MB' -f ((Get-Item $nodeExe).Length / 1MB))

# ---------- 2) 组装载荷 ----------
Info '同步 app 代码 → 载荷 ...'
if (-not (Test-Path $StageApp)) { New-Item -ItemType Directory -Path $StageApp -Force | Out-Null }
foreach ($f in @('server.js', 'supervisor.js', 'package.json', 'package-lock.json', 'start.bat', 'LICENSE', 'NOTICE')) {
    $src = Join-Path $Root $f
    if (Test-Path $src) { Copy-Item $src (Join-Path $StageApp $f) -Force }
}
Mirror (Join-Path $Root 'static')   (Join-Path $StageApp 'static')
Mirror (Join-Path $Root 'template') (Join-Path $StageApp 'template')
Info '同步运行依赖 node_modules ...'
Mirror (Join-Path $Root 'node_modules') (Join-Path $StageApp 'node_modules') @('@napi-rs')
Copy-Item $Logo (Join-Path $Stage 'logo.ico') -Force
Ok ('载荷就绪：{0:N1} MB，文件数 ' -f ((Get-ChildItem $Stage -Recurse -File | Measure-Object -Property Length -Sum).Sum / 1MB) +
    (Get-ChildItem $Stage -Recurse -File | Measure-Object).Count)

# ---------- 3) 编译托盘/服务控制器（系统自带 .NET Framework 编译器） ----------
if (-not $SkipISCC) {
    Info '编译控制器 ThingsManager.exe ...'
    if (-not (Test-Path $Controller)) { Fail ('缺少控制器源码：' + $Controller) }
    $fw = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319'
    $csc = Join-Path $fw 'csc.exe'
    if (-not (Test-Path $csc)) { Fail ('未找到 csc.exe：' + $csc) }
    $outExe = Join-Path $Stage 'ThingsManager.exe'
    & $csc /nologo /target:winexe /platform:anycpu /optimize+ /codepage:65001 `
        "/win32icon:$Logo" "/out:$outExe" `
        "/r:$fw\System.dll" /r:System.Core.dll /r:System.Windows.Forms.dll `
        /r:System.Drawing.dll /r:System.ServiceProcess.dll /r:System.Configuration.Install.dll `
        $Controller
    if ($LASTEXITCODE -ne 0) { Fail ('控制器编译失败（exit=' + $LASTEXITCODE + '）') }
    if (-not (Test-Path $outExe)) { Fail '控制器编译后未生成 ThingsManager.exe' }
    Ok ('控制器就绪：{0:N0} KB' -f ((Get-Item $outExe).Length / 1KB))
}

# ---------- 4) Inno Setup 编译安装包 ----------
if (-not $SkipISCC) {
    Info 'ISCC 编译安装包 ...'
    if (-not (Test-Path $IssFile)) { Fail ('缺少 Inno 脚本：' + $IssFile) }
    $iscc = $null
    foreach ($c in @(
            (Join-Path $env:LOCALAPPDATA 'Programs\Inno Setup 6\ISCC.exe'),
            (Join-Path ${env:ProgramFiles(x86)} 'Inno Setup 6\ISCC.exe'),
            (Join-Path $env:ProgramFiles 'Inno Setup 6\ISCC.exe'))) {
        if ($c -and (Test-Path $c)) { $iscc = $c; break }
    }
    if (-not $iscc) {
        $g = Get-Command ISCC.exe -ErrorAction SilentlyContinue
        if ($g) { $iscc = $g.Source }
    }
    if (-not $iscc) { Fail '未找到 Inno Setup 6 的 ISCC.exe（CI：choco install innosetup -y）' }
    if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }

    # 简体中文语言文件：**不是每个 Inno Setup 安装包都带**（CI 里 choco 装的 6.7.1 就没有中文，
    # 装成英文界面会直接编译失败：Couldn't open include file "...\Languages\ChineseSimplified.isl"）。
    # 规则：① 编译器目录自带的（译者在官方翻译页发布、随新版安装包分发）优先用；
    #       ② 没有就自动下载一份到 .cache\（不入版本库，与本机装的一致）。
    $LangIsl = $null
    foreach ($c in @((Join-Path (Split-Path -Parent $iscc) 'Languages\ChineseSimplified.isl'),
                     (Join-Path $CacheDir 'ChineseSimplified.isl'))) {
        if ($c -and (Test-Path $c)) { $LangIsl = $c; break }
    }
    if (-not $LangIsl) {
        # 第 1 条 = 官方翻译页（https://jrsoftware.org/files/istrans/）给出的直链，始终是最新官方版本；
        # 第 2 条 = 随 6.7.3 源码树发布的副本，作为兜底（内容略旧，缺部分卸载提示文案）。
        $islUrls = @(
            'https://raw.githubusercontent.com/jrsoftware/issrc/refs/heads/main/Files/Languages/ChineseSimplified.isl',
            'https://raw.githubusercontent.com/jrsoftware/issrc/is-6_7_3/Files/Languages/Unofficial/ChineseSimplified.isl'
        )
        New-Item -ItemType Directory -Path $CacheDir -Force | Out-Null
        $LangIsl = Join-Path $CacheDir 'ChineseSimplified.isl'
        $oldProgress = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        $islErr = ''
        foreach ($u in $islUrls) {
            try {
                Info ('下载简体中文语言文件：' + $u)
                Invoke-WebRequest -Uri $u -OutFile ($LangIsl + '.tmp') -UseBasicParsing
                Move-Item ($LangIsl + '.tmp') $LangIsl -Force
                $islErr = ''
                break
            } catch {
                $islErr = $_.Exception.Message
                Remove-Item ($LangIsl + '.tmp') -Force -ErrorAction SilentlyContinue
            }
        }
        $ProgressPreference = $oldProgress
        if ($islErr) { Fail ('简体中文语言文件下载失败：' + $islErr) }
    }
    Ok ('简体中文语言文件：' + $LangIsl + '（' + [math]::Round((Get-Item $LangIsl).Length / 1KB, 1) + ' KB）')

    & $iscc "/DMyAppVer=$Version" "/DVerInfoVersion=$Version.0" "/DMyArch=$Arch" `
        "/DOutBase=$OutBase" "/DStageSource=$Stage\*" "/DOutDir=$OutDir" "/DLangFile=$LangIsl" $IssFile
    if ($LASTEXITCODE -ne 0) { Fail ('ISCC 编译失败（exit=' + $LASTEXITCODE + '）') }
    if (-not (Test-Path $OutSetup)) { Fail ('编译后未找到：' + $OutSetup) }
}

# ---------- 5) 产物校验 ----------
Write-Host ''
Write-Host '=============================================================='
Ok '构建完成'
if (Test-Path $OutSetup) {
    $h = Get-FileHash $OutSetup -Algorithm SHA256
    $vi = (Get-Item $OutSetup).VersionInfo
    Write-Host ('  产物    : ' + $OutSetup)
    Write-Host ('  大小    : {0:N2} MB' -f ((Get-Item $OutSetup).Length / 1MB))
    Write-Host ('  产品版本: ' + $vi.ProductVersion + '   文件版本: ' + $vi.FileVersion)
    Write-Host ('  SHA256  : ' + $h.Hash)
    Write-Host ('  适用    : ' + $Arch + '（Windows 10/11）')
}
Write-Host '=============================================================='
