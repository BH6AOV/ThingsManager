#Requires -Version 5.1
























param(
    
    
    
    [ValidateSet('x64', 'arm64', 'x86')][string]$Arch = 'x64',
    [string]$Version = '',
    [string]$OutDir = '',
    [switch]$Portable,
    [switch]$SkipISCC,
    [switch]$ShowPlan
)
$ErrorActionPreference = 'Stop'
$OutputEncoding = [Console]::OutputEncoding = [Text.Encoding]::UTF8


$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path     
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



function Fail([string]$m) { Write-Host ('[X] ' + $m) -ForegroundColor Red; exit 1 }
function Info([string]$m) { Write-Host ('[.] ' + $m) -ForegroundColor Cyan }
function Ok([string]$m)   { Write-Host ('[OK] ' + $m) -ForegroundColor Green }


if (-not $Version) {
    $pkgFile = Join-Path $Root 'package.json'
    if (-not (Test-Path $pkgFile)) { Fail ('缺少 package.json：' + $pkgFile) }
    $Version = (Get-Content $pkgFile -Raw | ConvertFrom-Json).version
}
if (-not $Version) { Fail '无法确定版本号（package.json 里没有 version）' }
if ($Arch -eq 'x86') {
    Fail '不支持 32 位 x86：Node.js 官方自 v16 起不再提供 32 位 Windows 二进制，而本程序依赖 Node ≥ 22.13 的 node:sqlite。可选架构：x64 / arm64'
}

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
    if ($Portable) { Write-Host (' 5) 组装便携版并打包 ' + $OutBase + '_portable.zip') }
    Write-Host ' 6) 打印产物校验（大小 / SHA256）'
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


Info '同步 app 代码 → 载荷 ...'
if (-not (Test-Path $StageApp)) { New-Item -ItemType Directory -Path $StageApp -Force | Out-Null }
foreach ($f in @('server.js', 'supervisor.js', 'package.json', 'package-lock.json', 'start.bat', 'LICENSE', 'NOTICE', 'edition.default.json')) {
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

    
    
    
    
    $LangIsl = $null
    foreach ($c in @((Join-Path (Split-Path -Parent $iscc) 'Languages\ChineseSimplified.isl'),
                     (Join-Path $CacheDir 'ChineseSimplified.isl'))) {
        if ($c -and (Test-Path $c)) { $LangIsl = $c; break }
    }
    if (-not $LangIsl) {
        
        
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




$OutZip = Join-Path $OutDir ($OutBase + '_portable.zip')
if ($Portable) {
    Info '组装便携版并打包 zip ...'
    if (-not (Test-Path (Join-Path $Stage 'ThingsManager.exe'))) { Fail '便携版需要先编译控制器（请去掉 -SkipISCC）' }
    $PWork = Join-Path $OutDir ('.portable-' + $Arch)
    $PRoot = Join-Path $PWork 'ThingsManager'          
    if (Test-Path $PWork) { Remove-Item $PWork -Recurse -Force }
    New-Item -ItemType Directory -Path $PRoot -Force | Out-Null
    Copy-Item (Join-Path $Stage '*') $PRoot -Recurse -Force

    
    $flagText = "这个文件表示当前目录是 ThingsManager 便携版（免安装）。`r`n删除它之后，本程序会按「安装版」的方式运行（找系统服务）。`r`n"
    [IO.File]::WriteAllText((Join-Path $PRoot 'portable.flag'), $flagText, (New-Object Text.UTF8Encoding($true)))

    
    $cfg = @{ dataDir = '../data'; port = 3200; host = '0.0.0.0' } | ConvertTo-Json
    [IO.File]::WriteAllText((Join-Path $PRoot 'app\runtime.config.json'), $cfg, (New-Object Text.UTF8Encoding($false)))

    
    New-Item -ItemType Directory -Path (Join-Path $PRoot 'data') -Force | Out-Null

    $readme = @"
ThingsManager 便携版（免安装）V$Version
=========================================

1. 双击 ThingsManager.exe 启动，浏览器会自动打开管理面板（默认 http://127.0.0.1:3200）。
2. 关闭：右下角托盘图标（可能在“^”折叠区里）→「关闭程序」。
3. 数据存在本目录的 data 文件夹：备份或换电脑，拷走整个文件夹即可。
4. 请把整个文件夹解压到本地固定目录再运行（不要在压缩包里直接双击，也不要放在临时目录）。
5. 局域网访问：同网络设备打开 http://本机IP:3200（首次可能需在防火墙提示里允许）。
6. 端口被占用时：编辑 app\runtime.config.json 里的 port 换一个，重启程序生效。
7. 首次打开后建议：系统设置 → 账号与登录 里初始化管理员（默认开放模式，任何人都能改数据）。

LICENSE / NOTICE 见 app 文件夹；本项目代码采用 AGPL-3.0-or-later。
"@
    [IO.File]::WriteAllText((Join-Path $PRoot '使用说明.txt'), $readme, (New-Object Text.UTF8Encoding($true)))

    if (Test-Path $OutZip) { Remove-Item $OutZip -Force }
    
    & tar -a -c -f $OutZip -C $PWork 'ThingsManager'
    if ($LASTEXITCODE -ne 0) { Fail ('便携版 zip 打包失败（exit=' + $LASTEXITCODE + '）') }
    Remove-Item $PWork -Recurse -Force
    if (-not (Test-Path $OutZip)) { Fail ('便携版 zip 未生成：' + $OutZip) }
    Ok ('便携版就绪：{0} （{1:N2} MB）' -f (Split-Path -Leaf $OutZip), ((Get-Item $OutZip).Length / 1MB))
}


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
if ($Portable -and (Test-Path $OutZip)) {
    $hz = Get-FileHash $OutZip -Algorithm SHA256
    Write-Host ('  便携版  : ' + $OutZip)
    Write-Host ('  大小    : {0:N2} MB' -f ((Get-Item $OutZip).Length / 1MB))
    Write-Host ('  SHA256  : ' + $hz.Hash)
    Write-Host ('  用法    : 解压后双击 ThingsManager.exe（免安装，数据在本目录 data\）')
}
Write-Host '=============================================================='
