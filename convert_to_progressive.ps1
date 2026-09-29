<#
.SYNOPSIS
    将 videos 目录下的视频批量转换为「边下边播(progressive download)」格式，
    尽量保持原容器 + 原编码（只重封装），转换尽可能快，并记录每个视频耗时，生成统计表格。

.DESCRIPTION
    边下边播的关键：
      - MP4/MOV : -movflags +faststart，把 moov 挪到文件头部，边下边播（浏览器/播放器原生支持）
      - MKV/WebM: Matroska 本身就是流式容器，可保持原格式原编码直接复制

    「尽可能不要转换格式」的实现（策略由快到慢自动选择，失败自动降级重试）：
      1) remux         原容器可边下边播 + 编码通用 -> -c copy 原样重封装（毫秒级，无损、不改格式）
      2) recopy        原容器不适合流式(avi/wmv/mpg/ts/flv/3gp/ogv) 但编码是 H.264/通用音频
                       -> 只换容器到 MP4，-c copy + faststart（不重编码）
      3) copyvideo     视频编码通用、音频不通用 -> -c:v copy 只编码音轨
      4) transcode     其余编码(MPEG4/ProRes/WMV/MPEG2/Theora...) -> 重编码 H.264 + AAC 输出 MP4
                       优先硬件编码(NVENC/QSV/AMF，自动探测)，失败自动回退 libx264 -preset ultrafast

    产物（默认输出到 output 目录）：
      output\*.*                       转换后的视频（保持原扩展名，除非必须转 MP4）
      output\conversion_report.csv     表格(CSV，Excel 可直接打开)
      output\conversion_report.md      表格(Markdown)
      output\conversion_report.html    表格(HTML，浏览器直接打开)

.EXAMPLE
    powershell -ExecutionPolicy Bypass -NoProfile -File .\convert_to_progressive.ps1
    powershell -ExecutionPolicy Bypass -NoProfile -File .\convert_to_progressive.ps1 -Limit 2        # 先跑 2 个试速度
    powershell -ExecutionPolicy Bypass -NoProfile -File .\convert_to_progressive.ps1 -ForceMp4       # 全部统一输出 MP4
    powershell -ExecutionPolicy Bypass -NoProfile -File .\convert_to_progressive.ps1 -Encoder cpu    # 只用 CPU 软件编码
    powershell -ExecutionPolicy Bypass -NoProfile -File .\convert_to_progressive.ps1 -Parallel 8     # 提高并发
    powershell -ExecutionPolicy Bypass -NoProfile -File .\convert_to_progressive.ps1 -AllowHevc      # HEVC 也直接复制不重编码
#>

[CmdletBinding()]
param(
    [string] $InputDir  = '',
    [string] $OutputDir = '',
    [string] $Ffmpeg    = '',

    # 重编码时使用的编码器: auto=自动探测硬件加速 / cpu=libx264 / nvenc / qsv / amf
    [ValidateSet('auto', 'cpu', 'nvenc', 'qsv', 'amf')]
    [string] $Encoder = 'auto',

    # libx264 速度档（越靠前越快）: ultrafast < superfast < veryfast < faster < fast < medium
    [ValidateSet('ultrafast', 'superfast', 'veryfast', 'faster', 'fast', 'medium')]
    [string] $Preset = 'ultrafast',

    [int]    $Crf          = 26,     # 重编码质量（越小画质越好、体积越大）
    [string] $AudioBitrate = '128k',

    [switch] $ForceMp4,   # 全部统一输出 MP4（不保留 MKV/WebM/MOV 原容器）
    [switch] $AllowHevc,  # 允许 HEVC 直接复制（部分浏览器/设备不支持 HEVC）

    [string[]] $Extensions = @('.mp4', '.mkv', '.mov', '.avi', '.webm', '.flv',
                               '.wmv', '.mpg', '.mpeg', '.ts', '.3gp', '.ogv', '.m4v'),

    [int] $Parallel   = 0,     # 并发数，0 = 自动(min(4, CPU 核数))
    [int] $Limit      = 0,     # >0 时只转换前 N 个（用于试跑）
    [int] $TimeoutSec = 900,   # 单个视频超时秒数
    [switch] $Force            # 已存在输出时也重新转换
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------- 默认路径
if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    $ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
} else {
    $ScriptRoot = $PSScriptRoot
}
if ([string]::IsNullOrWhiteSpace($ScriptRoot)) { $ScriptRoot = (Get-Location).Path }
if ([string]::IsNullOrWhiteSpace($InputDir))  { $InputDir  = Join-Path $ScriptRoot 'videos' }
if ([string]::IsNullOrWhiteSpace($OutputDir)) { $OutputDir = Join-Path $ScriptRoot 'output' }
if ([string]::IsNullOrWhiteSpace($Ffmpeg))    { $Ffmpeg    = Join-Path $ScriptRoot 'tools\ffmpeg.exe' }

if ($Parallel -le 0) { $Parallel = [Math]::Min(4, [Environment]::ProcessorCount) }

# 可原生边下边播的容器（保持原格式时使用）
$KeepableContainers = @('.mp4', '.m4v', '.mov', '.mkv', '.webm')

# ---------------------------------------------------------------- 工具函数
function Write-Utf8File {
    param([string] $Path, [string] $Content)
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $enc = New-Object System.Text.UTF8Encoding($true)   # 带 BOM，Excel/记事本打开不乱码
    [System.IO.File]::WriteAllText($Path, $Content, $enc)
}

# 解析 ffmpeg -i 输出，拿到时长/分辨率/视频编码/音频编码
function Get-MediaInfo {
    param([string] $Ffmpeg, [string] $File)
    $info = [pscustomobject]@{ Duration = ''; Resolution = ''; VCodec = ''; ACodec = '' }
    try {
        $oldEap = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        $raw = (& $Ffmpeg -hide_banner -i $File 2>&1 | Out-String)
        $ErrorActionPreference = $oldEap
        $m = [regex]::Match($raw, 'Duration:\s*([0-9:.]+)')
        if ($m.Success) { $info.Duration = $m.Groups[1].Value }
        foreach ($line in ($raw -split "`r?`n")) {
            if ($line -match 'Stream\s+#\d+:\d+.*\bVideo:\s*([A-Za-z0-9_]+)') {
                $info.VCodec = $Matches[1].ToLower()
                $r = [regex]::Match($line, '(\d{2,5})x(\d{2,5})')
                if ($r.Success) { $info.Resolution = "$($r.Groups[1].Value)x$($r.Groups[2].Value)" }
            }
            elseif ($line -match 'Stream\s+#\d+:\d+.*\bAudio:\s*([A-Za-z0-9_]+)') {
                if (-not $info.ACodec) { $info.ACodec = $Matches[1].ToLower() }
            }
        }
    } catch { }
    return $info
}

# 探测可用的 H.264 硬件编码器，返回 '' / h264_nvenc / h264_qsv / h264_amf
function Resolve-HwEncoder {
    param([string] $Ffmpeg, [string] $Preference)
    if ($Preference -eq 'cpu') { return '' }
    $oldEap = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $list = (& $Ffmpeg -hide_banner -encoders 2>&1 | Out-String)
    $ErrorActionPreference = $oldEap
    $candidates = if ($Preference -eq 'auto') { @('h264_nvenc', 'h264_qsv', 'h264_amf') } else { @("h264_$Preference") }
    foreach ($name in $candidates) {
        if ([regex]::IsMatch($list, "\s$([regex]::Escape($name))\s")) { return $name }
    }
    if ($Preference -ne 'auto') { Write-Host "  [警告] 指定的编码器 $Preference 不可用，回退 CPU 软件编码" -ForegroundColor Yellow }
    return ''
}

# 视频编码在该容器里是否「可直接复制且能边下边播」
function Test-VideoOk {
    param([string] $Codec, [string] $Container, [bool] $Hevc)
    if ([string]::IsNullOrWhiteSpace($Codec)) { return $false }
    switch ($Container) {
        '.webm' { return ($Codec -in @('vp8', 'vp9', 'av1')) }
        '.mkv'  { return (($Codec -in @('h264', 'vp8', 'vp9', 'av1')) -or ($Hevc -and $Codec -in @('hevc', 'h265'))) }
        default { return (($Codec -in @('h264', 'av1')) -or ($Hevc -and $Codec -in @('hevc', 'h265'))) }
    }
}

# 音频编码在该容器里是否可直接复制
function Test-AudioOk {
    param([string] $Codec, [string] $Container)
    if ([string]::IsNullOrWhiteSpace($Codec)) { return $true }   # 无音轨
    switch ($Container) {
        '.webm' { return ($Codec -in @('opus', 'vorbis')) }
        '.mkv'  { return ($Codec -in @('aac', 'mp3', 'opus', 'vorbis', 'flac')) }
        default { return ($Codec -in @('aac', 'mp3')) }
    }
}

# 根据容器 + 编码决定「最快且尽量不改格式」的策略
function Get-Target {
    param([string] $SrcExt, [string] $VCodec, [string] $ACodec, [bool] $Hevc, [bool] $ForceMp4)
    $ext = $SrcExt.ToLower()
    $keep = (-not $ForceMp4) -and ($KeepableContainers -contains $ext)
    $target = if ($keep) { $ext } else { '.mp4' }

    $vOk = Test-VideoOk -Codec $VCodec -Container $target -Hevc $Hevc
    $aOk = Test-AudioOk -Codec $ACodec -Container $target

    if ($vOk -and $aOk) {
        if ($keep) { return [pscustomobject]@{ Plan = 'remux';   Ext = $target } }
        return            [pscustomobject]@{ Plan = 'recopy';  Ext = '.mp4' }
    }
    if ($vOk) {
        if ($keep) { return [pscustomobject]@{ Plan = 'copyvideo'; Ext = $target } }
        return            [pscustomobject]@{ Plan = 'copyvideo-mp4'; Ext = '.mp4' }
    }
    return [pscustomobject]@{ Plan = 'transcode'; Ext = '.mp4' }   # 编码不通用，只能重编码
}

# 容器 -> ffmpeg -f 封装格式
function Get-ContainerFormat {
    param([string] $Ext)
    switch ($Ext.ToLower()) {
        '.mkv'  { return 'matroska' }
        '.webm' { return 'webm' }
        '.mov'  { return 'mov' }
        default { return 'mp4' }
    }
}

# 容器 -> 音频编码参数（需要编码音轨时使用）
function Get-AudioArgs {
    param([string] $Ext, [string] $AudioBitrate)
    if ($Ext.ToLower() -eq '.webm') { return @('-c:a', 'libopus', '-b:a', $AudioBitrate) }
    return @('-c:a', 'aac', '-b:a', $AudioBitrate)
}

# 输出路径：保持原格式的策略用原扩展名，重编码/换容器统一 .mp4
function Get-OutPath {
    param($Item, [string] $Mode, [string] $OutputDir)
    $ext = if ($Mode -eq 'remux' -or $Mode -eq 'copyvideo') { $Item.KeepExt } else { '.mp4' }
    Join-Path $OutputDir "$($Item.File.BaseName)$ext"
}

# 判断输出文件是否为 moov 前置（可在下载过程中就开始播放）
function Test-FastStart {
    param([string] $Path)
    try {
        $fs = [System.IO.File]::OpenRead($Path)
        $len = [int][Math]::Min([long]1MB, [long]$fs.Length)
        $buf = New-Object byte[] $len
        [void]$fs.Read($buf, 0, $len)
        $fs.Dispose()
        return ([System.Text.Encoding]::ASCII.GetString($buf).IndexOf('moov') -ge 0)
    } catch { return $false }
}

# ---------------------------------------------------------------- ffmpeg 参数构建
function Build-FfmpegArgs {
    param([string] $Src, [string] $Out, [string] $Mode, [string] $Ext,
          [string] $Preset, [int] $Crf, [string] $AudioBitrate, [string] $HwEncoder)
    $a = New-Object System.Collections.Generic.List[string]
    foreach ($x in @('-y', '-hide_banner', '-nostdin', '-nostats', '-loglevel', 'error', '-i', $Src,
                     '-map', '0:v:0', '-map', '0:a:0?', '-sn', '-dn')) { $a.Add($x) }
    switch -Wildcard ($Mode) {
        'remux'  { $a.Add('-c'); $a.Add('copy') }
        'recopy' { $a.Add('-c'); $a.Add('copy') }
        'copyvideo*' {
            $a.Add('-c:v'); $a.Add('copy')
            foreach ($x in (Get-AudioArgs -Ext $Ext -AudioBitrate $AudioBitrate)) { $a.Add($x) }
        }
        default {
            $hw = if ($Mode -eq 'transcode:cpu') { '' } else { $HwEncoder }
            $v = switch ($hw) {
                'h264_nvenc' { @('-c:v', 'h264_nvenc', '-preset', 'p1', '-cq', "$Crf", '-b:v', '0', '-pix_fmt', 'yuv420p') }
                'h264_qsv'   { @('-c:v', 'h264_qsv', '-preset', 'veryfast', '-global_quality', "$Crf", '-look_ahead', '0', '-pix_fmt', 'nv12') }
                'h264_amf'   { @('-c:v', 'h264_amf', '-quality', 'speed', '-rc', 'cqp', '-qp', "$Crf", '-pix_fmt', 'yuv420p') }
                default      { @('-c:v', 'libx264', '-preset', $Preset, '-crf', "$Crf", '-pix_fmt', 'yuv420p', '-threads', '0') }
            }
            foreach ($x in $v) { $a.Add($x) }
            foreach ($x in (Get-AudioArgs -Ext $Ext -AudioBitrate $AudioBitrate)) { $a.Add($x) }
        }
    }
    $fmt = Get-ContainerFormat -Ext $Ext
    if ($fmt -eq 'mp4' -or $fmt -eq 'mov') { $a.Add('-movflags'); $a.Add('+faststart') }
    foreach ($x in @('-map_metadata', '0', '-f', $fmt, $Out)) { $a.Add($x) }
    return , $a.ToArray()
}

# 尝试序列: 最快策略 -> 全量重编码(硬件) -> 全量重编码(CPU)，任一成功即停
function Get-ModeSequence {
    param([string] $Plan, [string] $HwEncoder)
    $modes = New-Object System.Collections.Generic.List[string]
    $modes.Add($Plan)
    if ($Plan -ne 'transcode') { $modes.Add('transcode') }
    if ($HwEncoder -ne '')     { $modes.Add('transcode:cpu') }
    return , $modes.ToArray()
}

# 启动一个 ffmpeg 进程（异步，返回句柄）
# 注意: 不用 Start-Process，因为其返回的 Process 在退出后 ExitCode 不可用；这里直接用 .NET Process
function Start-FfmpegTask {
    param([string] $Ffmpeg, [string[]] $ArgList, $Item, [int] $Index, [int] $ModeIndex, [double] $Carry, [string] $Out)
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $Ffmpeg
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardError = $true
    $psi.RedirectStandardOutput = $true
    $sb = New-Object System.Text.StringBuilder
    foreach ($x in $ArgList) {
        if ($sb.Length -gt 0) { [void]$sb.Append(' ') }
        if ($x -match '[\s"]') { [void]$sb.Append('"'); [void]$sb.Append($x); [void]$sb.Append('"') }
        else { [void]$sb.Append($x) }
    }
    $psi.Arguments = $sb.ToString()
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $p = [System.Diagnostics.Process]::Start($psi)
    $errTask = $p.StandardError.ReadToEndAsync()
    $outTask = $p.StandardOutput.ReadToEndAsync()
    return [pscustomobject]@{
        Proc = $p; Sw = $sw; ErrTask = $errTask; OutTask = $outTask
        Item = $Item; Index = $Index; ModeIndex = $ModeIndex; Carry = $Carry
        Start = [DateTime]::Now; Out = $Out
    }
}

# 读取任务日志并释放进程句柄
function Read-TaskLog {
    param($Task)
    $log = ''
    try { $log = ([string]$Task.ErrTask.Result) + "`n" + ([string]$Task.OutTask.Result) } catch { }
    try { $Task.Proc.Close() } catch { }
    return [string]$log
}

# ---------------------------------------------------------------- 前置检查
if (-not (Test-Path -LiteralPath $Ffmpeg)) { throw "未找到 ffmpeg: $Ffmpeg" }
if (-not (Test-Path -LiteralPath $InputDir)) { throw "未找到输入目录: $InputDir" }
if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

$hw = Resolve-HwEncoder -Ffmpeg $Ffmpeg -Preference $Encoder

$files = @(Get-ChildItem -LiteralPath $InputDir -File |
           Where-Object { $Extensions -contains $_.Extension.ToLower() } |
           Sort-Object Name)
if ($Limit -gt 0) { $files = @($files | Select-Object -First $Limit) }
if ($files.Count -eq 0) { throw "输入目录中没有匹配的视频文件: $InputDir" }

Write-Host ''
Write-Host '=========== 边下边播转换 (尽量保持原格式) ===========' -ForegroundColor Cyan
Write-Host "输入目录    : $InputDir"
Write-Host "输出目录    : $OutputDir"
Write-Host "保持原容器  : $(if ($ForceMp4) { '否(-ForceMp4)' } else { '是 (mp4/m4v/mov/mkv/webm)' })"
Write-Host "重编码编码器: $(if ($hw) { $hw } else { 'libx264 (CPU)' })  (preset=$Preset, crf=$Crf)"
Write-Host "并发数      : $Parallel"
Write-Host "待转换文件  : $($files.Count) 个"
Write-Host '======================================================' -ForegroundColor Cyan
Write-Host ''

# ---------------------------------------------------------------- 探测源文件 & 制定策略
$items = New-Object System.Collections.Generic.List[object]
Write-Host '正在探测源文件编码...' -ForegroundColor DarkGray
foreach ($f in $files) {
    $info = Get-MediaInfo -Ffmpeg $Ffmpeg -File $f.FullName
    $t = Get-Target -SrcExt $f.Extension -VCodec $info.VCodec -ACodec $info.ACodec `
                    -Hevc ([bool]$AllowHevc) -ForceMp4 ([bool]$ForceMp4)
    $items.Add([pscustomobject]@{
        File    = $f
        Info    = $info
        Plan    = $t.Plan
        KeepExt = $t.Ext
        Out     = Get-OutPath -Item ([pscustomobject]@{ File = $f; KeepExt = $t.Ext }) -Mode $t.Plan -OutputDir $OutputDir
    })
}
Write-Host ("策略统计: 原样重封装 {0} 个 | 仅换容器复制 {1} 个 | 视频直拷+转音频 {2} 个 | 重编码 {3} 个" -f
    @($items | Where-Object { $_.Plan -eq 'remux' }).Count,
    @($items | Where-Object { $_.Plan -eq 'recopy' }).Count,
    @($items | Where-Object { $_.Plan -like 'copyvideo*' }).Count,
    @($items | Where-Object { $_.Plan -eq 'transcode' }).Count)
Write-Host ''

# ---------------------------------------------------------------- 执行转换（并发）
$resultsByIndex = @{}
$running = New-Object System.Collections.Generic.List[object]
$next = 0
$overall = [System.Diagnostics.Stopwatch]::StartNew()
$planLabel = @{ 'remux' = '原样重封装'; 'recopy' = '仅换容器(复制)'; 'copyvideo' = '视频直拷+转音频';
                'copyvideo-mp4' = '换容器+视频直拷'; 'transcode' = '重编码';
                'transcode:cpu' = '重编码(CPU回退)'; 'skip' = '跳过' }

while ($next -lt $items.Count -or $running.Count -gt 0) {
    # 填满进程池
    while ($running.Count -lt $Parallel -and $next -lt $items.Count) {
        $it = $items[$next]
        $index = $next
        $next++
        if ((Test-Path -LiteralPath $it.Out) -and -not $Force) {
            $resultsByIndex[$index] = [pscustomobject]@{
                Ok = $true; Plan = 'skip'; Seconds = 0
                OutBytes = (Get-Item -LiteralPath $it.Out).Length; Note = ''; Ext = [IO.Path]::GetExtension($it.Out)
            }
            Write-Host "  [跳过] $($it.File.Name)  已存在 (使用 -Force 覆盖)" -ForegroundColor Yellow
            continue
        }
        $modes = Get-ModeSequence -Plan $it.Plan -HwEncoder $hw
        $out0  = Get-OutPath -Item $it -Mode $modes[0] -OutputDir $OutputDir
        $ffArgs = Build-FfmpegArgs -Src $it.File.FullName -Out $out0 -Mode $modes[0] -Ext ([IO.Path]::GetExtension($out0)) `
                                   -Preset $Preset -Crf $Crf -AudioBitrate $AudioBitrate -HwEncoder $hw
        $running.Add((Start-FfmpegTask -Ffmpeg $Ffmpeg -ArgList $ffArgs -Item $it -Index $index -ModeIndex 0 -Carry 0 -Out $out0))
    }

    Start-Sleep -Milliseconds 150

    # 回收已完成 / 超时的进程（用 ToArray 快照，避免迭代中修改集合）
    foreach ($h in $running.ToArray()) {
        if (-not $h.Proc.HasExited) {
            if ((([DateTime]::Now) - $h.Start).TotalSeconds -gt $TimeoutSec) {
                try { $h.Proc.Kill(); [void]$h.Proc.WaitForExit(5000) } catch { }
            } else { continue }
        }
        $h.Sw.Stop()
        $spent = $h.Carry + $h.Sw.Elapsed.TotalSeconds
        $code = -1
        try { $code = $h.Proc.ExitCode } catch { }
        $ok = $false
        if ($code -eq 0 -and (Test-Path -LiteralPath $h.Out)) {
            try { $ok = ((Get-Item -LiteralPath $h.Out).Length -gt 1024) } catch { $ok = $false }
        }
        $modes = Get-ModeSequence -Plan $h.Item.Plan -HwEncoder $hw
        if ($ok) {
            [void]$running.Remove($h)
            $null = Read-TaskLog -Task $h
            $size = (Get-Item -LiteralPath $h.Out).Length
            $resultsByIndex[$h.Index] = [pscustomobject]@{
                Ok = $true; Plan = $modes[$h.ModeIndex]; Seconds = $spent; OutBytes = $size
                Note = ''; Ext = [IO.Path]::GetExtension($h.Out)
            }
            Write-Host ("  [完成] {0}  {1}  {2}s  {3}MB" -f $h.Item.File.Name, $planLabel[$modes[$h.ModeIndex]],
                        [math]::Round($spent, 2), [math]::Round($size / 1MB, 2)) -ForegroundColor Green
            continue
        }
        # 失败：清理并按降级序列重试
        $log = Read-TaskLog -Task $h
        Remove-Item -LiteralPath $h.Out -Force -ErrorAction SilentlyContinue
        $mi = $h.ModeIndex + 1
        if ($mi -lt $modes.Count) {
            $outN = Get-OutPath -Item $h.Item -Mode $modes[$mi] -OutputDir $OutputDir
            $ffArgs = Build-FfmpegArgs -Src $h.Item.File.FullName -Out $outN -Mode $modes[$mi] -Ext ([IO.Path]::GetExtension($outN)) `
                                       -Preset $Preset -Crf $Crf -AudioBitrate $AudioBitrate -HwEncoder $hw
            $h2 = Start-FfmpegTask -Ffmpeg $Ffmpeg -ArgList $ffArgs -Item $h.Item -Index $h.Index -ModeIndex $mi -Carry $spent -Out $outN
            [void]$running.Remove($h)
            $running.Add($h2)
        } else {
            [void]$running.Remove($h)
            $resultsByIndex[$h.Index] = [pscustomobject]@{
                Ok = $false; Plan = $h.Item.Plan; Seconds = $spent; OutBytes = 0
                Note = $log; Ext = [IO.Path]::GetExtension($h.Out)
            }
            Write-Host ("  [失败] {0}  {1}s" -f $h.Item.File.Name, [math]::Round($spent, 2)) -ForegroundColor Red
        }
    }
}
$overall.Stop()

# ---------------------------------------------------------------- 组装结果
$rows = New-Object System.Collections.Generic.List[object]
for ($i = 0; $i -lt $items.Count; $i++) {
    $it = $items[$i]
    $r  = $resultsByIndex[$i]
    $srcMB = [math]::Round($it.File.Length / 1MB, 2)
    $outMB = [math]::Round($r.OutBytes / 1MB, 2)
    $ratio = if ($srcMB -gt 0 -and $outMB -gt 0) { '{0:P1}' -f ($outMB / $srcMB) } else { '' }
    $rtf = ''
    try {
        $durSec = [double]([TimeSpan]::Parse($it.Info.Duration).TotalSeconds)
        if ($durSec -gt 0 -and $r.Seconds -gt 0) { $rtf = [math]::Round($durSec / $r.Seconds, 2) }
    } catch { }
    $outExt = if ($r.Ext) { $r.Ext.TrimStart('.').ToUpper() } else { '' }
    $stream = ''
    if ($r.Ok -and $r.Plan -ne 'skip') {
        if ($r.Ext -in @('.mp4', '.m4v', '.mov')) {
            $stream = if (Test-FastStart -Path (Get-OutPath -Item $it -Mode $r.Plan -OutputDir $OutputDir)) { '是(moov前置)' } else { '否' }
        } elseif ($r.Ext -in @('.mkv', '.webm')) {
            $stream = '是(容器原生流式)'
        }
    }
    $status = if (-not $r.Ok) { '失败' } elseif ($r.Plan -eq 'skip') { '跳过(已存在)' } else { '成功' }
    $encoderUsed = switch -Wildcard ($r.Plan) {
        'remux'         { '复制(无损)' }
        'recopy'        { '复制(无损)' }
        'copyvideo*'    { if ($r.Ext -eq '.webm') { '视频复制 + Opus' } else { '视频复制 + AAC' } }
        'transcode'     { if ($hw) { $hw } else { 'libx264' } }
        'transcode:cpu' { 'libx264(回退)' }
        default         { '' }
    }
    $note = if ($r.Ok) { '' } else { (($r.Note -split "`r?`n" | Where-Object { $_.Trim() } | Select-Object -Last 2) -join ' / ') }
    $rows.Add([pscustomobject]@{
        序号 = $i + 1
        源文件 = $it.File.Name
        输出文件 = "$($it.File.BaseName)$($r.Ext)"
        输出容器 = $outExt
        分辨率 = $it.Info.Resolution
        时长 = $it.Info.Duration
        源视频编码 = $it.Info.VCodec
        源音频编码 = $it.Info.ACodec
        策略 = if ($r.Plan -eq 'skip') { '跳过' } else { $planLabel[$r.Plan] }
        输出编码 = $encoderUsed
        源大小MB = $srcMB
        输出大小MB = $outMB
        体积比 = $ratio
        转换耗时秒 = [math]::Round($r.Seconds, 2)
        实时倍数 = $rtf
        边下边播 = $stream
        状态 = $status
        备注 = $note
    })
}

# ---------------------------------------------------------------- 汇总
$okList   = @($rows | Where-Object { $_.状态 -eq '成功' })
$skipList = @($rows | Where-Object { $_.状态 -like '跳过*' })
$failCnt  = @($rows | Where-Object { $_.状态 -eq '失败' }).Count
function Get-Sum { param($Items, [string] $Prop) $s = ($Items | Measure-Object -Property $Prop -Sum).Sum; if ($null -eq $s) { 0 } else { $s } }
$sumSrc  = [math]::Round((Get-Sum $rows '源大小MB'), 2)
$sumOut  = [math]::Round((Get-Sum $rows '输出大小MB'), 2)
$sumTime = [math]::Round((Get-Sum $rows '转换耗时秒'), 2)
$avgRatio = if ($sumSrc -gt 0) { '{0:P1}' -f ($sumOut / $sumSrc) } else { '-' }
$cntRemux = @($rows | Where-Object { $_.策略 -eq '原样重封装' }).Count
$cntRecopy = @($rows | Where-Object { $_.策略 -eq '仅换容器(复制)' }).Count
$cntCopyV = @($rows | Where-Object { $_.策略 -like '视频直拷*' -or $_.策略 -like '换容器*' }).Count
$cntTrans = @($rows | Where-Object { $_.策略 -like '重编码*' }).Count

Write-Host ''
Write-Host '===================== 转换结果表 =====================' -ForegroundColor Cyan
$rows | Format-Table -AutoSize -Property 序号, 源文件, 输出容器, 策略, 输出编码, 输出大小MB, 转换耗时秒, 实时倍数, 边下边播, 状态
Write-Host ("总文件 $($rows.Count) 个 | 成功 $($okList.Count) | 失败 $failCnt | 跳过 $($skipList.Count)")
Write-Host ("策略分布: 原样重封装 $cntRemux | 仅换容器 $cntRecopy | 视频直拷+转音频 $cntCopyV | 重编码 $cntTrans")
Write-Host ("总耗时 {0} | 累计转换耗时 {1:N1}s" -f $overall.Elapsed.ToString('hh\:mm\:ss'), $sumTime)
Write-Host ("总体积 {0} MB -> {1} MB | 体积比 {2}" -f $sumSrc, $sumOut, $avgRatio)

# ---------------------------------------------------------------- 输出表格文件
Write-Utf8File -Path (Join-Path $OutputDir 'conversion_report.csv') -Content (($rows | ConvertTo-Csv -NoTypeInformation) -join "`r`n")

$md = New-Object System.Collections.Generic.List[string]
$md.Add('| # | 源文件 | 输出文件 | 输出容器 | 分辨率 | 时长 | 源视频/音频编码 | 策略 | 输出编码 | 源大小(MB) | 输出大小(MB) | 体积比 | 转换耗时(s) | 实时倍数 | 边下边播 | 状态 |')
$md.Add('|---|--------|----------|----------|--------|------|-----------------|------|----------|-----------|-------------|--------|-------------|----------|----------|------|')
foreach ($r in $rows) {
    $md.Add("| $($r.序号) | $($r.源文件) | $($r.输出文件) | $($r.输出容器) | $($r.分辨率) | $($r.时长) | $($r.源视频编码)/$($r.源音频编码) | $($r.策略) | $($r.输出编码) | $($r.源大小MB) | $($r.输出大小MB) | $($r.体积比) | $($r.转换耗时秒) | $($r.实时倍数) | $($r.边下边播) | $($r.状态) |")
}
$md.Add('')
$md.Add("**汇总**: 总文件 $($rows.Count) 个，成功 $($okList.Count)，失败 $failCnt，跳过 $($skipList.Count)；" +
        "策略分布 原样重封装 $cntRemux / 仅换容器 $cntRecopy / 视频直拷 $cntCopyV / 重编码 $cntTrans；" +
        "总耗时 $($overall.Elapsed.ToString('hh\:mm\:ss'))，累计转换耗时 ${sumTime}s；" +
        "总体积 ${sumSrc} MB -> ${sumOut} MB (体积比 $avgRatio)。")
$md.Add('')
$md.Add("**参数**: 保持原容器 $(if ($ForceMp4) { '否' } else { '是' })；重编码编码器 $(if ($hw) { $hw } else { 'libx264' }) (preset=$Preset, crf=$Crf)；音频 $AudioBitrate；并发 $Parallel。")
Write-Utf8File -Path (Join-Path $OutputDir 'conversion_report.md') -Content ($md -join "`r`n")

$trs = New-Object System.Collections.Generic.List[string]
foreach ($r in $rows) {
    $cls = switch -Wildcard ($r.状态) { '成功' { 'ok' } '失败' { 'bad' } default { 'skip' } }
    $pcls = switch -Wildcard ($r.策略) {
        '原样重封装'      { 'fast' }
        '仅换容器*'       { 'fast' }
        '视频直拷*'       { 'mid' }
        '换容器*'         { 'mid' }
        default           { 'slow' }
    }
    $trs.Add(@"
      <tr>
        <td>$($r.序号)</td>
        <td class="file">$($r.源文件)</td>
        <td class="file">$($r.输出文件)</td>
        <td>$($r.输出容器)</td>
        <td>$($r.分辨率)</td>
        <td>$($r.时长)</td>
        <td>$($r.源视频编码) / $($r.源音频编码)</td>
        <td><span class="tag $pcls">$($r.策略)</span></td>
        <td>$($r.输出编码)</td>
        <td class="num">$($r.源大小MB)</td>
        <td class="num">$($r.输出大小MB)</td>
        <td class="num">$($r.体积比)</td>
        <td class="num time">$($r.转换耗时秒)</td>
        <td class="num">$($r.实时倍数)</td>
        <td>$($r.边下边播)</td>
        <td><span class="tag $cls">$($r.状态)</span></td>
      </tr>
"@)
}
$html = @"
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>边下边播转换报告</title>
<style>
  :root { --bg:#0f1115; --card:#171a21; --line:#262b36; --txt:#e6e9ef; --sub:#9aa4b2; --acc:#4ea1ff; }
  * { box-sizing:border-box; }
  body { margin:0; padding:32px; background:var(--bg); color:var(--txt);
         font:14px/1.6 "Segoe UI", "Microsoft YaHei", system-ui, sans-serif; }
  h1 { margin:0 0 4px; font-size:22px; }
  .sub { color:var(--sub); font-size:13px; margin-bottom:24px; }
  .cards { display:flex; flex-wrap:wrap; gap:12px; margin-bottom:24px; }
  .card { flex:1 1 150px; background:var(--card); border:1px solid var(--line);
          border-radius:12px; padding:14px 16px; }
  .card .k { color:var(--sub); font-size:12px; }
  .card .v { font-size:20px; font-weight:600; margin-top:2px; }
  .card .v small { font-size:12px; font-weight:400; color:var(--sub); }
  .wrap { overflow-x:auto; }
  table { width:100%; border-collapse:collapse; background:var(--card);
          border:1px solid var(--line); border-radius:12px; overflow:hidden; }
  th, td { padding:9px 12px; border-bottom:1px solid var(--line); text-align:left; white-space:nowrap; }
  th { background:#1c2029; font-size:12px; color:var(--sub); font-weight:600; position:sticky; top:0; }
  tbody tr:hover { background:#1d222c; }
  tbody tr:last-child td { border-bottom:none; }
  td.file { color:var(--acc); font-family:Consolas, monospace; }
  td.num { text-align:right; font-family:Consolas, monospace; }
  td.time { color:#ffd479; font-weight:600; }
  .tag { padding:2px 8px; border-radius:999px; font-size:12px; }
  .tag.ok { background:rgba(46,204,113,.15); color:#2ecc71; }
  .tag.bad { background:rgba(231,76,60,.15); color:#e74c3c; }
  .tag.skip { background:rgba(154,164,178,.15); color:var(--sub); }
  .tag.fast { background:rgba(78,161,255,.15); color:#4ea1ff; }
  .tag.mid { background:rgba(255,212,121,.15); color:#ffd479; }
  .tag.slow { background:rgba(155,89,182,.18); color:#b98ce0; }
</style>
</head>
<body>
  <h1>边下边播转换报告</h1>
  <div class="sub">保持原容器 $(if ($ForceMp4) { '否' } else { '是' }) · 编码器 $(if ($hw) { $hw } else { 'libx264' }) (preset=$Preset, crf=$Crf) · 音频 $AudioBitrate · 并发 $Parallel · 生成于 $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</div>
  <div class="cards">
    <div class="card"><div class="k">文件总数</div><div class="v">$($rows.Count)</div></div>
    <div class="card"><div class="k">成功 / 失败</div><div class="v">$($okList.Count) <small>/ $failCnt / 跳过 $($skipList.Count)</small></div></div>
    <div class="card"><div class="k">总耗时</div><div class="v">$($overall.Elapsed.ToString('hh\:mm\:ss'))</div></div>
    <div class="card"><div class="k">累计转换耗时</div><div class="v">${sumTime}<small> s</small></div></div>
    <div class="card"><div class="k">策略分布</div><div class="v">$cntRemux <small>重封装 /</small> $cntRecopy <small>换容器 /</small> $cntCopyV <small>直拷 /</small> $cntTrans <small>重编</small></div></div>
    <div class="card"><div class="k">体积变化</div><div class="v">${sumSrc} <small>MB →</small> ${sumOut} <small>MB</small></div></div>
    <div class="card"><div class="k">体积比</div><div class="v">$avgRatio</div></div>
  </div>
  <div class="wrap">
  <table>
    <thead>
      <tr><th>#</th><th>源文件</th><th>输出文件</th><th>容器</th><th>分辨率</th><th>时长</th><th>源编码(视频/音频)</th><th>策略</th><th>输出编码</th><th>源大小(MB)</th><th>输出大小(MB)</th><th>体积比</th><th>转换耗时(s)</th><th>实时倍数</th><th>边下边播</th><th>状态</th></tr>
    </thead>
    <tbody>
$($trs -join "`r`n")
    </tbody>
  </table>
  </div>
</body>
</html>
"@
Write-Utf8File -Path (Join-Path $OutputDir 'conversion_report.html') -Content $html

Write-Host ''
Write-Host '报告已生成:' -ForegroundColor Cyan
Write-Host "  $(Join-Path $OutputDir 'conversion_report.csv')"
Write-Host "  $(Join-Path $OutputDir 'conversion_report.md')"
Write-Host "  $(Join-Path $OutputDir 'conversion_report.html')"
