<#
.SYNOPSIS
    将 videos 目录下的视频批量转换为「边下边播」的 AV1 (MP4) 文件，并记录每个视频的转换耗时，生成统计表格。

.DESCRIPTION
    输出特点（边下边播）：
      - 容器: MP4(AV1 + AAC)，浏览器/播放器原生支持
      - faststart 模式（默认）: -movflags +faststart，moov 前置，下载即可开始播放
      - fragmented 模式: -movflags +frag_keyframe+empty_moov+default_base_is_moof，
                         输出 fMP4，更适合 HTTP 流式边下边播 / MSE
      - 固定 GOP(-g 250)、tiles 分块，利于拖动与并行解码

    产物:
      av1\*.av1.mp4                 转换后的视频
      av1\conversion_report.csv     表格(CSV，Excel 可直接打开)
      av1\conversion_report.md      表格(Markdown)
      av1\conversion_report.html    表格(HTML，可直接浏览器打开)

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\convert_to_av1.ps1
    powershell -ExecutionPolicy Bypass -File .\convert_to_av1.ps1 -Limit 2          # 先跑 2 个试速度
    powershell -ExecutionPolicy Bypass -File .\convert_to_av1.ps1 -CpuUsed 6 -Crf 30 # 更慢但画质更好
    powershell -ExecutionPolicy Bypass -File .\convert_to_av1.ps1 -StreamingMode fragmented
#>

[CmdletBinding()]
param(
    [string] $InputDir   = '',
    [string] $OutputDir  = '',
    [string] $Ffmpeg     = '',

    # 可用编码器: libaom-av1(通用、慢) / av1_qsv(Intel 核显硬件编码、快)
    [ValidateSet('libaom-av1', 'av1_qsv')]
    [string] $Encoder = 'libaom-av1',

    # libaom-av1: 0(最慢/最好) ~ 8(最快/最差)
    [int]    $CpuUsed = 8,
    [int]    $Crf     = 32,

    # 边下边播模式
    [ValidateSet('faststart', 'fragmented')]
    [string] $StreamingMode = 'faststart',

    [string] $AudioBitrate = '128k',

    [string[]] $Extensions = @('.mp4', '.mkv', '.mov', '.avi', '.webm', '.flv',
                               '.wmv', '.mpg', '.mpeg', '.ts', '.3gp', '.ogv', '.m4v'),

    [int]    $Limit = 0,      # >0 时只转换前 N 个（用于试跑）
    [switch] $Force           # 已存在输出时也重新转换
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
if ([string]::IsNullOrWhiteSpace($OutputDir)) { $OutputDir = Join-Path $ScriptRoot 'av1' }
if ([string]::IsNullOrWhiteSpace($Ffmpeg))    { $Ffmpeg    = Join-Path $ScriptRoot 'tools\ffmpeg.exe' }

# ---------------------------------------------------------------- 工具函数
function Write-Utf8File {
    param([string] $Path, [string] $Content)
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $enc = New-Object System.Text.UTF8Encoding($true)   # 带 BOM，Excel/记事本打开不乱码
    [System.IO.File]::WriteAllText($Path, $Content, $enc)
}

function Get-VideoInfo {
    param([string] $Ffmpeg, [string] $File)
    $info = [pscustomobject]@{ Duration = ''; Resolution = ''; VCodec = '' }
    try {
        # ffmpeg 把信息写到 stderr，且退出码非 0，临时放宽 ErrorAction 避免被当成终止错误
        $oldEap = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        $raw = (& $Ffmpeg -hide_banner -i $File 2>&1 | Out-String)
        $ErrorActionPreference = $oldEap
        $m = [regex]::Match($raw, 'Duration:\s*([0-9:.]+)')
        if ($m.Success) { $info.Duration = $m.Groups[1].Value }
        foreach ($line in ($raw -split "`r?`n")) {
            if ($line -match 'Stream .*Video:' ) {
                $info.VCodec = ([regex]::Match($line, 'Video:\s*([^,]+)')).Groups[1].Value.Trim()
                $r = [regex]::Match($line, '(\d{2,5})x(\d{2,5})')
                if ($r.Success) { $info.Resolution = "$($r.Groups[1].Value)x$($r.Groups[2].Value)" }
                break
            }
        }
    } catch { }
    return $info
}

function Get-EncoderArgs {
    param([string] $Encoder, [int] $Crf, [int] $CpuUsed)
    if ($Encoder -eq 'av1_qsv') {
        # Intel QSV 硬件 AV1：用 global_quality 控制质量(越小越好)
        return @('-c:v', 'av1_qsv', '-global_quality', '30', '-preset', 'medium', '-g', '250')
    }
    # libaom-av1 软件编码：cpu-used 越大越快
    return @('-c:v', 'libaom-av1', '-crf', "$Crf", '-b:v', '0',
             '-cpu-used', "$CpuUsed", '-row-mt', '1', '-tiles', '2x2',
             '-g', '250', '-threads', '0')
}

# ---------------------------------------------------------------- 前置检查
if (-not (Test-Path -LiteralPath $Ffmpeg)) { throw "未找到 ffmpeg: $Ffmpeg" }
if (-not (Test-Path -LiteralPath $InputDir)) { throw "未找到输入目录: $InputDir" }
if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

$movflags = if ($StreamingMode -eq 'fragmented') {
    '+frag_keyframe+empty_moov+default_base_is_moof'
} else {
    '+faststart'
}

$files = @(Get-ChildItem -LiteralPath $InputDir -File |
           Where-Object { $Extensions -contains $_.Extension.ToLower() } |
           Sort-Object Name)
if ($Limit -gt 0) { $files = @($files | Select-Object -First $Limit) }
if ($files.Count -eq 0) { throw "输入目录中没有匹配的视频文件: $InputDir" }

Write-Host ''
Write-Host "================ AV1 批量转换 ================" -ForegroundColor Cyan
Write-Host "输入目录    : $InputDir"
Write-Host "输出目录    : $OutputDir"
Write-Host "编码器      : $Encoder  (CRF=$Crf, cpu-used=$CpuUsed)"
Write-Host "边下边播模式: $StreamingMode  (movflags $movflags)"
Write-Host "待转换文件  : $($files.Count) 个"
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ''

$results = New-Object System.Collections.Generic.List[object]
$overall = [System.Diagnostics.Stopwatch]::StartNew()

for ($i = 0; $i -lt $files.Count; $i++) {
    $src    = $files[$i]
    $index  = $i + 1
    $outName = "$($src.BaseName).av1.mp4"
    $outPath = Join-Path $OutputDir $outName
    $srcMB   = [math]::Round($src.Length / 1MB, 2)

    Write-Host "[$index/$($files.Count)] $($src.Name)" -NoNewline

    if ((Test-Path -LiteralPath $outPath) -and -not $Force) {
        Write-Host "  -> 已存在，跳过 (使用 -Force 覆盖)" -ForegroundColor Yellow
        $results.Add([pscustomobject]@{
            序号 = $index; 源文件 = $src.Name; 分辨率 = ''; 时长 = ''
            源大小MB = $srcMB; 输出大小MB = [math]::Round((Get-Item -LiteralPath $outPath).Length / 1MB, 2)
            压缩率 = ''; 转换耗时秒 = 0; 实时倍数 = ''; 状态 = '跳过(已存在)'; 备注 = ''
        })
        continue
    }

    $info = Get-VideoInfo -Ffmpeg $Ffmpeg -File $src.FullName

    $args = @('-y', '-hide_banner', '-nostdin', '-nostats', '-loglevel', 'error',
              '-i', $src.FullName,
              '-map', '0:v:0', '-map', '0:a:0?', '-sn', '-dn')
    $args += Get-EncoderArgs -Encoder $Encoder -Crf $Crf -CpuUsed $CpuUsed
    $args += @('-pix_fmt', 'yuv420p',
               '-c:a', 'aac', '-b:a', $AudioBitrate,
               '-movflags', $movflags,
               '-map_metadata', '0',
               '-f', 'mp4', $outPath)

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $oldEap = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $ffOut = (& $Ffmpeg @args 2>&1 | Out-String)
        $code = $LASTEXITCODE
    } catch {
        $ffOut = $_.Exception.Message
        $code  = -1
    } finally {
        $ErrorActionPreference = $oldEap
    }
    $sw.Stop()

    $ok = ($code -eq 0) -and (Test-Path -LiteralPath $outPath)
    if ($ok) {
        $outMB = [math]::Round((Get-Item -LiteralPath $outPath).Length / 1MB, 2)
        $ratio = if ($srcMB -gt 0) { '{0:P1}' -f ($outMB / $srcMB) } else { '' }
        # 实时倍数 = 视频时长 / 转换耗时，>1 表示比播放速度还快
        $rtf = ''
        try {
            $durSec = [double]([TimeSpan]::Parse($info.Duration).TotalSeconds)
            if ($durSec -gt 0) { $rtf = [math]::Round($durSec / $sw.Elapsed.TotalSeconds, 2) }
        } catch { }
        Write-Host ("  -> 完成 {0:N1}s  {1}MB => {2}MB" -f $sw.Elapsed.TotalSeconds, $srcMB, $outMB) -ForegroundColor Green
        $results.Add([pscustomobject]@{
            序号 = $index; 源文件 = $src.Name; 分辨率 = $info.Resolution; 时长 = $info.Duration
            源大小MB = $srcMB; 输出大小MB = $outMB; 压缩率 = $ratio
            转换耗时秒 = [math]::Round($sw.Elapsed.TotalSeconds, 2); 实时倍数 = $rtf
            状态 = '成功'; 备注 = $info.VCodec
        })
    } else {
        $msg = (($ffOut -split "`r?`n" | Where-Object { $_.Trim() } | Select-Object -Last 2) -join ' / ')
        Write-Host ("  -> 失败 {0:N1}s  {1}" -f $sw.Elapsed.TotalSeconds, $msg) -ForegroundColor Red
        $results.Add([pscustomobject]@{
            序号 = $index; 源文件 = $src.Name; 分辨率 = $info.Resolution; 时长 = $info.Duration
            源大小MB = $srcMB; 输出大小MB = 0; 压缩率 = ''
            转换耗时秒 = [math]::Round($sw.Elapsed.TotalSeconds, 2); 实时倍数 = ''
            状态 = '失败'; 备注 = $msg
        })
        if (Test-Path -LiteralPath $outPath) { Remove-Item -LiteralPath $outPath -Force -ErrorAction SilentlyContinue }
    }
}
$overall.Stop()

# ---------------------------------------------------------------- 汇总
$okList  = @($results | Where-Object { $_.状态 -eq '成功' })
$failCnt = @($results | Where-Object { $_.状态 -eq '失败' }).Count
$skipCnt = @($results | Where-Object { $_.状态 -like '跳过*' }).Count
function Get-Sum { param($Items, [string] $Prop) $s = ($Items | Measure-Object -Property $Prop -Sum).Sum; if ($null -eq $s) { 0 } else { $s } }
$sumSrc  = [math]::Round((Get-Sum $results '源大小MB'), 2)
$sumOut  = [math]::Round((Get-Sum $okList  '输出大小MB'), 2)
$sumTime = [math]::Round((Get-Sum $results '转换耗时秒'), 2)
$avgRatio = if ($sumSrc -gt 0) { '{0:P1}' -f ($sumOut / $sumSrc) } else { '-' }

Write-Host ''
Write-Host '==================== 转换结果表 ====================' -ForegroundColor Cyan
$results | Format-Table -AutoSize -Property 序号, 源文件, 分辨率, 时长, 源大小MB, 输出大小MB, 压缩率, 转换耗时秒, 实时倍数, 状态
Write-Host "总文件 $($files.Count) 个 | 成功 $($okList.Count) | 失败 $failCnt | 跳过 $skipCnt"
Write-Host ("总耗时 {0} | 累计转换耗时 {1:N1}s" -f $overall.Elapsed.ToString('hh\:mm\:ss'), $sumTime)
Write-Host ("总体积 {0} MB -> {1} MB | 压缩率 {2}" -f $sumSrc, $sumOut, $avgRatio)

# ---------------------------------------------------------------- 输出表格文件
$csv = $results | ConvertTo-Csv -NoTypeInformation
Write-Utf8File -Path (Join-Path $OutputDir 'conversion_report.csv') -Content ($csv -join "`r`n")

$md = New-Object System.Collections.Generic.List[string]
$md.Add('| # | 源文件 | 分辨率 | 时长 | 源大小(MB) | 输出大小(MB) | 压缩率 | 转换耗时(s) | 实时倍数 | 状态 |')
$md.Add('|---|--------|--------|------|-----------|-------------|--------|------------|----------|------|')
foreach ($r in $results) {
    $md.Add("| $($r.序号) | $($r.源文件) | $($r.分辨率) | $($r.时长) | $($r.源大小MB) | $($r.输出大小MB) | $($r.压缩率) | $($r.转换耗时秒) | $($r.实时倍数) | $($r.状态) |")
}
$md.Add('')
$md.Add("**汇总**: 总文件 $($files.Count) 个，成功 $($okList.Count)，失败 $failCnt，跳过 $skipCnt；" +
        "总耗时 $($overall.Elapsed.ToString('hh\:mm\:ss'))，累计转换耗时 ${sumTime}s；" +
        "总体积 ${sumSrc} MB -> ${sumOut} MB (压缩率 $avgRatio)。")
$md.Add('')
$md.Add("**参数**: 编码器 $Encoder (CRF=$Crf, cpu-used=$CpuUsed)，音频 AAC $AudioBitrate，边下边播模式 $StreamingMode。")
Write-Utf8File -Path (Join-Path $OutputDir 'conversion_report.md') -Content ($md -join "`r`n")

$rows = New-Object System.Collections.Generic.List[string]
foreach ($r in $results) {
    $cls = switch -Wildcard ($r.状态) { '成功' { 'ok' } '失败' { 'bad' } default { 'skip' } }
    $rows.Add(@"
      <tr>
        <td>$($r.序号)</td>
        <td class="file">$($r.源文件)</td>
        <td>$($r.分辨率)</td>
        <td>$($r.时长)</td>
        <td class="num">$($r.源大小MB)</td>
        <td class="num">$($r.输出大小MB)</td>
        <td class="num">$($r.压缩率)</td>
        <td class="num time">$($r.转换耗时秒)</td>
        <td class="num">$($r.实时倍数)</td>
        <td><span class="tag $cls">$($r.状态)</span></td>
      </tr>
"@)
}
$html = @"
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>AV1 转换报告</title>
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
</style>
</head>
<body>
  <h1>AV1 转换报告</h1>
  <div class="sub">编码器 $Encoder (CRF=$Crf, cpu-used=$CpuUsed) · 音频 AAC $AudioBitrate · 边下边播 $StreamingMode · 生成于 $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</div>
  <div class="cards">
    <div class="card"><div class="k">文件总数</div><div class="v">$($files.Count)</div></div>
    <div class="card"><div class="k">成功 / 失败</div><div class="v">$($okList.Count) <small>/ $failCnt / 跳过 $skipCnt</small></div></div>
    <div class="card"><div class="k">总耗时</div><div class="v">$($overall.Elapsed.ToString('hh\:mm\:ss'))</div></div>
    <div class="card"><div class="k">累计转换耗时</div><div class="v">${sumTime}<small> s</small></div></div>
    <div class="card"><div class="k">体积变化</div><div class="v">${sumSrc} <small>MB →</small> ${sumOut} <small>MB</small></div></div>
    <div class="card"><div class="k">压缩率</div><div class="v">$avgRatio</div></div>
  </div>
  <table>
    <thead>
      <tr><th>#</th><th>源文件</th><th>分辨率</th><th>时长</th><th>源大小(MB)</th><th>输出大小(MB)</th><th>压缩率</th><th>转换耗时(s)</th><th>实时倍数</th><th>状态</th></tr>
    </thead>
    <tbody>
$($rows -join "`r`n")
    </tbody>
  </table>
</body>
</html>
"@
Write-Utf8File -Path (Join-Path $OutputDir 'conversion_report.html') -Content $html

Write-Host ''
Write-Host "报告已生成:" -ForegroundColor Cyan
Write-Host "  $(Join-Path $OutputDir 'conversion_report.csv')"
Write-Host "  $(Join-Path $OutputDir 'conversion_report.md')"
Write-Host "  $(Join-Path $OutputDir 'conversion_report.html')"
