#Requires -Version 5.1
<#
.SYNOPSIS
    批量生成多格式测试视频：每个文件 10 秒，分辨率 720p(1280x720) 与 1080p(1920x1080)，
    画面为 lavfi testsrc2 彩条 + 计时器，音轨为 440Hz 正弦音。

.PARAMETER Ffmpeg
    ffmpeg 可执行文件路径，默认取 ..\tools\ffmpeg.exe

.PARAMETER OutDir
    输出目录，默认脚本所在目录（videos）

.PARAMETER Duration
    每个视频时长（秒），默认 10

.PARAMETER Fps
    帧率，默认 30

.PARAMETER Resolutions
    需要生成的分辨率列表，默认 720p、1080p

.PARAMETER Formats
    需要生成的格式名（支持通配符），默认全部

.PARAMETER SkipVerify
    跳过生成后的解码自检

.PARAMETER VerifyOnly
    只校验已存在的文件（时长 + 完整解码自检），不重新生成

.EXAMPLE
    .\generate_test_videos.ps1

.EXAMPLE
    .\generate_test_videos.ps1 -Formats mp4_h264,webm_vp9 -Resolutions 720p
#>
[CmdletBinding()]
param(
    [string]$Ffmpeg = '',
    [string]$OutDir = '',
    [int]$Duration = 10,
    [int]$Fps = 30,
    [string[]]$Resolutions = @('720p', '1080p'),
    [string[]]$Formats = @('*'),
    [switch]$SkipVerify,
    [switch]$VerifyOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------- 脚本根目录 ----
# 某些宿主下 $PSScriptRoot 可能为空，回退到 $MyInvocation
$scriptRoot = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($scriptRoot)) {
    $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if ([string]::IsNullOrWhiteSpace($Ffmpeg)) {
    $Ffmpeg = Join-Path (Split-Path -Parent $scriptRoot) 'tools\ffmpeg.exe'
}
if ([string]::IsNullOrWhiteSpace($OutDir)) {
    $OutDir = $scriptRoot
}

# ---------------------------------------------------------------- 基础校验 ----
$ffmpegItem = Resolve-Path -LiteralPath $Ffmpeg -ErrorAction SilentlyContinue
if (-not $ffmpegItem) { throw "找不到 ffmpeg：$Ffmpeg" }
$ffmpeg = $ffmpegItem.Path

if (-not (Test-Path -LiteralPath $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
}
$OutDir = (Resolve-Path -LiteralPath $OutDir).Path

# ---------------------------------------------------------------- 分辨率表 ----
$ResolutionMap = [ordered]@{
    '720p'  = @{ W = 1280; H = 720  }
    '1080p' = @{ W = 1920; H = 1080 }
}

# ---------------------------------------------------------------- 格式清单 ----
# Name     : 文件名片段 / -Formats 过滤关键字
# Ext      : 输出扩展名
# Video    : 视频编码参数（可含 {BITRATE} 占位符）
# Audio    : 音频编码参数
# Extra    : 容器/封装相关的额外参数
# Bitrate  : 按分辨率给出的目标码率（仅需 CBR 的编码器使用），失败重试时会减半
# Vf       : 可选的视频滤镜（如 Theora 要求尺寸为 16 的倍数）
# Note     : 说明（仅用于输出展示）
$FormatList = @(
    [pscustomobject]@{
        Name    = 'mp4_h264'
        Ext     = 'mp4'
        Container = 'MP4 / H.264 + AAC'
        Video   = @('-c:v', 'libx264', '-preset', 'medium', '-crf', '23', '-pix_fmt', 'yuv420p')
        Audio   = @('-c:a', 'aac', '-b:a', '128k')
        Extra   = @('-movflags', '+faststart')
        Bitrate = $null
        Vf      = $null
        Note    = '最通用的组合'
    },
    [pscustomobject]@{
        Name    = 'mp4_hevc'
        Ext     = 'mp4'
        Container = 'MP4 / H.265 HEVC + AAC'
        Video   = @('-c:v', 'libx265', '-preset', 'medium', '-crf', '28', '-pix_fmt', 'yuv420p', '-tag:v', 'hvc1')
        Audio   = @('-c:a', 'aac', '-b:a', '128k')
        Extra   = @('-movflags', '+faststart')
        Bitrate = $null
        Vf      = $null
        Note    = 'hvc1 tag 提升 Apple 生态兼容性'
    },
    [pscustomobject]@{
        Name    = 'mkv_h264'
        Ext     = 'mkv'
        Container = 'Matroska / H.264 + AAC'
        Video   = @('-c:v', 'libx264', '-preset', 'medium', '-crf', '23', '-pix_fmt', 'yuv420p')
        Audio   = @('-c:a', 'aac', '-b:a', '128k')
        Extra   = @()
        Bitrate = $null
        Vf      = $null
        Note    = ''
    },
    [pscustomobject]@{
        Name    = 'mkv_vp9'
        Ext     = 'mkv'
        Container = 'Matroska / VP9 + Opus'
        Video   = @('-c:v', 'libvpx-vp9', '-crf', '32', '-b:v', '0', '-deadline', 'good', '-cpu-used', '4', '-row-mt', '1', '-pix_fmt', 'yuv420p')
        Audio   = @('-c:a', 'libopus', '-b:a', '96k')
        SampleRate = '48000'   # libopus 仅支持 48kHz
        Extra   = @()
        Bitrate = $null
        Vf      = $null
        Note    = ''
    },
    [pscustomobject]@{
        Name    = 'webm_vp9'
        Ext     = 'webm'
        Container = 'WebM / VP9 + Opus'
        Video   = @('-c:v', 'libvpx-vp9', '-crf', '32', '-b:v', '0', '-deadline', 'good', '-cpu-used', '4', '-row-mt', '1', '-pix_fmt', 'yuv420p')
        Audio   = @('-c:a', 'libopus', '-b:a', '96k')
        SampleRate = '48000'   # libopus 仅支持 48kHz
        Extra   = @()
        Bitrate = $null
        Vf      = $null
        Note    = ''
    },
    [pscustomobject]@{
        Name    = 'webm_vp8'
        Ext     = 'webm'
        Container = 'WebM / VP8 + Vorbis'
        Video   = @('-c:v', 'libvpx', '-crf', '10', '-b:v', '0', '-deadline', 'good', '-cpu-used', '4', '-pix_fmt', 'yuv420p')
        Audio   = @('-c:a', 'libvorbis', '-q:a', '4')
        Extra   = @()
        Bitrate = $null
        Vf      = $null
        Note    = ''
    },
    [pscustomobject]@{
        Name    = 'avi_mpeg4'
        Ext     = 'avi'
        Container = 'AVI / MPEG-4 + MP3'
        Video   = @('-c:v', 'mpeg4', '-q:v', '4', '-pix_fmt', 'yuv420p')
        Audio   = @('-c:a', 'libmp3lame', '-b:a', '128k')
        Extra   = @()
        Bitrate = $null
        Vf      = $null
        Note    = 'AVI 内 AAC 兼容性差，故用 MP3'
    },
    [pscustomobject]@{
        Name    = 'mov_h264'
        Ext     = 'mov'
        Container = 'QuickTime / H.264 + AAC'
        Video   = @('-c:v', 'libx264', '-preset', 'medium', '-crf', '23', '-pix_fmt', 'yuv420p')
        Audio   = @('-c:a', 'aac', '-b:a', '128k')
        Extra   = @('-movflags', '+faststart')
        Bitrate = $null
        Vf      = $null
        Note    = ''
    },
    [pscustomobject]@{
        Name    = 'mov_prores'
        Ext     = 'mov'
        Container = 'QuickTime / Apple ProRes + PCM'
        # ProRes 为帧内编码，码率极高；默认 profile 0 (Proxy) 控制体积，可改为 1/2/3
        Video   = @('-c:v', 'prores_ks', '-profile:v', '0')
        Audio   = @('-c:a', 'pcm_s16le')
        Extra   = @()
        Bitrate = $null
        Vf      = $null
        Note    = '不强制 pix_fmt，由编码器自动选 yuv422p10le'
    },
    [pscustomobject]@{
        Name    = 'ts_h264'
        Ext     = 'ts'
        Container = 'MPEG-TS / H.264 + AAC'
        Video   = @('-c:v', 'libx264', '-preset', 'medium', '-crf', '23', '-pix_fmt', 'yuv420p')
        Audio   = @('-c:a', 'aac', '-b:a', '128k')
        Extra   = @('-f', 'mpegts')
        Bitrate = $null
        Vf      = $null
        Note    = '可用于 HLS 切片测试'
    },
    [pscustomobject]@{
        Name    = 'flv_h264'
        Ext     = 'flv'
        Container = 'FLV / H.264 + MP3'
        Video   = @('-c:v', 'libx264', '-preset', 'veryfast', '-crf', '23', '-pix_fmt', 'yuv420p')
        Audio   = @('-c:a', 'libmp3lame', '-b:a', '128k')
        Extra   = @('-f', 'flv')
        Bitrate = $null
        Vf      = $null
        Note    = ''
    },
    [pscustomobject]@{
        Name    = 'asf_wmv2'
        Ext     = 'wmv'
        Container = 'ASF / WMV2 + WMA2'
        Video   = @('-c:v', 'wmv2', '-b:v', '{BITRATE}', '-pix_fmt', 'yuv420p')
        Audio   = @('-c:a', 'wmav2', '-b:a', '128k')
        Extra   = @('-f', 'asf')
        Bitrate = @{ '720p' = '2M'; '1080p' = '4M' }
        Vf      = $null
        Note    = 'WMV 编码器不支持 CRF，需显式码率'
    },
    [pscustomobject]@{
        Name    = 'mpg_mpeg2'
        Ext     = 'mpg'
        Container = 'MPEG-PS / MPEG-2 + MP2'
        Video   = @('-c:v', 'mpeg2video', '-b:v', '{BITRATE}', '-maxrate', '{BITRATE}', '-bufsize', '4M', '-pix_fmt', 'yuv420p')
        Audio   = @('-c:a', 'mp2', '-b:a', '192k')
        Extra   = @('-f', 'mpeg')
        Bitrate = @{ '720p' = '4M'; '1080p' = '8M' }
        Vf      = $null
        Note    = 'MPEG-PS 对码率敏感，失败会自动减半码率重试'
    },
    [pscustomobject]@{
        Name    = '3gp_mpeg4'
        Ext     = '3gp'
        Container = '3GP / MPEG-4 + AAC'
        # H.263 编码器最高只支持 CIF(352x288)，无法编码 720p/1080p，故改用 MPEG-4
        Video   = @('-c:v', 'mpeg4', '-b:v', '{BITRATE}', '-pix_fmt', 'yuv420p')
        Audio   = @('-c:a', 'aac', '-b:a', '64k')
        Extra   = @('-f', '3gp')
        Bitrate = @{ '720p' = '1500k'; '1080p' = '2500k' }
        Vf      = $null
        Note    = '原计划 H.263，因不支持 720p/1080p 改用 MPEG-4'
    },
    [pscustomobject]@{
        Name    = 'ogv_theora'
        Ext     = 'ogv'
        Container = 'Ogg / Theora + Vorbis'
        Video   = @('-c:v', 'libtheora', '-q:v', '6', '-pix_fmt', 'yuv420p')
        Audio   = @('-c:a', 'libvorbis', '-q:a', '4')
        Extra   = @('-f', 'ogg')
        Bitrate = $null
        # Theora 要求宽高均为 16 的倍数：1280x720 满足，1920x1080 会缩到 1920x1072
        Vf      = 'scale=trunc(iw/16)*16:trunc(ih/16)*16'
        Note    = '1080p 实际编码高度 1072（Theora 仅支持 16 的倍数）'
    }
)

# ---------------------------------------------------------------- 工具函数 ----
function Get-EntryProp {
    # 兼容 StrictMode：属性不存在时返回默认值
    param([object]$Obj, [string]$Name, [object]$Default)
    $prop = $Obj.PSObject.Properties[$Name]
    if ($null -ne $prop) { return $prop.Value }
    return $Default
}

function Expand-BitrateToken {
    param([string[]]$Args_, [string]$Bitrate)
    if (-not $Bitrate) { return @($Args_) }
    return @($Args_ | ForEach-Object { $_ -replace '\{BITRATE\}', $Bitrate })
}

function Get-HalfBitrate {
    param([string]$Bitrate)
    if ($Bitrate -match '^([0-9]+(?:\.[0-9]+)?)([kKMm]?)$') {
        $value = [double]$Matches[1] / 2
        return ('{0}{1}' -f $value.ToString('0.##'), $Matches[2])
    }
    return $Bitrate
}

function Invoke-Encode {
    param(
        [string]$Bin,
        [string[]]$Base,
        [string[]]$Video,
        [string[]]$Audio,
        [string[]]$Extra,
        [string]$AudioRate,
        [string]$Out
    )
    $argList = @('-hide_banner', '-nostdin', '-loglevel', 'error', '-stats', '-y')
    $argList += $Base
    $argList += $Video
    $argList += $Audio
    $argList += $Extra
    $argList += @('-t', "$Duration", '-ar', $AudioRate, '-ac', '2', $Out)

    $oldEap = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'   # 原生命令写 stderr 时不希望被 -ErrorAction Stop 中断
    $log = (& $Bin @argList 2>&1 | Out-String)
    $code = $LASTEXITCODE
    $ErrorActionPreference = $oldEap
    return [pscustomobject]@{ Code = $code; Log = $log }
}

function Get-MediaDuration {
    param([string]$Bin, [string]$Path)
    $oldEap = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $info = (& $Bin -hide_banner -nostdin -i $Path 2>&1 | Out-String)
    $ErrorActionPreference = $oldEap
    if ($info -match 'Duration:\s*([0-9]+):([0-9]+):([0-9]+\.?[0-9]*)') {
        return [math]::Round([int]$Matches[1] * 3600 + [int]$Matches[2] * 60 + [double]$Matches[3], 2)
    }
    return $null
}

function Test-Decodable {
    param([string]$Bin, [string]$Path)
    $oldEap = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $log = (& $Bin -hide_banner -nostdin -v error -i $Path -f null - 2>&1 | Out-String)
    $code = $LASTEXITCODE
    $ErrorActionPreference = $oldEap
    if ($code -ne 0) { return [pscustomobject]@{ Ok = $false; Message = $log.Trim() } }
    if ($log.Trim()) { return [pscustomobject]@{ Ok = $false; Message = $log.Trim() } }
    return [pscustomobject]@{ Ok = $true; Message = '' }
}

# ---------------------------------------------------------------- 主流程 ------
# powershell -File 调用时数组参数可能以 "a,b" 单个字符串传入，这里统一按逗号拆分
$formatPatterns = @($Formats | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
$resolutionKeys = @($Resolutions | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })

$selected = @($FormatList | Where-Object {
    $name = $_.Name
    @($formatPatterns | Where-Object { $name -like $_ }).Count -gt 0
})

$results = New-Object System.Collections.Generic.List[object]
$resList = @($resolutionKeys | Where-Object { $ResolutionMap.Contains($_) })
$total = @($selected).Count * $resList.Count
$index = 0

Write-Host "ffmpeg : $ffmpeg"
Write-Host "输出目录: $OutDir"
Write-Host ("选中格式: " + (($selected | ForEach-Object { $_.Name }) -join ', '))
Write-Host ("分辨率  : " + ($resList -join ', '))
if ($VerifyOnly) { Write-Host "共 $total 个文件待校验`n" } else { Write-Host "共 $total 个文件待生成`n" }

foreach ($fmt in $selected) {
    foreach ($resKey in $resList) {
        if (-not $ResolutionMap.Contains($resKey)) {
            Write-Warning "未知分辨率 $resKey，已跳过"
            continue
        }
        $index++
        $W = $ResolutionMap[$resKey].W
        $H = $ResolutionMap[$resKey].H
        $outName = "test_${resKey}_$($fmt.Name).$($fmt.Ext)"
        $outPath = Join-Path $OutDir $outName

        if ($VerifyOnly) {
            # 只校验已存在的文件：时长解析 + 完整解码自检
            $vsw = [System.Diagnostics.Stopwatch]::StartNew()
            $vrec = [pscustomobject]@{
                File = $outName; Format = $fmt.Container; Resolution = "$W x $H"
                SizeMB = 0; Duration = $null; Seconds = 0; Status = ''; Detail = ''
            }
            if (Test-Path -LiteralPath $outPath) {
                $vrec.SizeMB = [math]::Round((Get-Item -LiteralPath $outPath).Length / 1MB, 2)
                $vrec.Duration = Get-MediaDuration -Bin $ffmpeg -Path $outPath
                $check = Test-Decodable -Bin $ffmpeg -Path $outPath
                if ($check.Ok) {
                    $vrec.Status = 'OK'
                    Write-Host ("      校验通过 {0} MB / {1}s" -f $vrec.SizeMB, $vrec.Duration) -ForegroundColor Green
                }
                else {
                    $vrec.Status = '解码告警'
                    $vrec.Detail = ($check.Message -split "`n" | Select-Object -First 3) -join ' | '
                    Write-Host ("      解码告警: {0}" -f $vrec.Detail) -ForegroundColor Yellow
                }
            }
            else {
                $vrec.Status = '缺失'
                Write-Host "      文件缺失" -ForegroundColor Red
            }
            $vsw.Stop()
            $vrec.Seconds = [math]::Round($vsw.Elapsed.TotalSeconds, 1)
            $results.Add($vrec)
            continue
        }

        $bitrate = $null
        if ($fmt.Bitrate) { $bitrate = $fmt.Bitrate[$resKey] }
        $audioRate = Get-EntryProp -Obj $fmt -Name 'SampleRate' -Default '44100'

        $base = @(
            '-f', 'lavfi', '-i', "testsrc2=size=${W}x${H}:rate=${Fps}:duration=${Duration}",
            '-f', 'lavfi', '-i', "sine=frequency=440:duration=${Duration}",
            '-map', '0:v:0', '-map', '1:a:0'
        )
        if ($fmt.Vf) { $base += @('-vf', $fmt.Vf) }

        $video = Expand-BitrateToken -Args_ $fmt.Video -Bitrate $bitrate
        $audio = $fmt.Audio
        $extra = Expand-BitrateToken -Args_ $fmt.Extra -Bitrate $bitrate

        Write-Host "[$index/$total] $outName  ($($fmt.Container))" -ForegroundColor Cyan

        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $res = Invoke-Encode -Bin $ffmpeg -Base $base -Video $video -Audio $audio -Extra $extra -AudioRate $audioRate -Out $outPath
        $attempt = 1
        $usedBitrate = $bitrate

        if (($res.Code -ne 0) -and $bitrate) {
            # 码率型编码器失败时减半码率重试一次（MPEG-PS 常见 buffer underflow）
            $half = Get-HalfBitrate -Bitrate $bitrate
            Write-Host "      失败，码率 $bitrate -> $half 重试" -ForegroundColor Yellow
            $video = Expand-BitrateToken -Args_ $fmt.Video -Bitrate $half
            $extra = Expand-BitrateToken -Args_ $fmt.Extra -Bitrate $half
            $res = Invoke-Encode -Bin $ffmpeg -Base $base -Video $video -Audio $audio -Extra $extra -AudioRate $audioRate -Out $outPath
            $attempt = 2
            $usedBitrate = $half
        }
        $sw.Stop()

        $record = [pscustomobject]@{
            File       = $outName
            Format     = $fmt.Container
            Resolution = "$W x $H"
            SizeMB     = 0
            Duration   = $null
            Seconds    = [math]::Round($sw.Elapsed.TotalSeconds, 1)
            Status     = ''
            Detail     = ''
        }

        if ($res.Code -eq 0 -and (Test-Path -LiteralPath $outPath)) {
            $record.SizeMB = [math]::Round((Get-Item -LiteralPath $outPath).Length / 1MB, 2)
            $record.Duration = Get-MediaDuration -Bin $ffmpeg -Path $outPath

            if (-not $SkipVerify) {
                $check = Test-Decodable -Bin $ffmpeg -Path $outPath
                if ($check.Ok) {
                    $record.Status = 'OK'
                    Write-Host ("      完成 {0} MB / {1}s / 编码耗时 {2}s" -f $record.SizeMB, $record.Duration, $record.Seconds) -ForegroundColor Green
                }
                else {
                    $record.Status = '解码告警'
                    $record.Detail = ($check.Message -split "`n" | Select-Object -First 3) -join ' | '
                    Write-Host ("      解码自检告警: {0}" -f $record.Detail) -ForegroundColor Yellow
                }
            }
            else {
                $record.Status = 'OK'
                Write-Host ("      完成 {0} MB / {1}s / 编码耗时 {2}s" -f $record.SizeMB, $record.Duration, $record.Seconds) -ForegroundColor Green
            }
        }
        else {
            $record.Status = '失败'
            $record.Detail = (($res.Log -split "`n" | Where-Object { $_.Trim() } | Select-Object -Last 3) -join ' | ')
            if ($usedBitrate) { $record.Detail = "码率=$usedBitrate 尝试${attempt}次: " + $record.Detail }
            Write-Host ("      失败: {0}" -f $record.Detail) -ForegroundColor Red
        }

        $results.Add($record)
    }
}

# ---------------------------------------------------------------- 汇总 --------
Write-Host "`n================ 生成汇总 ================" -ForegroundColor Cyan
$results | Format-Table -AutoSize -Property File, Resolution, SizeMB, Duration, Status

$okCount = @($results | Where-Object { $_.Status -eq 'OK' }).Count
$warnCount = @($results | Where-Object { $_.Status -eq '解码告警' }).Count
$failCount = @($results | Where-Object { $_.Status -eq '失败' }).Count
$measure = $results | Measure-Object -Property SizeMB -Sum
$totalMB = if ($measure) { [math]::Round($measure.Sum, 2) } else { 0 }

Write-Host ("成功 {0} / 告警 {1} / 失败 {2} / 总体积 {3} MB" -f $okCount, $warnCount, $failCount, $totalMB) -ForegroundColor Cyan

if ($failCount -gt 0) {
    Write-Host "`n失败明细：" -ForegroundColor Red
    $results | Where-Object { $_.Status -eq '失败' } | ForEach-Object {
        Write-Host ("  {0}: {1}" -f $_.File, $_.Detail) -ForegroundColor Red
    }
}
