$ErrorActionPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$ffmpeg = 'c:\Users\al_ta\CodeBuddy\video_player\tools\ffmpeg.exe'
$dir    = 'c:\Users\al_ta\CodeBuddy\video_player\videos'

function Get-BoxLayout([string]$path) {
    $res = New-Object System.Collections.Specialized.OrderedDictionary
    $fs = [System.IO.File]::OpenRead($path)
    try {
        $br = New-Object System.IO.BinaryReader($fs)
        while ($fs.Position -lt $fs.Length) {
            $start = $fs.Position
            if (($fs.Length - $start) -lt 8) { break }
            $b = $br.ReadBytes(8)
            if ($b.Length -lt 8) { break }
            $size = [double]0
            for ($i = 0; $i -lt 4; $i++) { $size = $size * 256 + [double]$b[$i] }
            $type = [System.Text.Encoding]::ASCII.GetString($b, 4, 4)
            $hdr = 8
            if ($size -eq 1) {
                $b2 = $br.ReadBytes(8)
                $size = [double]0
                for ($i = 0; $i -lt 8; $i++) { $size = $size * 256 + [double]$b2[$i] }
                $hdr = 16
            } elseif ($size -eq 0) {
                $size = [double]($fs.Length - $start)
            }
            if (-not $res.Contains($type)) { $res.Add($type, [long]$start) }
            if ($size -lt $hdr) { break }
            $fs.Position = [long]($start + $size)
            if ($res.Count -gt 24) { break }
            if ($res.Contains('mdat') -and ($res.Contains('moov') -or $res.Contains('moof'))) { break }
        }
    } finally { $fs.Close() }
    return $res
}

$tmpLog = Join-Path $env:TEMP ("fflog_{0}.txt" -f [guid]::NewGuid().ToString('N'))

function Get-FfprobeInfo([string]$path) {
    if (Test-Path $tmpLog) { Remove-Item $tmpLog -Force }
    Start-Process -FilePath $ffmpeg -ArgumentList @('-hide_banner', '-i', $path) `
                  -RedirectStandardError $tmpLog -RedirectStandardOutput "$tmpLog.out" `
                  -NoNewWindow -Wait | Out-Null
    $out = ''
    if (Test-Path $tmpLog) { $out = [System.IO.File]::ReadAllText($tmpLog) }
    $fmt = '?'; $dur = '?'; $vcodec = '?'; $ares = '?'
    if ($out -match 'Input #0, ([^,]+),')      { $fmt = $Matches[1].Trim() }
    if ($out -match 'Duration: ([^,]+)')       { $dur = $Matches[1].Trim() }
    foreach ($line in ($out -split "`n")) {
        if ($line -match 'Stream #.*Video: ') {
            $tail = ($line -split 'Video: ', 2)[1]
            $vcodec = ($tail -split '[, ]+')[0]
            if ($tail -match '(\d{2,5})x(\d{2,5})') { $ares = $Matches[1] + 'x' + $Matches[2] }
            break
        }
    }
    return @{ Format = $fmt; Duration = $dur; VCodec = $vcodec; Res = $ares }
}

function Test-Progressive {
    param([string]$ext, $boxes)
    switch -Regex ($ext) {
        '\.(mp4|mov|m4v|3gp)$' {
            if ($boxes.Contains('moof')) { return @($true,  '分片 MP4 (fMP4)，天生支持流式') }
            if ($boxes.Contains('moov') -and $boxes.Contains('mdat')) {
                if ([long]$boxes['moov'] -lt [long]$boxes['mdat']) {
                    return @($true, ('moov 在 mdat 之前 (offset {0} < {1})，已 faststart' -f $boxes['moov'], $boxes['mdat']))
                } else {
                    return @($false, ('moov 在文件尾部 (offset {0} > mdat {1})，必须下载完才能起播' -f $boxes['moov'], $boxes['mdat']))
                }
            }
            if ($boxes.Contains('moov')) { return @($true, 'moov 位于文件头部') }
            return @($false, '未找到 moov box')
        }
        '\.(mkv|webm)$' { return @($true, 'Matroska/WebM 为流式容器，顺序边下边播') }
        '\.(flv)$'      { return @($true, 'FLV 元数据+数据在头部，支持渐进下载播放') }
        '\.(ts)$'       { return @($true, 'MPEG-TS 本身就是流媒体封装') }
        '\.(ogv|ogg)$'  { return @($true, 'Ogg 页结构为流式封装') }
        '\.(wmv|asf)$'  { return @($true, 'ASF Header Object 在头部，支持渐进/流式') }
        '\.(avi)$'      { return @($false, 'AVI 索引 idx1 在文件尾部，顺序可播但拖动需全文件') }
        '\.(mpg|mpeg)$' { return @($false, 'MPEG-PS 无全局索引，顺序可播但拖动需全文件') }
        default         { return @($false, '未知容器') }
    }
}

$rows = @()
Get-ChildItem -Path $dir -File | Where-Object { '.ps1','.csv','.md','.txt','.json' -notcontains $_.Extension.ToLower() } | Sort-Object Name | ForEach-Object {
    $ext = $_.Extension.ToLower()
    $boxes = $null
    if ($ext -match '\.(mp4|mov|m4v|3gp)$') { $boxes = Get-BoxLayout $_.FullName }
    $info = Get-FfprobeInfo $_.FullName
    $v = Test-Progressive -ext $ext -boxes $boxes
    $rows += [pscustomobject]@{
        File     = $_.Name
        SizeMB   = [math]::Round($_.Length / 1MB, 2)
        Container= $info.Format
        Duration = $info.Duration
        Res      = $info.Res
        VCodec   = $info.VCodec
        Support  = $(if ($v[0]) { 'YES' } else { 'NO' })
        Reason   = $v[1]
        MoovPos  = $(if ($boxes -and $boxes.Contains('moov')) { $boxes['moov'] } else { '' })
        MdatPos  = $(if ($boxes -and $boxes.Contains('mdat')) { $boxes['mdat'] } else { '' })
    }
}

$rows | Format-Table -AutoSize -Wrap
"==== 汇总 ===="
"支持边下边播: " + (($rows | Where-Object Support -eq 'YES').Count) + " / " + $rows.Count
"不支持:       " + (($rows | Where-Object Support -eq 'NO').Count)
$rows | Export-Csv -Path 'c:\Users\al_ta\CodeBuddy\video_player\tools\streaming_check.csv' -NoTypeInformation -Encoding UTF8
if (Test-Path $tmpLog) { Remove-Item $tmpLog -Force }
