$ErrorActionPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$ffmpeg = 'c:\Users\al_ta\CodeBuddy\video_player\tools\ffmpeg.exe'
$dir    = 'c:\Users\al_ta\CodeBuddy\video_player\videos'
$tmpDir = Join-Path $env:TEMP 'vp_partial_test'
$log    = Join-Path $tmpDir 'log.txt'
New-Item -ItemType Directory -Force -Path $tmpDir | Out-Null

$rows = @()
Get-ChildItem -Path $dir -File | Where-Object { '.ps1','.csv','.md','.txt','.json' -notcontains $_.Extension.ToLower() } | Sort-Object Name | ForEach-Object {
    $take = [int]($_.Length * 0.15)
    $partial = Join-Path $tmpDir $_.Name
    $fs = [System.IO.File]::OpenRead($_.FullName)
    $buf = New-Object byte[] $take
    $null = $fs.Read($buf, 0, $take)
    $fs.Close()
    [System.IO.File]::WriteAllBytes($partial, $buf)

    if (Test-Path $log) { Remove-Item $log -Force }
    Start-Process -FilePath $ffmpeg -ArgumentList @('-hide_banner', '-nostdin', '-i', $partial, '-frames:v', '2', '-f', 'null', '-') `
                  -RedirectStandardError $log -NoNewWindow -Wait | Out-Null
    $t = ''
    if (Test-Path $log) { $t = [System.IO.File]::ReadAllText($log) }

    $decoded = $t -match 'frame=\s*[1-9]'
    $err = ''
    if ($t -match 'moov atom not found')                 { $err = 'moov atom not found (元数据在文件尾部)' }
    elseif ($t -match 'Invalid data found when processing input') { $err = 'Invalid data found when processing input' }
    elseif ($t -match 'Error opening input')             { $err = '无法打开不完整文件' }
    elseif ($t -match 'error while decoding')            { $err = '解码错误' }

    $rows += [pscustomobject]@{
        File           = $_.Name
        PartialBytesKB = [math]::Round($take / 1KB, 0)
        StartFrom15pct = $(if ($decoded) { 'YES 可起播' } else { 'NO 无法起播' })
        Err            = $err
    }
}

$rows | Format-Table -AutoSize -Wrap
Remove-Item $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
