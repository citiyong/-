<#
  다운로드 / 바탕화면 정리 스크립트

  - 중복 파일(내용이 완전히 같은 파일)을 찾아 하나만 남깁니다.
  - 오래된 동영상을 삭제 후보로 찾습니다.
  - (선택) 남은 파일을 종류별 폴더(문서, 이미지 등)로 정리합니다.

  기본값은 "미리보기"입니다. 아무 파일도 건드리지 않고 바탕화면에 목록(CSV)만 만듭니다.
  목록에서 남기고 싶은 파일의 줄을 지우고 저장한 뒤 -Apply 로 실행하면,
  가장 최근 목록에 남아 있는 파일만 처리합니다. 삭제는 모두 "휴지통"으로 보내므로 되돌릴 수 있습니다.

  사용 예:
    powershell -ExecutionPolicy Bypass -File .\pc-cleanup.ps1                        # 미리보기
    powershell -ExecutionPolicy Bypass -File .\pc-cleanup.ps1 -Organize              # 미리보기 + 종류별 폴더 정리 포함
    powershell -ExecutionPolicy Bypass -File .\pc-cleanup.ps1 -VideoOlderThanDays 30 # 30일 넘은 동영상을 후보로
    powershell -ExecutionPolicy Bypass -File .\pc-cleanup.ps1 -Apply                 # 최근 목록대로 실행
#>
param(
    [switch]$Apply,
    [switch]$Organize,
    [int]$VideoOlderThanDays = 180,
    [int]$VideoMinSizeMB = 0,
    [switch]$SkipDuplicates,
    [switch]$SkipVideos,
    [string[]]$Folders,
    [string]$ReportDir,
    [string]$FromList
)

$ErrorActionPreference = 'Stop'

function Get-DownloadsPath {
    try {
        $p = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
        if ($p) { return $p }
    } catch { }
    return Join-Path $env:USERPROFILE 'Downloads'
}

if (-not $Folders) {
    # 바탕화면이 OneDrive로 옮겨진 경우도 GetFolderPath가 실제 위치를 돌려줍니다.
    $Folders = @((Get-DownloadsPath), [Environment]::GetFolderPath('Desktop'))
}
$Folders = $Folders | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -Unique

$reportRoot = [Environment]::GetFolderPath('Desktop')

$videoExt = @('.mp4', '.mkv', '.avi', '.mov', '.wmv', '.flv', '.webm', '.m4v', '.mpg', '.mpeg', '.3gp', '.ts', '.m2ts', '.vob')
$categories = [ordered]@{
    '문서'     = @('.pdf', '.doc', '.docx', '.xls', '.xlsx', '.csv', '.ppt', '.pptx', '.hwp', '.hwpx', '.txt', '.rtf', '.odt', '.md')
    '이미지'   = @('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.heic', '.svg', '.tif', '.tiff', '.psd', '.ai')
    '동영상'   = $videoExt
    '음악'     = @('.mp3', '.wav', '.flac', '.m4a', '.aac', '.ogg', '.wma')
    '압축파일' = @('.zip', '.rar', '.7z', '.alz', '.egg', '.tar', '.gz', '.iso')
    '설치파일' = @('.exe', '.msi', '.msix', '.appx', '.apk')
}
# 바탕화면 바로가기 등은 정리 대상에서 뺍니다.
$neverTouchExt = @('.lnk', '.url', '.ini')

$canRecycle = $false
try { Add-Type -AssemblyName Microsoft.VisualBasic; $canRecycle = $true } catch { }

function Remove-ToRecycleBin([string]$path) {
    if (-not $canRecycle) { throw '휴지통 기능을 쓸 수 없어 삭제하지 않았습니다.' }
    [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($path, 'OnlyErrorDialogs', 'SendToRecycleBin')
}

function Test-InFolders([string]$path) {
    foreach ($f in $Folders) {
        $root = [IO.Path]::GetFullPath($f).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
        if ([IO.Path]::GetFullPath($path).StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function Invoke-Actions($list, [string]$logDir) {
    $moveLog = New-Object System.Collections.Generic.List[object]
    $done = 0; $failed = 0
    foreach ($a in $list) {
        try {
            if (-not (Test-InFolders $a.파일)) { throw '정리 대상 폴더 밖의 파일이라 건드리지 않았습니다.' }
            if (-not (Test-Path -LiteralPath $a.파일)) { throw '파일이 이미 없습니다.' }
            if ($a.작업 -eq '폴더 이동') {
                if (-not (Test-InFolders (Join-Path $a.비고 'x'))) { throw '이동할 폴더가 정리 대상 폴더 밖입니다.' }
                New-Item -ItemType Directory -Force -Path $a.비고 | Out-Null
                $name = [IO.Path]::GetFileNameWithoutExtension($a.파일)
                $ext = [IO.Path]::GetExtension($a.파일)
                $dest = Join-Path $a.비고 ($name + $ext)
                $n = 1
                while (Test-Path -LiteralPath $dest) { $dest = Join-Path $a.비고 ("$name ($n)$ext"); $n++ }
                Move-Item -LiteralPath $a.파일 -Destination $dest
                $moveLog.Add([pscustomobject]@{ 원래위치 = $a.파일; 새위치 = $dest })
            } elseif ($a.작업 -eq '중복 삭제' -or $a.작업 -eq '동영상 삭제') {
                Remove-ToRecycleBin $a.파일
            } else {
                throw "알 수 없는 작업: $($a.작업)"
            }
            $done++
        } catch {
            $failed++
            Write-Warning "건너뜀: $($a.파일) - $($_.Exception.Message)"
        }
    }
    if ($moveLog.Count) {
        $moveLog | Export-Csv -LiteralPath (Join-Path $logDir '이동_기록.csv') -NoTypeInformation -Encoding UTF8
    }
    Write-Host ''
    Write-Host "완료: ${done}개 처리, ${failed}개 건너뜀. 삭제한 파일은 휴지통에서 되살릴 수 있습니다." -ForegroundColor Green
}

if ($Apply) {
    if (-not $FromList) {
        $latest = Get-ChildItem -LiteralPath $reportRoot -Directory -Filter 'PC정리_보고서_*' -ErrorAction SilentlyContinue |
            Sort-Object Name -Descending | Select-Object -First 1
        if ($latest) { $FromList = Join-Path $latest.FullName '정리_목록.csv' }
    }
    if (-not $FromList -or -not (Test-Path -LiteralPath $FromList)) {
        Write-Host '처리할 목록이 없습니다. 먼저 미리보기를 실행해 목록을 만드세요.' -ForegroundColor Red
        return
    }
    # 엑셀에서 저장하면 UTF-8 표시(BOM)가 빠지고 ANSI(CP949)로 저장될 수 있어 둘 다 받습니다.
    $bytes = [IO.File]::ReadAllBytes($FromList)
    $enc = if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { 'UTF8' } else { 'Default' }
    $list = @(Import-Csv -LiteralPath $FromList -Encoding $enc | Where-Object { $_.파일 })
    if (-not $list.Count) {
        Write-Host "목록을 읽지 못했거나 비어 있습니다: $FromList" -ForegroundColor Red
        return
    }
    Write-Host "목록대로 실행합니다: $FromList ($($list.Count)개)"
    Invoke-Actions $list (Split-Path -Parent $FromList)
    return
}

if (-not $ReportDir) {
    $ReportDir = Join-Path $reportRoot ('PC정리_보고서_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
}
New-Item -ItemType Directory -Force -Path $ReportDir | Out-Null
$reportDirFull = (Resolve-Path -LiteralPath $ReportDir).Path
$allFiles = foreach ($f in $Folders) {
    Get-ChildItem -LiteralPath $f -File -Recurse -Force -ErrorAction SilentlyContinue |
        Where-Object {
            -not ($_.Attributes -band [IO.FileAttributes]::Hidden) -and
            -not ($_.Attributes -band [IO.FileAttributes]::System) -and
            ($neverTouchExt -notcontains $_.Extension.ToLower()) -and
            -not $_.FullName.StartsWith($reportDirFull)
        }
}
$allFiles = @($allFiles | Sort-Object FullName -Unique)

function Format-Size([long]$bytes) {
    if ($bytes -ge 1GB) { return '{0:N2} GB' -f ($bytes / 1GB) }
    if ($bytes -ge 1MB) { return '{0:N1} MB' -f ($bytes / 1MB) }
    return '{0:N0} KB' -f ($bytes / 1KB)
}

$actions = New-Object System.Collections.Generic.List[object]
$removed = @{}

# 1) 중복 파일: 크기가 같은 것끼리 먼저 묶고, 그중에서만 해시를 비교합니다.
if (-not $SkipDuplicates) {
    Write-Host '중복 파일을 찾는 중...'
    $sizeGroups = $allFiles | Where-Object { $_.Length -gt 0 } | Group-Object Length | Where-Object { $_.Count -gt 1 }
    foreach ($sg in $sizeGroups) {
        $hashed = foreach ($file in $sg.Group) {
            try { [pscustomobject]@{ File = $file; Hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash } } catch { }
        }
        foreach ($hg in ($hashed | Group-Object Hash | Where-Object { $_.Count -gt 1 })) {
            # 남길 파일: "(1)", "복사본", "copy" 가 붙지 않은 이름 → 이름이 짧은 것 → 먼저 만든 것
            $ordered = $hg.Group | Sort-Object `
                @{ Expression = { if ($_.File.BaseName -match '\(\d+\)$|복사본|- Copy|copy$') { 1 } else { 0 } } },
                @{ Expression = { $_.File.Name.Length } },
                @{ Expression = { $_.File.CreationTime } }
            $keep = $ordered[0].File
            foreach ($dup in ($ordered | Select-Object -Skip 1)) {
                $actions.Add([pscustomobject]@{
                    작업 = '중복 삭제'; 파일 = $dup.File.FullName; 크기 = Format-Size $dup.File.Length
                    수정일 = $dup.File.LastWriteTime.ToString('yyyy-MM-dd'); 비고 = "남기는 파일: $($keep.FullName)"
                    Bytes = $dup.File.Length
                })
                $removed[$dup.File.FullName] = $true
            }
        }
    }
}

# 2) 동영상: 기준보다 오래되고 기준보다 큰 동영상
if (-not $SkipVideos) {
    Write-Host '동영상을 찾는 중...'
    $cutoff = (Get-Date).AddDays(-$VideoOlderThanDays)
    $allFiles | Where-Object {
        ($videoExt -contains $_.Extension.ToLower()) -and -not $removed[$_.FullName] -and
        $_.LastWriteTime -lt $cutoff -and $_.Length -ge ($VideoMinSizeMB * 1MB)
    } | ForEach-Object {
        $actions.Add([pscustomobject]@{
            작업 = '동영상 삭제'; 파일 = $_.FullName; 크기 = Format-Size $_.Length
            수정일 = $_.LastWriteTime.ToString('yyyy-MM-dd'); 비고 = "${VideoOlderThanDays}일 넘게 수정 안 됨"
            Bytes = $_.Length
        })
        $removed[$_.FullName] = $true
    }
}

# 3) 종류별 폴더 정리: 각 폴더 맨 위에 있는 파일만 옮깁니다(이미 만든 하위 폴더는 그대로).
if ($Organize) {
    foreach ($root in $Folders) {
        Get-ChildItem -LiteralPath $root -File -Force -ErrorAction SilentlyContinue | Where-Object {
            -not $removed[$_.FullName] -and -not ($_.Attributes -band [IO.FileAttributes]::Hidden) -and
            ($neverTouchExt -notcontains $_.Extension.ToLower())
        } | ForEach-Object {
            $ext = $_.Extension.ToLower()
            $cat = '기타'
            foreach ($k in $categories.Keys) { if ($categories[$k] -contains $ext) { $cat = $k; break } }
            $actions.Add([pscustomobject]@{
                작업 = '폴더 이동'; 파일 = $_.FullName; 크기 = Format-Size $_.Length
                수정일 = $_.LastWriteTime.ToString('yyyy-MM-dd'); 비고 = (Join-Path $root $cat)
                Bytes = $_.Length
            })
        }
    }
}

$reportCsv = Join-Path $ReportDir '정리_목록.csv'
$actions | Select-Object 작업, 파일, 크기, 수정일, 비고 | Export-Csv -LiteralPath $reportCsv -NoTypeInformation -Encoding UTF8

$dupBytes = ($actions | Where-Object 작업 -eq '중복 삭제' | Measure-Object Bytes -Sum).Sum
$vidBytes = ($actions | Where-Object 작업 -eq '동영상 삭제' | Measure-Object Bytes -Sum).Sum
Write-Host ''
Write-Host "대상 폴더: $($Folders -join ', ')"
Write-Host ("중복 파일: {0}개 ({1})" -f @($actions | Where-Object 작업 -eq '중복 삭제').Count, (Format-Size $dupBytes))
Write-Host ("오래된 동영상: {0}개 ({1})" -f @($actions | Where-Object 작업 -eq '동영상 삭제').Count, (Format-Size $vidBytes))
if ($Organize) { Write-Host ("종류별 폴더로 옮길 파일: {0}개" -f @($actions | Where-Object 작업 -eq '폴더 이동').Count) }
Write-Host "목록: $reportCsv"

Write-Host ''
Write-Host '미리보기만 했습니다. 아무 파일도 바뀌지 않았습니다.' -ForegroundColor Yellow
Write-Host '목록에서 남기고 싶은 파일의 줄을 지우고 저장한 뒤, 실행(-Apply)하세요.' -ForegroundColor Yellow
try { Invoke-Item -LiteralPath $reportCsv } catch { }
