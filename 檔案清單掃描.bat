<# :
@echo off
setlocal
title File Inventory Scanner
set "SCAN_BAT=%~f0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$p=$env:SCAN_BAT; & ([ScriptBlock]::Create([IO.File]::ReadAllText($p, [Text.Encoding]::UTF8)))"
echo.
pause
exit /b
#>
param(
    [string]$TargetPath,
    [string]$OutputCsv
)
# ============================================================
#  檔案清單掃描（File Inventory）
#  掃描指定資料夾內所有檔案，輸出檔案清單 CSV。
#  只讀取檔案資訊（名稱、路徑、大小、日期），不開啟、不修改任何檔案。
#  輸出的 CSV 可直接匯入「檔案整理工具（DocRoute）」的素材索引。
#
#  用法一：點兩下「檔案清單掃描.bat」，依畫面提示操作。
#  用法二（排程或批次）：
#    powershell -ExecutionPolicy Bypass -File 檔案清單掃描.ps1 -TargetPath "\\伺服器\共用資料夾" -OutputCsv "D:\清單\共用資料夾(全).csv"
# ============================================================

$ErrorActionPreference = 'Stop'
$Interactive = -not ($TargetPath -and $OutputCsv)

function Write-Title([string]$t) {
    Write-Host ''
    Write-Host '========================================' -ForegroundColor DarkCyan
    Write-Host " $t" -ForegroundColor Cyan
    Write-Host '========================================' -ForegroundColor DarkCyan
    Write-Host ''
}

function Select-Folder {
    $typed = Read-Host '請輸入要掃描的資料夾路徑（可貼上 \\伺服器\共用資料夾 這類網路路徑；直接按 Enter 則開啟選擇視窗）'
    $typed = $typed.Trim().Trim('"')
    if ($typed) { return $typed }
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = '請選擇要掃描的資料夾'
        $dlg.ShowNewFolderButton = $false
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.SelectedPath }
    } catch { }
    return $null
}

Write-Title '檔案清單掃描'

# ---------- 1. 掃描目標 ----------
if (-not $TargetPath) { $TargetPath = Select-Folder }
if (-not $TargetPath) { Write-Host '未選擇資料夾，結束。' -ForegroundColor Yellow; exit 1 }
$TargetPath = $TargetPath.Trim().Trim('"')
if (-not (Test-Path -LiteralPath $TargetPath -PathType Container)) {
    Write-Host "找不到資料夾，或沒有讀取權限：$TargetPath" -ForegroundColor Red
    exit 1
}

# ---------- 2. 輸出位置 ----------
if (-not $OutputCsv) {
    $leaf = Split-Path $TargetPath.TrimEnd('\') -Leaf
    if (-not $leaf -or $leaf -match '^[A-Za-z]:$') { $leaf = '掃描結果' }
    $leaf = $leaf -replace '[\\/:*?"<>|]', '_'
    $desk = [Environment]::GetFolderPath('Desktop'); if (-not $desk) { $desk = (Get-Location).Path }
    $default = Join-Path $desk ("{0}(全)_{1}.csv" -f $leaf, (Get-Date -Format 'yyyyMMdd'))
    Write-Host ''
    Write-Host "CSV 預設存到：$default"
    $typed = Read-Host '直接按 Enter 使用預設；或輸入其他完整檔名（.csv）'
    $typed = $typed.Trim().Trim('"')
    $OutputCsv = if ($typed) { $typed } else { $default }
    if ($OutputCsv -notmatch '\.csv$') { $OutputCsv += '.csv' }
}

Write-Host ''
Write-Host "掃描資料夾：$TargetPath"
Write-Host "輸出 CSV ：$OutputCsv"
Write-Host ''

# ---------- 3. 掃描 ----------
$ScanStart = Get-Date
$NormalizedRoot = $TargetPath.TrimEnd('\')
$Inventory = New-Object System.Collections.Generic.List[object]
$ScanErrors = @()
$count = 0
$tick = [Diagnostics.Stopwatch]::StartNew()

Write-Host '掃描中，檔案多時需要數分鐘，請勿關閉視窗…'
Get-ChildItem -LiteralPath $TargetPath -File -Recurse -Force -ErrorAction SilentlyContinue -ErrorVariable +ScanErrors |
    ForEach-Object {
        $File = $_
        $count++
        if ($tick.ElapsedMilliseconds -ge 1000) {
            Write-Progress -Activity '掃描檔案中' -Status ("已找到 {0:N0} 個檔案　{1}" -f $count, $File.DirectoryName)
            $tick.Restart()
        }
        if ($File.FullName.StartsWith($NormalizedRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
            $RelativePath = $File.FullName.Substring($NormalizedRoot.Length).TrimStart('\')
        } else {
            $RelativePath = $File.FullName
        }
        $Inventory.Add([PSCustomObject]@{
            FileName       = $File.Name
            BaseName       = $File.BaseName
            Extension      = $File.Extension.ToLower()
            FullPath       = $File.FullName
            Directory      = $File.DirectoryName
            RelativePath   = $RelativePath
            SizeBytes      = $File.Length
            SizeKB         = [math]::Round($File.Length / 1KB, 2)
            SizeMB         = [math]::Round($File.Length / 1MB, 2)
            CreationTime   = $File.CreationTime
            LastWriteTime  = $File.LastWriteTime
            LastAccessTime = $File.LastAccessTime
            Attributes     = $File.Attributes.ToString()
            IsReadOnly     = $File.IsReadOnly
            ScanTime       = $ScanStart
        })
    }
Write-Progress -Activity '掃描檔案中' -Completed
Write-Host ("掃描完成，共 {0:N0} 個檔案。" -f $Inventory.Count)

# ---------- 4. 寫出 CSV（UTF-8 含 BOM，Excel 可直接開啟） ----------
Write-Host '寫入 CSV 中…'
$OutputDirectory = Split-Path -Parent $OutputCsv
if ($OutputDirectory -and -not (Test-Path -LiteralPath $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}
$enc = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
try {
    $Inventory | Sort-Object FullPath | Export-Csv -LiteralPath $OutputCsv -NoTypeInformation -Encoding $enc
} catch {
    Write-Host "無法寫入 CSV（檔案可能正在 Excel 中開啟，或沒有寫入權限）：$($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# ---------- 5. 結果 ----------
$TotalSize = ($Inventory | Measure-Object -Property SizeBytes -Sum).Sum
if ($null -eq $TotalSize) { $TotalSize = 0 }
$Duration = (Get-Date) - $ScanStart

Write-Title '掃描完成'
Write-Host ("檔案數　：{0:N0}" -f $Inventory.Count)
Write-Host ("總大小　：{0:N0} MB" -f ($TotalSize / 1MB))
Write-Host ("耗時　　：{0:N0} 秒" -f $Duration.TotalSeconds)
Write-Host "CSV 位置：$OutputCsv"
if ($ScanErrors.Count) {
    $log = [IO.Path]::ChangeExtension($OutputCsv, $null).TrimEnd('.') + '_無法讀取清單.txt'
    $ScanErrors | ForEach-Object { if ($_.TargetObject) { [string]$_.TargetObject } else { $_.Exception.Message } } |
        Sort-Object -Unique | Set-Content -LiteralPath $log -Encoding UTF8
    Write-Host ("有 {0:N0} 個資料夾或檔案無法讀取（沒有權限或路徑過長），已略過，清單在：{1}" -f $ScanErrors.Count, $log) -ForegroundColor Yellow
}
Write-Host ''

if ($Interactive) {
    try { Start-Process explorer.exe -ArgumentList "/select,`"$OutputCsv`"" } catch { }
}
exit 0
