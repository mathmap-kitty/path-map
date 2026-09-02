<#
.SYNOPSIS
    註冊（或移除）「path-map 體育班單招每週檢核」的 Windows 工作排程。

.DESCRIPTION
    重灌或換電腦之後，clone 完 repo 跑這支就能把每週自動檢核裝回來。
    工作排程本身不在 repo 裡，只有這支註冊腳本在。

    用 pythonw.exe 而不是 .bat：Task Scheduler 執行 .bat 時，
    cmd 的主控台一關就會回 0xC000013A（Ctrl-C 結束），任務等於沒跑成。
    pythonw 沒有主控台，不會有這個問題；fetch_sport.py 本身會把輸出
    導進 _build\sport_watch_stdout.log，執行摘要寫進 sport_watch_log.md。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File _build\register_sport_watch.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File _build\register_sport_watch.ps1 -Unregister
#>
param(
    [switch]$Unregister,
    [string]$DayOfWeek = 'Sunday',
    [string]$At = '09:00'
)

$TaskName = 'path-map 體育班單招每週檢核'
$BuildDir = Split-Path -Parent $MyInvocation.MyCommand.Path

if ($Unregister) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "已移除工作排程：$TaskName"
    exit 0
}

# 找 pythonw.exe：優先用 PATH 上的 python 旁邊那支，找不到再退回常見安裝路徑
$Pythonw = $null
$py = Get-Command python.exe -ErrorAction SilentlyContinue
if ($py) {
    $candidate = Join-Path (Split-Path -Parent $py.Source) 'pythonw.exe'
    if (Test-Path $candidate) { $Pythonw = $candidate }
}
if (-not $Pythonw) {
    $fallback = Get-Command pythonw.exe -ErrorAction SilentlyContinue
    if ($fallback) { $Pythonw = $fallback.Source }
}
if (-not $Pythonw) {
    Write-Error "找不到 pythonw.exe。請先安裝 Python 並確認它在 PATH 上，或自行修改本腳本的路徑。"
    exit 1
}

$script = Join-Path $BuildDir 'fetch_sport.py'
if (-not (Test-Path $script)) {
    Write-Error "找不到 $script"
    exit 1
}

$action   = New-ScheduledTaskAction -Execute $Pythonw -Argument 'fetch_sport.py' -WorkingDirectory $BuildDir
$trigger  = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $DayOfWeek -At $At
# 三個設定都是實測踩出來的：
#   -AllowStartIfOnBatteries / -DontStopIfGoingOnBatteries
#       New-ScheduledTaskSettingsSet 預設「使用電池時不啟動」，筆電沒插電時工作會
#       一直卡在 Queued 不執行，而且 LastResult 還是 0，很難發現。
#   不加 -RunOnlyIfNetworkAvailable
#       同樣會讓工作卡在 Queued。抓不到網頁時 fetch_sport.py 本來就會記 log 並
#       保留舊資料，讓它跑下去再失敗比較好。
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable `
                                         -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
                                         -ExecutionTimeLimit (New-TimeSpan -Minutes 20) -Hidden

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings `
    -Description "每週$DayOfWeek $At 檢查大專校院運動績優單獨招生資訊有無更新，有變動就重建 sport.html，摘要寫入 _build\sport_watch_log.md" `
    -Force | Out-Null

Write-Host "已註冊工作排程：$TaskName"
Write-Host "  執行檔：$Pythonw fetch_sport.py"
Write-Host "  工作目錄：$BuildDir"
Get-ScheduledTaskInfo -TaskName $TaskName | Select-Object NextRunTime | Format-List

Write-Host "先手動跑一次確認（約 30 秒）："
Write-Host "  Start-ScheduledTask -TaskName '$TaskName'"
Write-Host "跑完看 _build\sport_watch_log.md 最下面那一段；LastTaskResult 應為 0："
Write-Host "  Get-ScheduledTaskInfo -TaskName '$TaskName' | Select-Object LastRunTime,LastTaskResult"
