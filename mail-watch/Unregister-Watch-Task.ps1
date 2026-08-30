#requires -Version 5.1
<#
.SYNOPSIS
    タスクスケジューラから MailWatchClient タスクを削除する。

.DESCRIPTION
    schtasks.exe /Delete を使い、Register-Watch-Task.ps1 で登録したタスクを削除する。
    実行前に確認プロンプトを表示する。

.NOTES
    ★このスクリプトは開発機では実行しないこと。
#>

[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$taskName = "MailWatchClient"

$queryOutput = & schtasks.exe /Query /TN $taskName 2>&1
$exists = ($LASTEXITCODE -eq 0)

if (-not $exists) {
    Write-Host ("タスク '{0}' は登録されていません。削除の必要はありません。" -f $taskName)
    return
}

Write-Host ("タスク '{0}' を削除します。" -f $taskName)
$answer = Read-Host "本当に削除しますか？ (y/N)"
if ($answer -ne "y" -and $answer -ne "Y") {
    Write-Host "削除を中止しました。"
    return
}

& schtasks.exe /Delete /TN $taskName /F

if ($LASTEXITCODE -eq 0) {
    Write-Host ("タスク '{0}' を削除しました。" -f $taskName)
}
else {
    Write-Host ("削除に失敗しました（終了コード {0}）。" -f $LASTEXITCODE)
}
