#requires -Version 5.1
<#
.SYNOPSIS
    Watch-ClientMail.ps1 をタスクスケジューラへ登録する（5分間隔・現在ユーザー・ログオン時のみ）。

.DESCRIPTION
    姉妹ツール mail-attach-saver の Register-Task.ps1 と同方式。
    schtasks.exe /Create を使い、Watch-ClientMail.ps1 を5分ごとに実行するタスクを登録する。
    /RU（実行ユーザー）を指定しないことで、schtasks.exe の既定動作である
    「このコマンドを実行した現在のユーザーとして、ログオン中のみ対話トークンで実行する」
    設定になる（パスワード保存が不要で、資格情報のハードコードを避けられる）。
    ウィンドウは -WindowStyle Hidden で非表示にする。

    実行前に登録内容を表示し、y/N の確認を求める。

.NOTES
    ★このスクリプトは開発機では実行しないこと（本人のPCで、動作確認が済んでから実行する）。
    ★常駐化する前に、README.md の「セキュリティ確認ダイアログ」の注意を必ず読むこと。
      ダイアログが出る状態のままタスク化すると、5分毎に無人で処理が止まり続ける。
    タスク名: MailWatchClient
#>

[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$taskName = "MailWatchClient"
$scriptPath = Join-Path $PSScriptRoot "Watch-ClientMail.ps1"

if (-not (Test-Path -LiteralPath $scriptPath)) {
    throw "Watch-ClientMail.ps1 が見つかりません: $scriptPath"
}

# ------------------------------------------------------------
# ネイティブコマンド実行ヘルパー
#   PowerShellの配列引数渡しは、値に " が含まれる文字列を渡すと
#   引用符が失われて空白でトークン分割されてしまう既知の問題があるため
#   （mail-attach-saver側で実測確認済み）、ProcessStartInfo.Arguments に
#   生のコマンドライン文字列を直接渡す方式を使う。/TR の値に含まれる " は
#   \" にエスケープしてから外側を " で囲む（標準的なコマンドライン引数のエスケープ規則）。
# ------------------------------------------------------------
function Invoke-NativeCommand {
    param(
        [string]$FileName,
        [string]$Arguments
    )
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $FileName
    $psi.Arguments = $Arguments
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true

    $proc = [System.Diagnostics.Process]::Start($psi)
    $stdout = $proc.StandardOutput.ReadToEnd()
    $stderr = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()

    return [PSCustomObject]@{
        ExitCode = $proc.ExitCode
        StdOut   = $stdout
        StdErr   = $stderr
    }
}

# 実行させたいコマンド（人間が読む用・エスケープ前）
$quote = '"'
$taskRun = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File " + $quote + $scriptPath + $quote

Write-Host "===== 以下の内容でタスクスケジューラへ登録します ====="
Write-Host ("タスク名     : {0}" -f $taskName)
Write-Host ("実行コマンド : {0}" -f $taskRun)
Write-Host "実行間隔     : 5分ごと"
Write-Host ("実行ユーザー : 現在ログオン中のユーザー（{0}）※パスワードは保存しない" -f $env:USERNAME)
Write-Host "実行条件     : ログオン時のみ・ウィンドウ非表示・実行レベル=標準（管理者権限なし）"
Write-Host "備考         : Outlookが起動していない時間帯は、本体スクリプトが自分で判断して何もせず終了する"
Write-Host "確認         : セキュリティ確認ダイアログが出ない状態（README.md参照）で手動実行を確認済みか？"
Write-Host "=================================================="

$answer = Read-Host "この内容で登録しますか？ (y/N)"
if ($answer -ne "y" -and $answer -ne "Y") {
    Write-Host "登録を中止しました。"
    return
}

# 既存の同名タスクを確認（あれば知らせたうえで /F により上書きする）
$queryResult = Invoke-NativeCommand -FileName "schtasks.exe" -Arguments ("/Query /TN " + $quote + $taskName + $quote)
if ($queryResult.ExitCode -eq 0) {
    Write-Host ("既存タスク '{0}' が見つかりました。設定を上書きします。" -f $taskName)
}

# /TR の値だけ、内側の " を \" にエスケープしてから外側を " で囲む
$taskRunEscaped = $taskRun.Replace('"', '\"')

$createArgs = "/Create /TN {0}{1}{0} /TR {0}{2}{0} /SC MINUTE /MO 5 /RL LIMITED /F" -f $quote, $taskName, $taskRunEscaped

$createResult = Invoke-NativeCommand -FileName "schtasks.exe" -Arguments $createArgs

if ($createResult.ExitCode -eq 0) {
    Write-Host ("タスク '{0}' の登録に成功しました。" -f $taskName)
    if ($createResult.StdOut) { Write-Host $createResult.StdOut }
    Write-Host "確認方法: タスクスケジューラを開き「MailWatchClient」を確認、または次を実行:"
    Write-Host ('  schtasks.exe /Query /TN "MailWatchClient" /V /FO LIST')
}
else {
    Write-Host ("登録に失敗しました（終了コード {0}）。" -f $createResult.ExitCode)
    if ($createResult.StdErr) { Write-Host $createResult.StdErr }
    if ($createResult.StdOut) { Write-Host $createResult.StdOut }
}
