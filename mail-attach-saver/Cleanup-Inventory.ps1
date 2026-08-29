#requires -Version 5.1
<#
.SYNOPSIS
    保存先ルート配下のフォルダを棚卸しし、コード形式フォルダと旧形式疑いフォルダに分類する。

.DESCRIPTION
    ★現地でのみ実行すること。
    ★まずは -Delete を付けない「引数なし（-Rootのみ）」の一覧表示オプションで実行し、
      出力されたCSVをクライアントと一緒に確認してから、必要な場合のみ -Delete を検討すること。

    -Root 配下の直下サブフォルダを列挙し、CreationTime が -From ～ -To の範囲内のものを対象に、
      - コード形式（既定では同フォルダの config.local.ps1 の $SubjectPatterns から自動導出。
        導出できない場合のみ ^(DJ|BJ|CH)\d{5}_$ ＝ Save-MailAttachments.ps1 の既定命名規則にフォールバック）
      - その他（旧形式の疑い。件名がそのままフォルダ名になっている等）
    の2種類に分類し、コンソールのテーブルと logs\cleanup-inventory-<日時>.csv に出力する。

    既定では読み取り専用（一覧表示のみ）。何も削除・変更しない。

.PARAMETER Root
    棚卸し対象の保存先ルートフォルダ（必須）。

.PARAMETER From
    対象とするフォルダのCreationTimeの下限。省略時は制限なし相当（2000-01-01）。

.PARAMETER To
    対象とするフォルダのCreationTimeの上限。省略時は現在時刻。

.PARAMETER CodePatterns
    「コード形式」とみなす正規表現の配列。
    [レビュー反映] 省略時は同フォルダの config.local.ps1（無ければ config.sample.ps1）の
    $SubjectPatterns から自動導出するため、config.local.ps1 に件名パターンを追加しても
    このパラメータを手動で追従させる必要はない。config が読めない場合のみ
    Save-MailAttachments.ps1 の既定命名規則（DJ|BJ|CH + 数字5桁 + "_"）にフォールバックする。
    明示的に指定した場合はその値が優先される。

.PARAMETER Delete
    「その他」分類のみを削除候補とする。-Confirm との併用が必須（二重ゲート）。
    実行時にはさらに削除件数の入力による確認を求める。既定では何も削除しない。

.PARAMETER Confirm
    -Delete と併用して初めて削除が有効になる、明示的な確認スイッチ。
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Root,

    [Parameter(Mandatory = $false)]
    [datetime]$From = [datetime]"2000-01-01",

    [Parameter(Mandatory = $false)]
    [datetime]$To = (Get-Date),

    [Parameter(Mandatory = $false)]
    [string[]]$CodePatterns,

    [switch]$Delete,
    [switch]$Confirm
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $Root)) {
    Write-Host ("保存先ルートが見つかりません: {0}" -f $Root)
    return
}

$scriptLogFolder = Join-Path $PSScriptRoot "logs"
if (-not (Test-Path -LiteralPath $scriptLogFolder)) {
    New-Item -ItemType Directory -Path $scriptLogFolder -Force | Out-Null
}

# ------------------------------------------------------------
# [レビュー反映・修正2] 分類パターンをconfigの $SubjectPatterns と連動させる
#   -CodePatterns が明示指定されなかった場合、同フォルダの config.local.ps1
#   （無ければ config.sample.ps1）を読み込み、$SubjectPatterns の各パターンを
#   "^(?:" + パターン + ")_$" に変換して既定値にする
#   （(?!\d) 等の先読みは末尾の "_$" の前でもそのまま機能するため変換は不要）。
#   config が読めない・$SubjectPatternsが無い場合のみ、従来のハードコード既定値に
#   フォールバックし警告を表示する。
# ------------------------------------------------------------
function Get-DefaultCodePatterns {
    param([string]$ToolRoot)

    $fallback = @('^(DJ|BJ|CH)\d{5}_$')

    $localConfig = Join-Path $ToolRoot "config.local.ps1"
    $sampleConfig = Join-Path $ToolRoot "config.sample.ps1"

    $configPath = $null
    if (Test-Path -LiteralPath $localConfig) { $configPath = $localConfig }
    elseif (Test-Path -LiteralPath $sampleConfig) { $configPath = $sampleConfig }

    if ($null -eq $configPath) {
        Write-Host "警告: config.local.ps1 / config.sample.ps1 のどちらも見つからないため、既定の分類パターンを使用します。"
        return $fallback
    }

    try {
        # $SubjectPatterns 以外の変数（$SaveRoot等）でこのスクリプトのスコープを汚さないよう、
        # 子スコープ (& { ... }) の中だけで config を dot-source する。
        $subjectPatterns = & {
            . $configPath
            return $SubjectPatterns
        }

        if ($null -eq $subjectPatterns -or @($subjectPatterns).Count -eq 0) {
            Write-Host ("警告: {0} から `$SubjectPatterns を取得できなかったため、既定の分類パターンを使用します。" -f (Split-Path -Leaf $configPath))
            return $fallback
        }

        return @($subjectPatterns | ForEach-Object { "^(?:" + $_ + ")_$" })
    }
    catch {
        Write-Host ("警告: {0} の読み込みに失敗したため、既定の分類パターンを使用します（{1}）。" -f (Split-Path -Leaf $configPath), $_.Exception.Message)
        return $fallback
    }
}

if (-not $PSBoundParameters.ContainsKey('CodePatterns') -or $null -eq $CodePatterns -or @($CodePatterns).Count -eq 0) {
    $CodePatterns = Get-DefaultCodePatterns -ToolRoot $PSScriptRoot
}

Write-Host "この分類パターンで判定します（コード形式とみなす正規表現）:"
foreach ($p in $CodePatterns) { Write-Host ("  - {0}" -f $p) }

# ------------------------------------------------------------
# 棚卸し
# ------------------------------------------------------------
$subDirs = Get-ChildItem -LiteralPath $Root -Directory -ErrorAction Stop

$results = New-Object System.Collections.Generic.List[object]

foreach ($dir in $subDirs) {
    if ($dir.CreationTime -lt $From -or $dir.CreationTime -gt $To) { continue }

    $isCodeFormat = $false
    foreach ($pattern in $CodePatterns) {
        if ($dir.Name -match $pattern) {
            $isCodeFormat = $true
            break
        }
    }
    $category = if ($isCodeFormat) { "コード形式" } else { "その他(旧形式疑い)" }

    $fileCount = 0
    $totalSizeBytes = 0
    try {
        $files = Get-ChildItem -LiteralPath $dir.FullName -File -Recurse -ErrorAction SilentlyContinue
        if ($files) {
            $fileCount = @($files).Count
            $sum = ($files | Measure-Object -Property Length -Sum).Sum
            if ($null -ne $sum) { $totalSizeBytes = $sum }
        }
    }
    catch {
        # サイズ集計に失敗しても棚卸し自体は継続する
    }

    $results.Add([PSCustomObject]@{
        フォルダ名   = $dir.Name
        分類         = $category
        作成日時     = $dir.CreationTime
        ファイル数   = $fileCount
        合計サイズKB = [math]::Round($totalSizeBytes / 1KB, 1)
        フルパス     = $dir.FullName
    })
}

$codeCount = @($results | Where-Object { $_.分類 -eq "コード形式" }).Count
$otherCount = @($results | Where-Object { $_.分類 -eq "その他(旧形式疑い)" }).Count

Write-Host ("===== 棚卸し結果: {0} 〜 {1}（{2}件中: コード形式 {3} / その他 {4}） =====" -f `
    $From.ToString("yyyy-MM-dd"), $To.ToString("yyyy-MM-dd"), $results.Count, $codeCount, $otherCount)

$results | Sort-Object 分類, フォルダ名 | Format-Table フォルダ名, 分類, 作成日時, ファイル数, 合計サイズKB -AutoSize | Out-Host

# ------------------------------------------------------------
# CSV出力
# ------------------------------------------------------------
$csvPath = Join-Path $scriptLogFolder ("cleanup-inventory-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".csv")
$results | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
Write-Host ("CSV出力: {0}" -f $csvPath)

# ------------------------------------------------------------
# 削除（既定では何もしない。-Delete と -Confirm の両方が必須の二重ゲート）
# ------------------------------------------------------------
if (-not $Delete) {
    Write-Host ""
    Write-Host "（読み取り専用モードで終了。削除は行っていません。削除を検討する場合は -Delete -Confirm を指定してください）"
    return
}

if (-not $Confirm) {
    Write-Host ""
    Write-Host "-Delete のみでは削除を実行しません。-Delete -Confirm の両方を指定した場合のみ削除処理に進みます（二重ゲート）。"
    return
}

$targets = @($results | Where-Object { $_.分類 -eq "その他(旧形式疑い)" })

if ($targets.Count -eq 0) {
    Write-Host ""
    Write-Host "削除対象（その他分類）はありませんでした。"
    return
}

Write-Host ""
Write-Host "----- 削除対象一覧（その他(旧形式疑い) のみ。コード形式フォルダは対象外） -----"
$targets | Format-Table フォルダ名, 作成日時, ファイル数, フルパス -AutoSize | Out-Host
Write-Host ("削除対象: {0} 件" -f $targets.Count)

$typed = Read-Host ("本当に上記 {0} 件を削除する場合は、その件数（半角数字）を入力してください" -f $targets.Count)
if ($typed -ne $targets.Count.ToString()) {
    Write-Host "入力された件数が一致しなかったため、削除を中止しました（何も削除していません）。"
    return
}

$deletedCount = 0
foreach ($t in $targets) {
    try {
        Remove-Item -LiteralPath $t.フルパス -Recurse -Force
        Write-Host ("削除しました: {0}" -f $t.フルパス)
        $deletedCount++
    }
    catch {
        Write-Host ("削除に失敗しました: {0} ({1})" -f $t.フルパス, $_.Exception.Message)
    }
}
Write-Host ("完了: {0} / {1} 件を削除しました。" -f $deletedCount, $targets.Count)
