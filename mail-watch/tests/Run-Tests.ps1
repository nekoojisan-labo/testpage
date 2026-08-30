#requires -Version 5.1
<#
    tests\Run-Tests.ps1  (mail-watch)

    MailAttachLib.ps1 の純粋関数に対する合成テスト。
    Outlook不要・ネットワーク不要。ファイル操作は $env:TEMP 配下に作成する
    専用一時フォルダのみで行い、テスト終了時に削除する。

    実行方法:
        powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests\Run-Tests.ps1

    終了コード:
        0 = 全PASS
        1 = 1件以上FAIL（詳細を標準出力に列挙）
#>

$ErrorActionPreference = "Stop"

$libPath = Join-Path $PSScriptRoot "..\MailAttachLib.ps1"
. $libPath

# ------------------------------------------------------------
# 簡易テストランナー（mail-attach-saver側と同一パターン）
# ------------------------------------------------------------
$script:Results = New-Object System.Collections.Generic.List[object]

function Add-Result {
    param(
        [string]$Name,
        [bool]$Pass,
        [string]$Detail = ""
    )
    $script:Results.Add([PSCustomObject]@{
        Name   = $Name
        Pass   = $Pass
        Detail = $Detail
    })
}

function Assert-Equal {
    param(
        [string]$Name,
        $Expected,
        $Actual
    )
    $isEqual = $false
    if ($null -eq $Expected -and $null -eq $Actual) {
        $isEqual = $true
    }
    elseif ($null -eq $Expected -or $null -eq $Actual) {
        $isEqual = $false
    }
    else {
        $isEqual = ($Expected.ToString() -ceq $Actual.ToString())
    }

    $detail = "期待値=[{0}] 実際=[{1}]" -f $Expected, $Actual
    Add-Result -Name $Name -Pass $isEqual -Detail $detail
}

function Assert-True {
    param(
        [string]$Name,
        [bool]$Condition,
        [string]$Detail = ""
    )
    Add-Result -Name $Name -Pass $Condition -Detail $Detail
}

# ------------------------------------------------------------
# テスト用一時フォルダ（$env:TEMP 配下のみ使用）
# ------------------------------------------------------------
$testRoot = Join-Path $env:TEMP ("mailwatch-tests-" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $testRoot -Force | Out-Null

try {

    # ============================================================
    # 1. Test-SenderMatch（差出人ドメイン一致判定）
    #    ※ 実際のドメイン例はテスト用の架空値（example.co.jp はRFC2606等で
    #      予約された例示用ドメイン）を使用し、実在の会社名・ドメインは書かない。
    # ============================================================
    $domains = @("example.co.jp")

    # 完全一致ドメイン
    $r = Test-SenderMatch -SenderEmailAddress "tanaka@example.co.jp" -AllowedDomains $domains
    Assert-Equal -Name "送信者一致: 完全一致ドメイン" -Expected $true -Actual $r

    # 大文字混じり
    $r = Test-SenderMatch -SenderEmailAddress "Tanaka@EXAMPLE.CO.JP" -AllowedDomains $domains
    Assert-Equal -Name "送信者一致: 大文字混じりでも一致" -Expected $true -Actual $r

    # サブドメイン扱い ケース1: "@ドメイン" で完全一致（上のケースと同義だが明示的に再確認）
    $r = Test-SenderMatch -SenderEmailAddress "user@example.co.jp" -AllowedDomains @("example.co.jp")
    Assert-Equal -Name "送信者一致: サブドメイン扱いケース1(@ドメイン)" -Expected $true -Actual $r

    # サブドメイン扱い ケース2: "mail.example.co.jp" は "example.co.jp" 指定で一致(.ドメイン)
    $r = Test-SenderMatch -SenderEmailAddress "user@mail.example.co.jp" -AllowedDomains @("example.co.jp")
    Assert-Equal -Name "送信者一致: サブドメイン扱いケース2(.ドメイン)" -Expected $true -Actual $r

    # 不一致: 紛らわしい別ドメイン（"example.co.jp"を含むが別物）
    $r = Test-SenderMatch -SenderEmailAddress "user@notexample.co.jp" -AllowedDomains $domains
    Assert-Equal -Name "送信者一致: 紛らわしい別ドメインは不一致" -Expected $false -Actual $r

    # 不一致: 全く無関係のドメイン
    $r = Test-SenderMatch -SenderEmailAddress "user@other-company.example" -AllowedDomains $domains
    Assert-Equal -Name "送信者一致: 無関係ドメインは不一致" -Expected $false -Actual $r

    # Exchange DN でフォールバック不能→不一致（本体側の解決処理がCOM依存のため、
    # ここでは「解決できず $null が渡ってきた」ケースと「解決できず生DN文字列の
    # ままだった」ケースの両方を検証する）
    $r = Test-SenderMatch -SenderEmailAddress $null -AllowedDomains $domains
    Assert-Equal -Name "送信者一致: Exchange DN解決失敗(null)は不一致" -Expected $false -Actual $r

    $r = Test-SenderMatch -SenderEmailAddress "/O=EXCHANGE/OU=FIRST ADMINISTRATIVE GROUP/CN=RECIPIENTS/CN=someone" -AllowedDomains $domains
    Assert-Equal -Name "送信者一致: Exchange DN解決失敗(生DN文字列)は不一致" -Expected $false -Actual $r

    # ============================================================
    # 2. New-MailEventLine（着信記録JSONLの1行生成）
    # ============================================================
    $receivedTime = Get-Date "2026-08-29 18:52:52"
    $line = New-MailEventLine -ReceivedTime $receivedTime -FromDomain "example.co.jp" -FromName "山田太郎" `
        -Subject "至急のご確認をお願いいたします【請求書】" -Folder "received\2026-08-29_1852_至急のご確認" `
        -Files @("invoice.pdf", "photo.jpg") -SavedCount 2

    # JSONとしてパース可能
    $parsed = $null
    $parseThrew = $false
    try { $parsed = $line | ConvertFrom-Json } catch { $parseThrew = $true }
    Assert-True -Name "イベント記録: JSONとしてパース可能" -Condition (-not $parseThrew) -Detail "line=[$line]"

    # 必須キー全あり
    $requiredKeys = @("receivedTime", "fromDomain", "fromName", "subject", "folder", "files", "savedCount", "ts")
    foreach ($k in $requiredKeys) {
        $hasKey = ($null -ne $parsed) -and ([bool]($parsed.PSObject.Properties.Name -contains $k))
        Assert-True -Name "イベント記録: 必須キー[$k]がある" -Condition $hasKey
    }

    # 日本語件名・氏名が壊れない（往復一致）
    Assert-Equal -Name "イベント記録: 日本語件名が壊れない" -Expected "至急のご確認をお願いいたします【請求書】" -Actual $parsed.subject
    Assert-Equal -Name "イベント記録: 日本語氏名が壊れない" -Expected "山田太郎" -Actual $parsed.fromName
    Assert-Equal -Name "イベント記録: savedCountが一致" -Expected 2 -Actual $parsed.savedCount
    Assert-Equal -Name "イベント記録: filesの件数が一致" -Expected 2 -Actual (@($parsed.files).Count)

    # 1行に収まっている（JSONL要件。改行を含まない）
    Assert-True -Name "イベント記録: 1行に収まっている(改行を含まない)" -Condition (-not ($line.Contains("`n")))

    # ============================================================
    # 3. Get-EventFolderName（フォルダ名生成: 40字切り・サニタイズ・衝突連番）
    # ============================================================
    $folderTestDir = Join-Path $testRoot "folder-name"
    New-Item -ItemType Directory -Path $folderTestDir -Force | Out-Null
    $rt = Get-Date "2026-08-29 18:52:00"

    # 40字切り（全角込みで長い件名。サニタイズ後に先頭40字へ切り詰められること）
    $longSubject = "あ" * 60
    $name = Get-EventFolderName -ParentPath $folderTestDir -ReceivedTime $rt -Subject $longSubject
    $prefixLen = "2026-08-29_1852_".Length
    $subjectPart = $name.Substring($prefixLen)
    Assert-True -Name "フォルダ名生成: 件名部分が40字以内に切り詰められる" -Condition ($subjectPart.Length -le 40) -Detail "len=$($subjectPart.Length) name=[$name]"
    Assert-True -Name "フォルダ名生成: 日時プレフィックスが正しい" -Condition ($name.StartsWith("2026-08-29_1852_")) -Detail "name=[$name]"

    # サニタイズ（禁止文字を含む件名）
    $dirtySubject = 'a<b>:c"d'
    $name = Get-EventFolderName -ParentPath $folderTestDir -ReceivedTime $rt -Subject $dirtySubject
    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    $hasInvalid = $false
    foreach ($ch in $invalidChars) { if ($name.IndexOf($ch) -ge 0) { $hasInvalid = $true } }
    Assert-True -Name "フォルダ名生成: 禁止文字がサニタイズされる" -Condition (-not $hasInvalid) -Detail "name=[$name]"

    # 空件名でも壊れない
    $name = Get-EventFolderName -ParentPath $folderTestDir -ReceivedTime $rt -Subject ""
    Assert-True -Name "フォルダ名生成: 空件名でも生成できる" -Condition (-not [string]::IsNullOrWhiteSpace($name)) -Detail "name=[$name]"

    # 同名衝突連番: 同じ日時・件名で2回作ると2回目は "(2)" が付く
    $subject3 = "衝突テスト用の件名"
    $baseName = Get-EventFolderName -ParentPath $folderTestDir -ReceivedTime $rt -Subject $subject3
    New-Item -ItemType Directory -Path (Join-Path $folderTestDir $baseName) -Force | Out-Null
    $secondName = Get-EventFolderName -ParentPath $folderTestDir -ReceivedTime $rt -Subject $subject3
    Assert-Equal -Name "フォルダ名生成: 同名衝突は(2)連番になる" -Expected ($baseName + " (2)") -Actual $secondName

    # (2)まで衝突していれば(3)まで進む
    New-Item -ItemType Directory -Path (Join-Path $folderTestDir $secondName) -Force | Out-Null
    $thirdName = Get-EventFolderName -ParentPath $folderTestDir -ReceivedTime $rt -Subject $subject3
    Assert-Equal -Name "フォルダ名生成: 連番は(3)まで進む" -Expected ($baseName + " (3)") -Actual $thirdName

    # ============================================================
    # 4. DryRun不変条件
    #    書き込み系統（Initialize-TargetFolder・Save-AttachmentToTarget・
    #    Add-ProcessedId・Add-MailEventRecord）が -DryRun:$true のとき
    #    ファイルシステムに一切書き込まない（一時ファイル・イベント記録ファイルも作らない）ことを検証する。
    # ============================================================
    $dryRunRoot = Join-Path $testRoot "dryrun-invariant"
    New-Item -ItemType Directory -Path $dryRunRoot -Force | Out-Null

    function Get-DirSnapshot {
        param([string]$Path)
        # 注意: 空フォルダの場合 `return @(パイプライン)` はPowerShellの仕様上 $null に
        # 収束してしまうことがある（実測確認済み）。`,$items`（単項カンマ演算子）で
        # 配列であることを保証してから返す。
        $items = @(Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty FullName | Sort-Object)
        return , $items
    }

    $beforeSnapshot = @(Get-DirSnapshot -Path $dryRunRoot)

    # Initialize-TargetFolder
    $notYetExistingFolder = Join-Path $dryRunRoot "2026-08-29_1900_dryrun"
    $folderResult = Initialize-TargetFolder -Path $notYetExistingFolder -DryRun
    Assert-Equal -Name "DryRun不変: Initialize-TargetFolderはExisted=falseを返す" -Expected $false -Actual $folderResult.Existed
    Assert-Equal -Name "DryRun不変: Initialize-TargetFolderはCreated=falseを返す" -Expected $false -Actual $folderResult.Created
    Assert-Equal -Name "DryRun不変: フォルダは実際には作られていない" -Expected $false -Actual (Test-Path -LiteralPath $notYetExistingFolder)

    # Save-AttachmentToTarget: -WriteTempFile には「呼ばれたら即FAILする」ブロックを渡す
    $writerWasCalled = $false
    $writerThatMustNotRun = {
        param($tmpPath)
        $script:writerWasCalled = $true
        throw "DryRunなのにWriteTempFileが呼ばれてしまった: $tmpPath"
    }
    $tmpFolderForTest = Join-Path $dryRunRoot "tmp"
    $saveResult = Save-AttachmentToTarget -TargetFolder $notYetExistingFolder -SafeFileName "invoice.pdf" `
        -TempFolder $tmpFolderForTest -WriteTempFile $writerThatMustNotRun -EstimatedSize 12345 -DryRun
    Assert-Equal -Name "DryRun不変: Save-AttachmentToTargetはWriteTempFileを呼ばない" -Expected $false -Actual $writerWasCalled
    Assert-True -Name "DryRun不変: 戻り値はSaveEstimateかSkipEstimate" `
        -Condition ($saveResult.Action -eq "SaveEstimate" -or $saveResult.Action -eq "SkipEstimate") -Detail "Action=$($saveResult.Action)"
    Assert-Equal -Name "DryRun不変: 一時フォルダも作られない" -Expected $false -Actual (Test-Path -LiteralPath $tmpFolderForTest)

    # Add-ProcessedId
    $processedIdsPathForTest = Join-Path $dryRunRoot "processed-ids.txt"
    $addResult = Add-ProcessedId -Path $processedIdsPathForTest -EntryId "DUMMY-ENTRY-ID" -DryRun
    Assert-Equal -Name "DryRun不変: Add-ProcessedIdはRecorded=falseを返す" -Expected $false -Actual $addResult.Recorded
    Assert-Equal -Name "DryRun不変: processed-ids.txtも作られない" -Expected $false -Actual (Test-Path -LiteralPath $processedIdsPathForTest)

    # Add-MailEventRecord（mail-watch固有の4本目の書き込み系統）
    $eventLogPathForTest = Join-Path $dryRunRoot "logs\inbox-events.jsonl"
    $eventResult = Add-MailEventRecord -Path $eventLogPathForTest -JsonLine '{"dummy":true}' -DryRun
    Assert-Equal -Name "DryRun不変: Add-MailEventRecordはRecorded=falseを返す" -Expected $false -Actual $eventResult.Recorded
    Assert-Equal -Name "DryRun不変: inbox-events.jsonlも作られない" -Expected $false -Actual (Test-Path -LiteralPath $eventLogPathForTest)

    $afterSnapshot = @(Get-DirSnapshot -Path $dryRunRoot)
    if ($beforeSnapshot.Count -eq 0 -and $afterSnapshot.Count -eq 0) {
        $snapshotDiff = @()
    }
    else {
        $snapshotDiff = @(Compare-Object -ReferenceObject $beforeSnapshot -DifferenceObject $afterSnapshot -ErrorAction SilentlyContinue)
    }
    Assert-Equal -Name "DryRun不変: 4関数呼び出し前後でファイル/フォルダが1つも増減していない" -Expected 0 -Actual $snapshotDiff.Count

    # ============================================================
    # 5. Add-MailEventRecord（本実行時の動作・BOM無しUTF-8確認）
    # ============================================================
    $realEventLogPath = Join-Path $testRoot "realrun-events.jsonl"
    $line1 = New-MailEventLine -ReceivedTime (Get-Date "2026-08-29 09:00:00") -FromDomain "example.co.jp" `
        -FromName "テスト太郎" -Subject "1通目" -Folder "received\1" -Files @("a.pdf") -SavedCount 1
    $line2 = New-MailEventLine -ReceivedTime (Get-Date "2026-08-29 10:00:00") -FromDomain "example.co.jp" `
        -FromName "テスト花子" -Subject "2通目" -Folder "received\2" -Files @("b.pdf", "c.pdf") -SavedCount 2

    [void](Add-MailEventRecord -Path $realEventLogPath -JsonLine $line1)
    [void](Add-MailEventRecord -Path $realEventLogPath -JsonLine $line2)

    Assert-True -Name "イベント記録(本実行): ファイルが作られる" -Condition (Test-Path -LiteralPath $realEventLogPath -PathType Leaf)

    # 注意: このファイルは仕様によりBOM無しUTF-8で書かれる。Get-Content を
    # -Encoding 指定なしで呼ぶと、Windows PowerShell 5.1はBOMが無い場合に
    # システムの既定コードページ（日本語環境ではCP932）とみなして読むため、
    # 日本語部分が文字化けする（実測確認済み。.ps1のBOM問題と同じ原因の逆方向）。
    # 読み込み側は必ず -Encoding UTF8 を明示すること（README.mdにも明記）。
    $writtenLines = Get-Content -LiteralPath $realEventLogPath -Encoding UTF8
    Assert-Equal -Name "イベント記録(本実行): 1メール=1行で2行になる" -Expected 2 -Actual (@($writtenLines).Count)

    $parsedLine2 = $writtenLines[1] | ConvertFrom-Json
    Assert-Equal -Name "イベント記録(本実行): 2行目の内容が正しい" -Expected "2通目" -Actual $parsedLine2.subject

    # BOM無しUTF-8であることの確認（先頭3バイトが EF BB BF ではない）
    $bytes = [System.IO.File]::ReadAllBytes($realEventLogPath)
    $hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
    Assert-Equal -Name "イベント記録(本実行): BOM無しUTF-8で書かれる" -Expected $false -Actual $hasBom

}
finally {
    # 後始末（$env:TEMP配下に作った専用一時フォルダのみ削除。他の場所には一切触れない）
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# ------------------------------------------------------------
# 結果サマリ
# ------------------------------------------------------------
$failedResults = @($script:Results | Where-Object { -not $_.Pass })
$passCount = $script:Results.Count - $failedResults.Count

if ($failedResults.Count -gt 0) {
    Write-Host ""
    Write-Host "=== FAILED ===" -ForegroundColor Red
    foreach ($f in $failedResults) {
        Write-Host ("[FAIL] {0}" -f $f.Name) -ForegroundColor Red
        if ($f.Detail) {
            Write-Host ("       {0}" -f $f.Detail) -ForegroundColor Red
        }
    }
    Write-Host ""
    Write-Host ("結果サマリ: 総数 {0} / PASS {1} / FAIL {2}" -f $script:Results.Count, $passCount, $failedResults.Count) -ForegroundColor Red
    exit 1
}
else {
    Write-Host ""
    Write-Host ("結果サマリ: 総数 {0} / PASS {1} / FAIL 0 ― 全PASS" -f $script:Results.Count, $passCount) -ForegroundColor Green
    exit 0
}
