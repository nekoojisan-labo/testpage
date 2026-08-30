#requires -Version 5.1
<#
    tests\Run-Tests.ps1

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
# 簡易テストランナー
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
$testRoot = Join-Path $env:TEMP ("mailattach-tests-" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $testRoot -Force | Out-Null

try {

    # ============================================================
    # 1. Get-CodeFromSubject（件名からのコード抽出）
    # ============================================================
    $defaultPatterns = @('(DJ|BJ|CH)\d{5}(?!\d)')

    $r = Get-CodeFromSubject -Subject "RE: DJ26779 の件" -Patterns $defaultPatterns
    Assert-Equal -Name "コード抽出: 通常の件名" -Expected "DJ26779" -Actual $r.Code
    Assert-Equal -Name "コード抽出: 通常の件名 (複数フラグ)" -Expected $false -Actual $r.MultipleFound

    $r = Get-CodeFromSubject -Subject "dj26779 の見積について" -Patterns $defaultPatterns
    Assert-Equal -Name "コード抽出: 小文字は大文字化して返す" -Expected "DJ26779" -Actual $r.Code

    $r = Get-CodeFromSubject -Subject "DJ123456 は6桁なので誤検知しない" -Patterns $defaultPatterns
    Assert-Equal -Name "コード抽出: 6桁は誤検知しない" -Expected $null -Actual $r.Code

    $r = Get-CodeFromSubject -Subject "BJ12345とCH99999の2件について" -Patterns $defaultPatterns
    Assert-Equal -Name "コード抽出: 複数コードは先頭を採用" -Expected "BJ12345" -Actual $r.Code
    Assert-Equal -Name "コード抽出: 複数コードで複数フラグが立つ" -Expected $true -Actual $r.MultipleFound

    $r = Get-CodeFromSubject -Subject "本日の会議室予約のお願い" -Patterns $defaultPatterns
    Assert-Equal -Name "コード抽出: 無関係の件名はnull" -Expected $null -Actual $r.Code

    $r = Get-CodeFromSubject -Subject "【至急】DJ26779 ご確認をお願いいたします" -Patterns $defaultPatterns
    Assert-Equal -Name "コード抽出: 全角混じりの件名でも動く" -Expected "DJ26779" -Actual $r.Code

    # ============================================================
    # 2. ConvertTo-SafeName（ファイル名サニタイズ）
    # ============================================================
    $safe = ConvertTo-SafeName -Name 'a<b>:c"/\|?*d'
    $hasInvalidChar = $false
    foreach ($ch in [System.IO.Path]::GetInvalidFileNameChars()) {
        if ($safe.IndexOf($ch) -ge 0) { $hasInvalidChar = $true }
    }
    Assert-True -Name "サニタイズ: 禁止文字が残っていない" -Condition (-not $hasInvalidChar) -Detail "結果=[$safe]"
    Assert-Equal -Name "サニタイズ: 先頭a・末尾dは保持" -Expected $true -Actual ($safe.StartsWith("a") -and $safe.EndsWith("d"))

    $safe = ConvertTo-SafeName -Name "invoice..."
    Assert-Equal -Name "サニタイズ: 末尾ドット除去" -Expected "invoice" -Actual $safe

    $safe = ConvertTo-SafeName -Name "invoice   "
    Assert-Equal -Name "サニタイズ: 末尾空白除去" -Expected "invoice" -Actual $safe

    $safe = ConvertTo-SafeName -Name ""
    Assert-Equal -Name "サニタイズ: 空文字はattachmentになる" -Expected "attachment" -Actual $safe

    # ============================================================
    # 3. Get-SaveFileName（保存ファイル名決定・実サイズ/実ハッシュベース・実フォルダ使用）
    #    [レビュー反映] att.Sizeのような近似値ではなく、実際に書き出したファイルの
    #    実サイズ・MD5ハッシュを引数で受けて判定する形に変更されたため、
    #    テストも実ファイルの実サイズ・実ハッシュを使う。
    # ============================================================
    $saveDir = Join-Path $testRoot "savefile"
    New-Item -ItemType Directory -Path $saveDir -Force | Out-Null

    # 分岐A: 同名・同実サイズ・同ハッシュ → スキップ($null)。[レビュー反映・軽微3]
    # -MatchedName で実際に一致したファイル名（この場合はベース名そのもの）も検証する。
    $xPath = Join-Path $saveDir "x.pdf"
    Set-Content -LiteralPath $xPath -Value ("A" * 100) -NoNewline -Encoding Ascii
    $sizeA = (Get-Item -LiteralPath $xPath).Length
    $hashA = (Get-FileHash -LiteralPath $xPath -Algorithm MD5).Hash
    $matchedRefA = [ref]$null
    $name = Get-SaveFileName -Folder $saveDir -FileName "x.pdf" -ActualSize $sizeA -ActualHash $hashA -MatchedName $matchedRefA
    Assert-Equal -Name "保存名決定: 同名同実サイズ同ハッシュはスキップ(null)" -Expected $null -Actual $name
    Assert-Equal -Name "保存名決定: MatchedNameにベース名が入る" -Expected "x.pdf" -Actual $matchedRefA.Value

    # 分岐B: 同名・別実サイズ → 連番 (2)（サイズが異なる時点で確定するのでハッシュは不問）
    $name = Get-SaveFileName -Folder $saveDir -FileName "x.pdf" -ActualSize ($sizeA + 999) -ActualHash "unused-because-size-differs"
    Assert-Equal -Name "保存名決定: 同名別実サイズは連番(2)" -Expected "x (2).pdf" -Actual $name

    # 分岐C: 未存在 → そのまま
    $name = Get-SaveFileName -Folder $saveDir -FileName "y.pdf" -ActualSize 12345 -ActualHash "unused-because-not-existing"
    Assert-Equal -Name "保存名決定: 未存在はそのまま" -Expected "y.pdf" -Actual $name

    # 分岐D（新規）: 同名・同実サイズだがハッシュ不一致 → 別内容とみなし連番保存（安全側）
    $wPath = Join-Path $saveDir "w.pdf"
    Set-Content -LiteralPath $wPath -Value "1111" -NoNewline -Encoding Ascii
    $sizeW = (Get-Item -LiteralPath $wPath).Length
    $hashOfW = (Get-FileHash -LiteralPath $wPath -Algorithm MD5).Hash

    $diffContentPath = Join-Path $saveDir "_w_different_content.tmp"
    Set-Content -LiteralPath $diffContentPath -Value "2222" -NoNewline -Encoding Ascii
    $differentHashSameSize = (Get-FileHash -LiteralPath $diffContentPath -Algorithm MD5).Hash
    Remove-Item -LiteralPath $diffContentPath -Force

    Assert-True -Name "保存名決定: テスト前提(同サイズ4バイトで実際に別ハッシュ)" `
        -Condition ($differentHashSameSize -ne $hashOfW) -Detail "hashOfW=$hashOfW other=$differentHashSameSize"

    $name = Get-SaveFileName -Folder $saveDir -FileName "w.pdf" -ActualSize $sizeW -ActualHash $differentHashSameSize
    Assert-Equal -Name "保存名決定: 同実サイズでもハッシュ不一致なら連番保存(安全側)" -Expected "w (2).pdf" -Actual $name

    # 分岐E: 連番が(3)まで進むケース（z.pdf・z (2).pdf の両方が別実サイズで使用中）
    Set-Content -LiteralPath (Join-Path $saveDir "z.pdf") -Value ("B" * 50) -NoNewline -Encoding Ascii
    Set-Content -LiteralPath (Join-Path $saveDir "z (2).pdf") -Value ("C" * 80) -NoNewline -Encoding Ascii
    $sizeZ = (Get-Item -LiteralPath (Join-Path $saveDir "z.pdf")).Length
    $sizeZ2 = (Get-Item -LiteralPath (Join-Path $saveDir "z (2).pdf")).Length
    $newSize = $sizeZ + $sizeZ2 + 1   # 既存のどちらの実サイズとも一致しない値
    $name = Get-SaveFileName -Folder $saveDir -FileName "z.pdf" -ActualSize $newSize -ActualHash "unused-because-size-differs"
    Assert-Equal -Name "保存名決定: 連番は(3)まで進む" -Expected "z (3).pdf" -Actual $name

    # 分岐F（新規・軽微3の再現テスト）: ベース名(v.pdf)とは別内容だが、連番側(v (2).pdf)と
    # 実サイズ・ハッシュが一致する場合 → スキップし、MatchedNameには実際に一致した
    # "v (2).pdf"（連番側）が入ること（ベース名 "v.pdf" ではないこと）を検証する。
    # これが無いと、ログに「実際に一致したファイル名」でなくベース名が出てしまうバグになる。
    $vPath = Join-Path $saveDir "v.pdf"
    $vPath2 = Join-Path $saveDir "v (2).pdf"
    Set-Content -LiteralPath $vPath -Value "base-content-XYZ" -NoNewline -Encoding Ascii
    Set-Content -LiteralPath $vPath2 -Value "numbered-content-12345" -NoNewline -Encoding Ascii
    $sizeV2 = (Get-Item -LiteralPath $vPath2).Length
    $hashV2 = (Get-FileHash -LiteralPath $vPath2 -Algorithm MD5).Hash
    $matchedRefF = [ref]$null
    $name = Get-SaveFileName -Folder $saveDir -FileName "v.pdf" -ActualSize $sizeV2 -ActualHash $hashV2 -MatchedName $matchedRefF
    Assert-Equal -Name "保存名決定(軽微3): 連番側と一致した場合はスキップ(null)" -Expected $null -Actual $name
    Assert-Equal -Name "保存名決定(軽微3): MatchedNameはベース名でなく連番側になる" -Expected "v (2).pdf" -Actual $matchedRefF.Value

    # ============================================================
    # 4. Limit-PathLength（パス長の切り詰め）
    # ============================================================
    $shortPath = "C:\short\dir\file.pdf"
    $r = Limit-PathLength -FullPath $shortPath -Max 250
    Assert-Equal -Name "パス長: 短いパスは変更されない" -Expected $shortPath -Actual $r.Path
    Assert-Equal -Name "パス長: 短いパスはTruncated=false" -Expected $false -Actual $r.Truncated

    $longName = ("N" * 300) + ".pdf"
    $longFullPath = Join-Path "C:\short\dir" $longName
    $r = Limit-PathLength -FullPath $longFullPath -Max 250
    Assert-True -Name "パス長: 250字超は切り詰められる" -Condition ($r.Path.Length -le 250) -Detail "結果長=$($r.Path.Length)"
    Assert-Equal -Name "パス長: Truncated=trueになる" -Expected $true -Actual $r.Truncated
    Assert-True -Name "パス長: 拡張子は保持される" -Condition ($r.Path.EndsWith(".pdf")) -Detail "結果=[$($r.Path)]"

    # ============================================================
    # 5. Select-MailsToProcess（"候補"に対する処理件数上限の選別）
    #    [レビュー反映・バグ2] この関数自体のsort+cap実装は変わっていないが、
    #    実機リハーサルで「上限がコード照合の前に効いてしまい、新しい側にある
    #    コード一致メールが上限超過側に落ちてDryRun集計が実態と合わない」
    #    という問題が判明したため、本体(Save-MailAttachments.ps1)側の役割を
    #    変更した: 上限はもう「受信トレイの全未処理メール」には適用せず、
    #    「件名がコードに一致し、かつ保存対象の添付がある候補」だけに適用する。
    #    コード不一致・添付なしのメールは候補にすらならず、上限に関係なく
    #    即座にprocessed-idsへ記録してよい（走査自体は実測1〜3秒/100通と安価）。
    #    そのため、このテストの $mailInfos は「（本体側で既に絞り込まれた）候補」を
    #    模したものとして扱う。
    # ============================================================
    $mailInfos = @()
    $baseDate = Get-Date "2026-01-01"
    for ($i = 0; $i -lt 60; $i++) {
        $mailInfos += [PSCustomObject]@{
            EntryID      = "ENTRY-$i"
            ReceivedTime = $baseDate.AddDays($i)
            Code         = "DJ{0:D5}" -f (10000 + $i)   # 候補である証として、疑似的にコードも持たせておく
        }
    }
    # 入力順をシャッフルし、「日付昇順に並べ替えてから選別する」ことを検証する
    $shuffled = @($mailInfos | Sort-Object { Get-Random })

    $sel = Select-MailsToProcess -MailInfos $shuffled -Max 50
    Assert-Equal -Name "候補選択: 60件中50件を選択" -Expected 50 -Actual $sel.Selected.Count
    Assert-Equal -Name "候補選択: 残り10件" -Expected 10 -Actual $sel.Remaining
    Assert-Equal -Name "候補選択: 最古の1件が先頭(日付昇順)" -Expected "ENTRY-0" -Actual $sel.Selected[0].EntryID
    Assert-Equal -Name "候補選択: 50件目はENTRY-49" -Expected "ENTRY-49" -Actual $sel.Selected[49].EntryID

    $selBackfill = Select-MailsToProcess -MailInfos $shuffled -Max 50 -Backfill
    Assert-Equal -Name "候補選択: Backfillで60件全件" -Expected 60 -Actual $selBackfill.Selected.Count
    Assert-Equal -Name "候補選択: Backfillで残り0件" -Expected 0 -Actual $selBackfill.Remaining

    # ============================================================
    # 6. Test-InlineAttachment（インライン画像判定）
    # ============================================================
    $r = Test-InlineAttachment -HiddenProp $true -ContentId $null -HtmlBody $null
    Assert-Equal -Name "インライン判定: hidden=true は除外" -Expected $true -Actual $r

    $r = Test-InlineAttachment -HiddenProp $false -ContentId "logo001" -HtmlBody "<html><body><img src='cid:logo001'></body></html>"
    Assert-Equal -Name "インライン判定: hidden=false でもcid参照ありは除外" -Expected $true -Actual $r

    $r = Test-InlineAttachment -HiddenProp $null -ContentId $null -HtmlBody $null
    Assert-Equal -Name "インライン判定: 両方取得不能なら安全側(保存=false)" -Expected $false -Actual $r

    # 備考: $ExcludeInlineImages = $false 相当（この判定自体を呼び出すかどうかの分岐）は
    #       Save-MailAttachments.ps1（本体側）の責務であり、Outlook非依存のこのテストの対象外。

    # ============================================================
    # 7. [レビュー反映・修正3] DryRun不変条件
    #    フォルダ作成(Initialize-TargetFolder)・添付保存(Save-AttachmentToTarget)・
    #    processed-ids追記(Add-ProcessedId) は、-DryRun 指定時にファイルシステムへ
    #    一切書き込まない（一時ファイルすら作らない）ことを検証する。
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

    # Initialize-TargetFolder: 存在しないフォルダでも作らない
    $notYetExistingFolder = Join-Path $dryRunRoot "DJ99999_"
    $folderResult = Initialize-TargetFolder -Path $notYetExistingFolder -DryRun
    Assert-Equal -Name "DryRun不変: Initialize-TargetFolderはExisted=falseを返す" -Expected $false -Actual $folderResult.Existed
    Assert-Equal -Name "DryRun不変: Initialize-TargetFolderはCreated=falseを返す" -Expected $false -Actual $folderResult.Created
    Assert-Equal -Name "DryRun不変: フォルダは実際には作られていない" -Expected $false -Actual (Test-Path -LiteralPath $notYetExistingFolder)

    # Save-AttachmentToTarget: -WriteTempFile には「呼ばれたら即FAILする」ブロックを渡し、
    # DryRunでは絶対に呼ばれない（＝一時ファイルすら作らない）ことを検証する
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

    # Add-ProcessedId: 追記しない・ファイルも作らない
    $processedIdsPathForTest = Join-Path $dryRunRoot "processed-ids.txt"
    $addResult = Add-ProcessedId -Path $processedIdsPathForTest -EntryId "DUMMY-ENTRY-ID" -DryRun
    Assert-Equal -Name "DryRun不変: Add-ProcessedIdはRecorded=falseを返す" -Expected $false -Actual $addResult.Recorded
    Assert-Equal -Name "DryRun不変: processed-ids.txtも作られない" -Expected $false -Actual (Test-Path -LiteralPath $processedIdsPathForTest)

    $afterSnapshot = @(Get-DirSnapshot -Path $dryRunRoot)
    if ($beforeSnapshot.Count -eq 0 -and $afterSnapshot.Count -eq 0) {
        $snapshotDiff = @()
    }
    else {
        $snapshotDiff = @(Compare-Object -ReferenceObject $beforeSnapshot -DifferenceObject $afterSnapshot -ErrorAction SilentlyContinue)
    }
    Assert-Equal -Name "DryRun不変: 3関数呼び出し前後でファイル/フォルダが1つも増減していない" -Expected 0 -Actual $snapshotDiff.Count

    # ============================================================
    # 8. Save-AttachmentToTarget（本実行時の一時ファイル経由フローの動作確認）
    #    Outlookの Attachment オブジェクトの代わりに Set-Content する偽の
    #    WriteTempFile を注入し、COM非依存のまま一時ファイル→実サイズ/ハッシュ判定
    #    →Move-Item の流れ全体を検証する。
    # ============================================================
    $realRunDir = Join-Path $testRoot "realrun"
    $realTargetFolder = Join-Path $realRunDir "DJ11111_"
    $realTmpFolder = Join-Path $realRunDir "tmp"
    New-Item -ItemType Directory -Path $realTargetFolder -Force | Out-Null

    $fakeContentA = "hello-world-content-A"
    $writerA = { param($tmpPath) Set-Content -LiteralPath $tmpPath -Value $fakeContentA -NoNewline -Encoding Ascii }

    # 1回目: 新規保存
    $r1 = Save-AttachmentToTarget -TargetFolder $realTargetFolder -SafeFileName "report.pdf" -TempFolder $realTmpFolder -WriteTempFile $writerA
    Assert-Equal -Name "添付保存(本実行): 1回目はSaved" -Expected "Saved" -Actual $r1.Action
    Assert-True -Name "添付保存(本実行): 保存先に実ファイルができている" -Condition (Test-Path -LiteralPath $r1.SavedPath -PathType Leaf) -Detail "SavedPath=$($r1.SavedPath)"
    Assert-Equal -Name "添付保存(本実行): 保存後、一時フォルダにファイルが残っていない" -Expected 0 -Actual (@(Get-ChildItem -LiteralPath $realTmpFolder -File -ErrorAction SilentlyContinue).Count)

    # 2回目: 全く同じ内容を再度保存しようとする → 冪等スキップ
    $r2 = Save-AttachmentToTarget -TargetFolder $realTargetFolder -SafeFileName "report.pdf" -TempFolder $realTmpFolder -WriteTempFile $writerA
    Assert-Equal -Name "添付保存(本実行): 同内容の再実行はSkip(冪等性)" -Expected "Skip" -Actual $r2.Action
    Assert-Equal -Name "添付保存(本実行): スキップ後も保存先ファイルは1つだけ" -Expected 1 -Actual (@(Get-ChildItem -LiteralPath $realTargetFolder -Filter "report*.pdf").Count)

    # 3回目: 同名だが別内容（別サイズ）の添付 → 連番で保存される
    $fakeContentB = "different-content-with-another-length-B"
    $writerB = { param($tmpPath) Set-Content -LiteralPath $tmpPath -Value $fakeContentB -NoNewline -Encoding Ascii }
    $r3 = Save-AttachmentToTarget -TargetFolder $realTargetFolder -SafeFileName "report.pdf" -TempFolder $realTmpFolder -WriteTempFile $writerB
    Assert-Equal -Name "添付保存(本実行): 別内容は連番で保存される" -Expected "Saved" -Actual $r3.Action
    Assert-True -Name "添付保存(本実行): 連番ファイル名になっている" -Condition ($r3.SavedPath -like "*report (2).pdf") -Detail "SavedPath=$($r3.SavedPath)"
    Assert-Equal -Name "添付保存(本実行): 最終的に保存先ファイルは2つ" -Expected 2 -Actual (@(Get-ChildItem -LiteralPath $realTargetFolder -Filter "report*.pdf").Count)

    # ============================================================
    # 9. [レビュー反映・バグ1／mail-watch対応] Show-SaveNotification（通知の判定ロジック）
    #    [シグネチャ変更] メッセージ文言はもう内部で組み立てず、呼び出し側が
    #    -Message で渡す（mail-watch等ツールごとに文言が違うため一般化した）。
    #    実際のポップアップ表示（WScript.Shell.Popup）はOutlook非依存のこの
    #    テストでは呼べない/呼ぶべきでないため、-ShowPopup にモックの
    #    スクリプトブロックを注入し、「呼ばれたか・何回か・どんなメッセージで」を検証する。
    # ============================================================

    # 抑制条件1: DryRun中は出さない・ShowPopupも呼ばない
    $mockCallCount = 0
    $mockShowPopup = { param($m) $script:mockCallCount++ }

    $script:mockCallCount = 0
    $n1 = Show-SaveNotification -SavedFileCount 3 -Message "test-message" -EnableNotification $true -ShowPopup $mockShowPopup -DryRun
    Assert-Equal -Name "通知: DryRun中はSkipped" -Expected "Skipped" -Actual $n1.Action
    Assert-Equal -Name "通知: DryRun中はShowPopupを呼ばない" -Expected 0 -Actual $script:mockCallCount

    # 抑制条件2: $EnableNotification=$false のときは出さない
    $script:mockCallCount = 0
    $n2 = Show-SaveNotification -SavedFileCount 3 -Message "test-message" -EnableNotification $false -ShowPopup $mockShowPopup
    Assert-Equal -Name "通知: EnableNotification=falseはSkipped" -Expected "Skipped" -Actual $n2.Action
    Assert-Equal -Name "通知: EnableNotification=falseはShowPopupを呼ばない" -Expected 0 -Actual $script:mockCallCount

    # 抑制条件3: 保存0件のときは出さない
    $script:mockCallCount = 0
    $n3 = Show-SaveNotification -SavedFileCount 0 -Message "test-message" -EnableNotification $true -ShowPopup $mockShowPopup
    Assert-Equal -Name "通知: 保存0件はSkipped" -Expected "Skipped" -Actual $n3.Action
    Assert-Equal -Name "通知: 保存0件はShowPopupを呼ばない" -Expected 0 -Actual $script:mockCallCount

    # 通常ケース: 条件を満たせば1回だけ呼ばれ、渡したメッセージがそのままShowPopupへ渡る
    $script:mockCallCount = 0
    $script:mockLastMessage = $null
    $mockShowPopupCapture = { param($m) $script:mockCallCount++; $script:mockLastMessage = $m }
    $n4 = Show-SaveNotification -SavedFileCount 4 -Message "添付ファイルを 4 件保存しました。DJ26779_" -EnableNotification $true -ShowPopup $mockShowPopupCapture
    Assert-Equal -Name "通知: 条件を満たせばShown" -Expected "Shown" -Actual $n4.Action
    Assert-Equal -Name "通知: ShowPopupは1回だけ呼ばれる" -Expected 1 -Actual $script:mockCallCount
    Assert-Equal -Name "通知: 渡したメッセージがそのままShowPopupへ渡る" -Expected "添付ファイルを 4 件保存しました。DJ26779_" -Actual $script:mockLastMessage
    Assert-Equal -Name "通知: 戻り値のMessageも一致する" -Expected "添付ファイルを 4 件保存しました。DJ26779_" -Actual $n4.Message

    # ShowPopupが失敗した場合はFailedを返し、例外は外に漏れない
    $mockShowPopupThrows = { param($m) throw "COM呼び出しが失敗したという想定のテスト例外" }
    $n6 = Show-SaveNotification -SavedFileCount 2 -Message "test-message" -EnableNotification $true -ShowPopup $mockShowPopupThrows
    Assert-Equal -Name "通知: ShowPopup失敗時はFailed" -Expected "Failed" -Actual $n6.Action
    Assert-True -Name "通知: Failed時にErrorメッセージが入る" -Condition (-not [string]::IsNullOrEmpty($n6.Error)) -Detail "Error=[$($n6.Error)]"

    # ============================================================
    # 10. [レビュー反映・mail-watch対応] Format-TruncatedList（一覧の先頭N件+他n件要約）
    #     元はShow-SaveNotification内蔵だったフォルダ一覧要約ロジックを汎用化したもの。
    # ============================================================
    $shortList = @("A", "B", "C")
    $r = Format-TruncatedList -Items $shortList -Threshold 6 -ShowCount 5
    Assert-Equal -Name "一覧要約: 閾値以下はそのまま(件数)" -Expected 3 -Actual (@($r).Count)
    Assert-Equal -Name "一覧要約: 閾値以下はそのまま(中身)" -Expected "A,B,C" -Actual (($r) -join ",")

    $manyFolders = @(1..8 | ForEach-Object { "CODE{0:D5}_" -f $_ })
    $r = Format-TruncatedList -Items $manyFolders -Threshold 6 -ShowCount 5
    Assert-Equal -Name "一覧要約: 8件は先頭5件+他3件で計6件" -Expected 6 -Actual (@($r).Count)
    Assert-Equal -Name "一覧要約: 先頭5件目まで含まれる" -Expected "CODE00005_" -Actual $r[4]
    Assert-Equal -Name "一覧要約: 最後は「他3件」" -Expected "他 3 件" -Actual $r[5]
    Assert-True -Name "一覧要約: 6件目以降(CODE00006_)は含まれない" -Condition (($r -join ",") -notlike "*CODE00006_*")

    $r = Format-TruncatedList -Items @() -Threshold 6 -ShowCount 5
    Assert-Equal -Name "一覧要約: 空配列は空配列のまま" -Expected 0 -Actual (@($r).Count)

    $r = Format-TruncatedList -Items $null -Threshold 6 -ShowCount 5
    Assert-Equal -Name "一覧要約: nullも空配列扱い" -Expected 0 -Actual (@($r).Count)

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
