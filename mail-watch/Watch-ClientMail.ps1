#requires -Version 5.1
<#
.SYNOPSIS
    Outlook受信トレイを監視し、監視対象ドメインからの添付付きメールをローカルへ保存、
    ポップアップ通知とJSONL着信記録を残す（開発者本人のPC用の個人ツール）。

.DESCRIPTION
    起動中のOutlook（クラシック版）にのみ接続し、受信トレイの直近N日分から
    差出人ドメインが一致し、かつ保存対象の添付があるメールを検知して、
    添付ファイルをローカルの received フォルダへ保存する。
    メール本体（既読/未読・内容・場所）は一切変更しない。

    運用イメージ: メールが届く→本ツールを実行（当面は手動）→ポップアップで気づく→
    「メール確認」とClaudeに伝える→Claudeが logs\inbox-events.jsonl を読んで
    次のたたき台準備に入る。

.PARAMETER DryRun
    実際の保存・フォルダ作成・processed-ids追記・イベント記録追記・通知を
    一切行わず（一時ファイルも作らない）、対象件数などの見込みだけを
    コンソールとログに出力する。

.NOTES
    このスクリプトは開発機（Outlookプロファイル無し）では実行できない。
    姉妹ツール mail-attach-saver（クライアント納品キット）とMailAttachLib.ps1を
    共有している（本フォルダのものは複製。正本はmail-attach-saver側）。
#>

[CmdletBinding()]
param(
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"

# ============================================================
# 0. 初期化: ライブラリ・設定の読み込み
# ============================================================
. (Join-Path $PSScriptRoot "MailAttachLib.ps1")

$configPath = Join-Path $PSScriptRoot "config.local.ps1"
if (-not (Test-Path -LiteralPath $configPath)) {
    throw "config.local.ps1 が見つかりません。config.sample.ps1 をコピーして作成してください: $configPath"
}
. $configPath

# 必須設定の存在チェック（設定漏れ・タイプミスの早期検知）
$requiredVars = @(
    "WatchSenderDomains", "RequireAttachment", "RetentionDays", "MaxMailsPerRun",
    "EnableNotification", "SaveRoot", "EventLogPath", "ProcessedIdsPath", "LogFolder"
)
foreach ($v in $requiredVars) {
    if (-not (Get-Variable -Name $v -Scope Script -ErrorAction SilentlyContinue)) {
        throw "config.local.ps1 に必須設定 `$$v がありません。config.sample.ps1 と見比べてください。"
    }
}

# ログ・processed-ids・イベント記録・保存先用フォルダを用意
# （mail-attach-saverと異なり $SaveRoot はローカルフォルダなので、Z:のような
#   「未接続かもしれない」リスクが無い。存在しなければ素直に作成する）
foreach ($dir in @($LogFolder, $SaveRoot, (Split-Path -Parent $ProcessedIdsPath), (Split-Path -Parent $EventLogPath))) {
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

$logFile = Join-Path $LogFolder ("run-" + (Get-Date -Format "yyyyMM") + ".log")

# 添付の一時保存先はローカル(logs\tmp)。実行毎にクリーンアップする
# （前回異常終了時の残骸を残さないため）。DryRunでは一切使わない。
$tmpFolder = Join-Path $LogFolder "tmp"

function Write-Log {
    param([string]$Message)
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    try {
        Add-Content -LiteralPath $logFile -Value $line -Encoding UTF8
    }
    catch {
        # ログ書き込み自体の失敗で本処理を止めない
    }
    Write-Verbose $line
}

$modeLabel = if ($DryRun) { "DryRun" } else { "通常" }
Write-Log ("===== 実行開始 (モード: {0}) =====" -f $modeLabel)

# ============================================================
# 1. 多重起動防止（kitと同じMutex方式。名前はこちら専用にする）
# ============================================================
$mutex = New-Object System.Threading.Mutex($false, "MailWatchClient_SingleInstance_Mutex")
$acquired = $false
try {
    $acquired = $mutex.WaitOne(0)
}
catch [System.Threading.AbandonedMutexException] {
    # 前回実行が異常終了していた場合。ミューテックスの所有権は取得できているので続行する。
    $acquired = $true
}

if (-not $acquired) {
    Write-Log "前回の実行がまだ処理中の可能性があるため、今回の実行はスキップします（多重起動防止）。"
    return
}

try {
    # 実行毎に一時フォルダをクリーンアップ（前回異常終了時の残骸掃除も兼ねる。DryRunでは触らない）
    if (-not $DryRun -and (Test-Path -LiteralPath $tmpFolder)) {
        try {
            Remove-Item -LiteralPath $tmpFolder -Recurse -Force -ErrorAction Stop
        }
        catch {
            Write-Log ("一時フォルダの初期クリーンアップに失敗しました（続行します）: {0}" -f $_.Exception.Message)
        }
    }

    # ============================================================
    # 2. Outlookへの接続
    #
    #    [レビュー反映] 実機で新たに判明した障害への対処。
    #    ★正本は本ファイルと Save-MailAttachments.ps1（mail-attach-saver側）の
    #      両方に同一実装として置いてある。改修時は両方へ反映すること
    #      （Outlook COM依存のためMailAttachLib.ps1には入れず、そちらの
    #      Outlook非依存を保つ）。
    #
    #    実測事実: Outlook(クラシック)が起動中・応答中・受信トレイ表示中でも、
    #    GetActiveObject("Outlook.Application")がMK_E_UNAVAILABLE(0x800401E3)で
    #    失敗する状態が発生する（ROT=Running Object Table への登録が無い状態。
    #    メインウィンドウを閉じて通知領域の常駐だけ残った後に再表示した、等の
    #    経緯で発生し、以後この状態が続く）。
    #
    #    この状態でも、OUTLOOK.EXEプロセスが存在する間に
    #    New-Object -ComObject Outlook.Application を実行すると、一時的に2本目の
    #    プロセスが現れるがOutlook自身の単一インスタンスガードにより数秒で1本へ
    #    収束し、返ってきたオブジェクトで実際にSessionへアクセスできる
    #    （実測: 受信トレイ739件取得に成功）。
    #
    #    ただし、OUTLOOK.EXEプロセスが1本も無い状態でNew-Objectを呼ぶと
    #    「Outlookが起動していない前提で静かに終了する」という設計方針に反して
    #    本当にOutlookを起動してしまうため、プロセス存在チェックで必ずゲートする。
    #
    #    3段構え:
    #      1. GetActiveObject を試す → 成功なら採用（Status="Normal"）
    #      2. 失敗時、OUTLOOK.EXEプロセスが1件以上あるときだけ New-Object を試し、
    #         取得後に実際にSessionへアクセスして接続実効性を確認する
    #         → 成功なら採用（Status="Piggyback"）
    #      3. どちらも不可:
    #         - プロセスが無い → Status="NotRunning"（＝本当に未起動）
    #         - プロセスはあるのに接続できない → Status="ProcessExistsButFailed"
    #           （＝Outlook再起動で解消する見込み、という別メッセージにする）
    # ============================================================
    function Connect-Outlook {
        $outlookApp = $null

        try {
            $outlookApp = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
            if ($null -ne $outlookApp) {
                return [PSCustomObject]@{ Outlook = $outlookApp; Status = "Normal" }
            }
        }
        catch {
            $outlookApp = $null
        }

        # プロセス存在チェック（これより後のNew-Objectは、プロセスが本当に
        # 存在する場合にのみ試みる。存在しない状態で呼ぶとOutlookを本当に
        # 起動させてしまうため、このゲートを必ず先に通す）
        $outlookProcesses = @(Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue)
        if ($outlookProcesses.Count -eq 0) {
            return [PSCustomObject]@{ Outlook = $null; Status = "NotRunning" }
        }

        try {
            $outlookApp = New-Object -ComObject Outlook.Application
            # 接続実効性の確認: 実際にSessionへアクセスできるかどうかを見る
            # （実測でこの経路が成功していたSession.GetDefaultFolder(6)を使う）
            $verifyInbox = $outlookApp.Session.GetDefaultFolder(6)
            if ($null -ne $verifyInbox) {
                return [PSCustomObject]@{ Outlook = $outlookApp; Status = "Piggyback" }
            }
            return [PSCustomObject]@{ Outlook = $null; Status = "ProcessExistsButFailed" }
        }
        catch {
            return [PSCustomObject]@{ Outlook = $null; Status = "ProcessExistsButFailed" }
        }
    }

    $connectResult = Connect-Outlook
    $outlook = $connectResult.Outlook

    switch ($connectResult.Status) {
        "Normal" {
            Write-Log "Outlook接続: 通常"
        }
        "Piggyback" {
            Write-Log "Outlook接続: 相乗り(ROT未登録を回避)"
        }
        "NotRunning" {
            Write-Log "Outlookが起動していないため、処理をスキップして終了します（次回実行時に処理されます）。"
            return
        }
        "ProcessExistsButFailed" {
            Write-Log "Outlookは起動していますが接続できませんでした（Outlook再起動で解消する見込み）。"
            return
        }
    }

    $namespace = $null
    $inbox = $null
    $comItemsByEntryId = @{}

    try {
        $namespace = $outlook.GetNamespace("MAPI")
        $inbox = $namespace.GetDefaultFolder(6)  # 6 = olFolderInbox

        # ============================================================
        # 3. 対象メールの取得（Restrictで直近N日に絞り込み、受信トレイ全件走査を避ける）
        # ============================================================
        $cutoff = (Get-Date).AddDays(-$RetentionDays)
        # 注意: Outlook Restrict の日付書式はWindowsの地域設定に依存する（mail-attach-saver
        # と同じ既知の注意点）。"g"（カレントカルチャ準拠）を使うが、対象件数が
        # 不自然な場合はこれを疑うこと。
        $filterDateString = $cutoff.ToString("g")
        $filter = "[ReceivedTime] >= '" + $filterDateString + "'"

        $restrictedItems = $inbox.Items.Restrict($filter)

        # 処理済みEntryIDの読み込み（重複防止の主軸）
        $processedIds = New-Object System.Collections.Generic.HashSet[string]
        if (Test-Path -LiteralPath $ProcessedIdsPath) {
            foreach ($line in Get-Content -LiteralPath $ProcessedIdsPath) {
                $trimmed = $line.Trim()
                if ($trimmed.Length -gt 0) {
                    [void]$processedIds.Add($trimmed)
                }
            }
        }

        # COMアイテムを辞書に落とし込む（このあとの評価フェーズで使う）
        $allUnprocessedCount = 0

        foreach ($item in $restrictedItems) {
            try {
                if ($item.Class -ne 43) { continue }  # 43 = olMail 以外（会議通知等）は対象外
                $entryId = $item.EntryID
                if ($processedIds.Contains($entryId)) { continue }

                $allUnprocessedCount++
                $comItemsByEntryId[$entryId] = $item
            }
            catch {
                Write-Log ("メール一覧取得中にエラー（該当アイテムをスキップ）: {0}" -f $_.Exception.Message)
            }
        }

        # ============================================================
        # 4. 候補の評価（全件・上限なし）
        #    上限 $MaxMailsPerRun は「差出人ドメイン一致かつ保存対象添付あり」の
        #    候補にのみ適用する（mail-attach-saverのバグ2修正と同じ3フェーズ思想。
        #    ドメイン不一致・添付なしのメールは上限に関係なく即座にprocessed-ids記録）。
        # ============================================================
        $candidates = New-Object System.Collections.Generic.List[object]
        $nonCandidateCount = 0

        foreach ($entryId in @($comItemsByEntryId.Keys)) {
            $mailItem = $comItemsByEntryId[$entryId]
            $subjectForLog = ""

            try {
                $subject = $mailItem.Subject
                $subjectForLog = $subject
                $receivedTime = $mailItem.ReceivedTime

                # 差出人アドレスの解決（Exchange DN形式ならSMTPアドレスへ変換を試みる。
                # 取れなければ $null にし、以後は確実に不一致として扱う＝安全側）
                $resolvedAddr = $null
                try {
                    $rawAddr = $mailItem.SenderEmailAddress
                    if (-not [string]::IsNullOrWhiteSpace($rawAddr)) {
                        if ($rawAddr.StartsWith("/O=", [System.StringComparison]::OrdinalIgnoreCase)) {
                            try {
                                $exUser = $mailItem.Sender.GetExchangeUser()
                                if ($null -ne $exUser -and -not [string]::IsNullOrWhiteSpace($exUser.PrimarySmtpAddress)) {
                                    $resolvedAddr = $exUser.PrimarySmtpAddress
                                }
                            }
                            catch {
                                $resolvedAddr = $null
                            }
                        }
                        else {
                            $resolvedAddr = $rawAddr
                        }
                    }
                }
                catch {
                    $resolvedAddr = $null
                }

                $isMatch = Test-SenderMatch -SenderEmailAddress $resolvedAddr -AllowedDomains $WatchSenderDomains

                if (-not $isMatch) {
                    # ドメイン不一致＝監視対象外。上限を消費せずその場で処理済み記録する。
                    [void](Add-ProcessedId -Path $ProcessedIdsPath -EntryId $entryId -DryRun:$DryRun)
                    $nonCandidateCount++
                    continue
                }

                # 添付の選別（インライン画像を除く）。列挙自体は$RequireAttachmentの値に
                # 関わらず常に行い、「0件を候補から外すかどうか」だけを$RequireAttachmentで
                # 切り替える（$false時にも実在する添付を取りこぼさないようにするため）。
                $htmlBody = $null
                try { $htmlBody = $mailItem.HTMLBody } catch { $htmlBody = $null }

                $targetAttachments = New-Object System.Collections.Generic.List[object]
                $attachments = $mailItem.Attachments
                for ($i = 1; $i -le $attachments.Count; $i++) {
                    $att = $attachments.Item($i)
                    $hiddenProp = $null
                    $contentId = $null
                    try { $hiddenProp = $att.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x7FFE000B") } catch { $hiddenProp = $null }
                    try { $contentId = $att.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F") } catch { $contentId = $null }

                    $shouldExclude = Test-InlineAttachment -HiddenProp $hiddenProp -ContentId $contentId -HtmlBody $htmlBody
                    if (-not $shouldExclude) {
                        $targetAttachments.Add($att)
                    }
                }

                if ($RequireAttachment -and $targetAttachments.Count -eq 0) {
                    # 保存対象添付なし＝候補にしない（既定仕様）。上限を消費せずその場で処理済み記録する。
                    [void](Add-ProcessedId -Path $ProcessedIdsPath -EntryId $entryId -DryRun:$DryRun)
                    $nonCandidateCount++
                    continue
                }

                # ドメイン一致（かつ設定により添付あり）＝候補。上限の適用対象。
                $fromDomainForRecord = ""
                if (-not [string]::IsNullOrEmpty($resolvedAddr) -and $resolvedAddr.Contains("@")) {
                    $fromDomainForRecord = $resolvedAddr.Substring($resolvedAddr.LastIndexOf("@") + 1)
                }
                $fromNameForRecord = ""
                try { $fromNameForRecord = $mailItem.SenderName } catch { $fromNameForRecord = "" }

                $candidates.Add([PSCustomObject]@{
                    EntryID           = $entryId
                    ReceivedTime      = $receivedTime
                    Subject           = $subject
                    FromDomain        = $fromDomainForRecord
                    FromName          = $fromNameForRecord
                    TargetAttachments = $targetAttachments
                })
            }
            catch {
                Write-Log ("メール評価中にエラー（該当メールをスキップ・次回リトライ）: EntryID={0} 件名=[{1}] {2}" -f $entryId, $subjectForLog, $_.Exception.Message)
            }
        }

        # ============================================================
        # 5. 上限の適用（候補のみに対して）
        # ============================================================
        $selection = Select-MailsToProcess -MailInfos $candidates.ToArray() -Max $MaxMailsPerRun

        Write-Log ("候補メール（ドメイン一致かつ添付あり）: {0} 件（未処理合計 {1} 件中・対象外で {2} 件を即記録）" -f `
            $candidates.Count, $allUnprocessedCount, $nonCandidateCount)
        Write-Log ("今回処理: {0} 件・残り: {1} 件" -f $selection.Selected.Count, $selection.Remaining)
        if ($selection.Remaining -gt 0) {
            Write-Log ("上限超過: 残り {0} 件は次回実行で処理します。" -f $selection.Remaining)
        }

        # DryRun集計用
        $dryPlannedFolders = New-Object System.Collections.Generic.HashSet[string]
        $dryPlannedFiles = 0
        $drySkipEstimate = 0

        # 通知・イベント記録用集計
        $savedFileCount = 0
        $touchedSubjects = New-Object System.Collections.Generic.List[string]

        # ============================================================
        # 6. 候補ごとの保存処理（1通単位でtry/catch。失敗はログに残して継続、
        #    processed-idsには書かない＝次回自動リトライ）
        # ============================================================
        foreach ($candidate in $selection.Selected) {
            try {
                if ($candidate.TargetAttachments.Count -eq 0) {
                    # $RequireAttachment=$false で添付0件のまま候補になったケース。
                    # 空フォルダを作らず（仕様4と同じ考え方）、処理済みとして記録するだけにする。
                    [void](Add-ProcessedId -Path $ProcessedIdsPath -EntryId $candidate.EntryID -DryRun:$DryRun)
                    continue
                }

                $folderName = Get-EventFolderName -ParentPath $SaveRoot -ReceivedTime $candidate.ReceivedTime -Subject $candidate.Subject
                $targetDir = Join-Path $SaveRoot $folderName

                $folderResult = Initialize-TargetFolder -Path $targetDir -DryRun:$DryRun
                if (-not $DryRun -and $folderResult.Created) {
                    Write-Log ("フォルダ作成: {0}" -f $targetDir)
                }
                elseif ($DryRun -and -not $folderResult.Existed) {
                    [void]$dryPlannedFolders.Add($folderName)
                }

                $savedFileNames = New-Object System.Collections.Generic.List[string]
                $savedAnyForThisMail = $false

                foreach ($att in $candidate.TargetAttachments) {
                    $originalName = $att.FileName
                    $safeName = ConvertTo-SafeName -Name $originalName
                    $estimatedSize = [long]$att.Size

                    # Outlook COMへの依存はこのスクリプトブロックだけに閉じ込める。
                    $writeTempFile = { param($tmpPath) $att.SaveAsFile($tmpPath) }

                    $result = Save-AttachmentToTarget -TargetFolder $targetDir -SafeFileName $safeName `
                        -TempFolder $tmpFolder -WriteTempFile $writeTempFile -EstimatedSize $estimatedSize -DryRun:$DryRun

                    if ($DryRun) {
                        if ($result.Action -eq "SkipEstimate") {
                            $drySkipEstimate++
                        }
                        else {
                            $dryPlannedFiles++
                        }
                        continue
                    }

                    if ($result.Action -eq "Skip") {
                        Write-Log ("スキップ（保存済み・実サイズ/ハッシュ一致）: {0}\{1}" -f $folderName, $result.MatchedName)
                        continue
                    }

                    if ($result.Truncated) {
                        Write-Log ("パス長切り詰め: {0}" -f $result.SavedPath)
                    }
                    Write-Log ("保存: {0}" -f $result.SavedPath)
                    $savedFileCount++
                    $savedAnyForThisMail = $true
                    $savedFileNames.Add((Split-Path -Leaf $result.SavedPath))
                }

                # イベント記録（JSONL）: このメールで新規に保存されたファイルがある場合のみ記録する
                # （Claudeが読むのは「新着」の合図のため。全件が冪等スキップだった場合は
                #   新しく伝えるべき情報が無いのでノイズにしない）
                if (-not $DryRun -and $savedAnyForThisMail) {
                    $relativeFolder = Join-Path (Split-Path -Leaf $SaveRoot) $folderName
                    $eventLine = New-MailEventLine -ReceivedTime $candidate.ReceivedTime -FromDomain $candidate.FromDomain `
                        -FromName $candidate.FromName -Subject $candidate.Subject -Folder $relativeFolder `
                        -Files @($savedFileNames) -SavedCount $savedFileNames.Count
                    $eventResult = Add-MailEventRecord -Path $EventLogPath -JsonLine $eventLine
                    if ($eventResult.Recorded) {
                        Write-Log ("着信記録追記: {0}" -f $relativeFolder)
                    }
                }

                if ($savedAnyForThisMail -and -not $touchedSubjects.Contains($candidate.Subject)) {
                    $touchedSubjects.Add($candidate.Subject)
                }

                # メール本体には一切書き込まない（UnRead代入禁止・Save()禁止・移動削除禁止）。
                [void](Add-ProcessedId -Path $ProcessedIdsPath -EntryId $candidate.EntryID -DryRun:$DryRun)
            }
            catch {
                Write-Log ("エラー: EntryID={0} 件名=[{1}] {2}" -f $candidate.EntryID, $candidate.Subject, $_.Exception.Message)
            }
        }

        # ============================================================
        # 7. DryRun集計の出力
        # ============================================================
        if ($DryRun) {
            $summaryLines = @(
                "----- DryRun 集計 -----",
                ("対象メール（ドメイン一致かつ添付あり）: {0} 件（今回処理: {1} 件・残り: {2} 件）" -f `
                    $candidates.Count, $selection.Selected.Count, $selection.Remaining),
                ("作成予定フォルダ: {0} 件" -f $dryPlannedFolders.Count),
                ("保存予定ファイル数: {0} 件" -f $dryPlannedFiles),
                ("スキップ見込み（保存済みと思われる件数）: {0} 件" -f $drySkipEstimate),
                "※スキップ見込みは概算です（添付のSize情報のみによる判定のため）。実行時は一時ファイルへの書き出し後の実サイズ・MD5ハッシュで正確に判定します。"
            )
            if ($selection.Remaining -gt 0) {
                $summaryLines += ("上限超過につき今回対象外: {0} 件" -f $selection.Remaining)
            }
            foreach ($line in $summaryLines) {
                Write-Host $line
                Write-Log $line
            }
            if ($dryPlannedFolders.Count -gt 0) {
                Write-Host "作成予定フォルダ一覧:"
                foreach ($f in $dryPlannedFolders) { Write-Host ("  - {0}" -f $f) }
            }
        }

        # ============================================================
        # 8. 通知（1件以上保存した実行のみ・DryRunでは出さない・設定でOFF可）
        #    Show-SaveNotificationを流用。メッセージ文言はmail-watch独自
        #    （「法人案件Aメールn件・保存mファイル」+件名先頭30字を最大5件）。
        # ============================================================
        $subjectPreviewLines = @($touchedSubjects | ForEach-Object {
            if ($_.Length -gt 30) { $_.Substring(0, 30) + "…" } else { $_ }
        })
        $subjectPreviewLines = Format-TruncatedList -Items $subjectPreviewLines -Threshold 5 -ShowCount 5

        $notifyMessage = "法人案件Aメール {0} 件・保存 {1} ファイル`r`n`r`n" -f $touchedSubjects.Count, $savedFileCount
        $notifyMessage += ($subjectPreviewLines -join "`r`n")

        $notifyResult = Show-SaveNotification -SavedFileCount $savedFileCount -Message $notifyMessage `
            -EnableNotification ([bool]$EnableNotification) -DryRun:$DryRun -ShowPopup {
                param($popupMessage)
                Write-Log ("通知表示: 保存 {0} 件" -f $savedFileCount)
                $shell = New-Object -ComObject WScript.Shell
                # 64=vbInformation, 4096=vbSystemModal を加算(4160)。システムモーダルにすることで
                # 他アプリの背後に隠れて見落とされる事故を防ぐ（timeout=0=手動で閉じるまで残る仕様は維持）。
                [void]$shell.Popup($popupMessage, 0, "法人案件Aメール受信", 4160)
            }

        if ($notifyResult.Action -eq "Failed") {
            Write-Log ("通知ポップアップの表示に失敗しました: {0}" -f $notifyResult.Error)
        }

        Write-Log ("===== 実行終了 (保存 {0} 件・対象メール {1} 件) =====" -f $savedFileCount, $touchedSubjects.Count)
    }
    catch {
        Write-Log ("予期しないエラーのため処理を中断しました: {0}" -f $_.Exception.Message)
    }
    finally {
        # 一時フォルダ(logs\tmp)は実行終了時にも掃除しておく（DryRunでは触らない）。
        if (-not $DryRun -and (Test-Path -LiteralPath $tmpFolder)) {
            try { Remove-Item -LiteralPath $tmpFolder -Recurse -Force -ErrorAction Stop } catch {}
        }

        # COMオブジェクトの解放
        foreach ($key in @($comItemsByEntryId.Keys)) {
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($comItemsByEntryId[$key]) } catch {}
        }
        if ($null -ne $inbox) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($inbox) } catch {} }
        if ($null -ne $namespace) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) } catch {} }
        if ($null -ne $outlook) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) } catch {} }
    }
}
finally {
    if ($acquired) {
        try {
            $mutex.ReleaseMutex()
        }
        catch {}
    }
    $mutex.Dispose()
}
