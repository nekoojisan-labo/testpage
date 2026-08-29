#requires -Version 5.1
<#
.SYNOPSIS
    Outlook受信トレイの添付ファイルを、件名から抽出した管理コードごとのフォルダへ自動保存する。

.DESCRIPTION
    起動中のOutlook（クラシック版）にのみ接続し、受信トレイの直近N日分から
    件名パターンに一致するメールを検知して、添付ファイルを共有ドライブ上の
    コード別フォルダへ保存する。メール本体（既読/未読・内容・場所）は一切変更しない。

    タスクスケジューラから5分間隔で呼び出される運用を前提とし、Outlookが
    起動していなければ何もせず終了する（次回に任せる）。

.PARAMETER DryRun
    実際の保存・フォルダ作成・processed-ids追記・通知を一切行わず（一時ファイルも作らない）、
    対象件数や保存予定件数などの見込みだけをコンソールとログに出力する。
    スキップ見込みは添付のSize情報のみによる概算（実行時は実サイズ・ハッシュで正確に判定する）。

.PARAMETER Backfill
    1回あたりの処理件数上限（$MaxMailsPerRun）を無視し、対象メールを全件処理する。
    [レビュー反映] この上限は「件名がコードに一致し、かつ保存対象の添付がある候補」に
    のみ適用される（コード不一致・添付なしのメールは上限に関係なく毎回評価される）。
    導入初回のみ使用する想定。

.NOTES
    このスクリプトは開発機（Outlookプロファイル無し・Z:未接続）では実行できない。
    現地導入手順は docs\ONSITE-SETUP.md を参照。
#>

[CmdletBinding()]
param(
    [switch]$DryRun,
    [switch]$Backfill
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
    "SaveRoot", "RetentionDays", "SubjectPatterns", "ExcludeInlineImages",
    "MaxMailsPerRun", "EnableNotification", "ProcessedIdsPath", "LogFolder"
)
foreach ($v in $requiredVars) {
    if (-not (Get-Variable -Name $v -Scope Script -ErrorAction SilentlyContinue)) {
        throw "config.local.ps1 に必須設定 `$$v がありません。config.sample.ps1 と見比べてください。"
    }
}

# ログ・processed-ids用フォルダを用意
foreach ($dir in @($LogFolder, (Split-Path -Parent $ProcessedIdsPath))) {
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

$logFile = Join-Path $LogFolder ("run-" + (Get-Date -Format "yyyyMM") + ".log")

# [レビュー反映・修正1] 添付の一時保存先はローカル(logs\tmp)。$SaveRoot(Z:等)ではない。
# 実行毎にクリーンアップする（前回異常終了時の残骸を残さないため）。DryRunでは一切使わない。
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

$modeLabel = if ($DryRun -and $Backfill) { "DryRun+Backfill" } elseif ($DryRun) { "DryRun" } elseif ($Backfill) { "Backfill" } else { "通常" }
Write-Log ("===== 実行開始 (モード: {0}) =====" -f $modeLabel)

# ============================================================
# 1. 多重起動防止（タスクスケジューラの5分間隔実行が重ならないようにする）
#    前回実行が長引いている場合、今回はスキップして次回に任せる。
# ============================================================
$mutex = New-Object System.Threading.Mutex($false, "MailAttachSaver_SingleInstance_Mutex")
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
    # ============================================================
    # 2. 保存先ルートの健全性チェック（Outlookに触る前に確認する）
    #    config.local.ps1 がプレースホルダのまま／Z:が未接続の場合はここで検知する。
    # ============================================================
    $saveRootOk = $false
    try {
        $saveRootOk = Test-Path -LiteralPath $SaveRoot
    }
    catch {
        $saveRootOk = $false
    }
    if (-not $saveRootOk) {
        Write-Log ("保存先ルートが見つからないか、パスが不正です。処理を中断します: {0}" -f $SaveRoot)
        return
    }

    # [レビュー反映・修正1] 一時フォルダ(logs\tmp)を実行毎にクリーンアップする。
    # 前回が異常終了して一時ファイルが残っていた場合の掃除も兼ねる。DryRunでは触らない。
    if (-not $DryRun -and (Test-Path -LiteralPath $tmpFolder)) {
        try {
            Remove-Item -LiteralPath $tmpFolder -Recurse -Force -ErrorAction Stop
        }
        catch {
            Write-Log ("一時フォルダの初期クリーンアップに失敗しました（続行します）: {0}" -f $_.Exception.Message)
        }
    }

    # ============================================================
    # 3. Outlookへの接続（起動中のインスタンスにのみ接続。未起動なら静かに終了）
    # ============================================================
    $outlook = $null
    try {
        $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    }
    catch {
        Write-Log "Outlookが起動していないため、処理をスキップして終了します（次回実行時に処理されます）。"
        return
    }

    $namespace = $null
    $inbox = $null
    $comItemsByEntryId = @{}

    try {
        $namespace = $outlook.GetNamespace("MAPI")
        $inbox = $namespace.GetDefaultFolder(6)  # 6 = olFolderInbox

        # ============================================================
        # 4. 対象メールの取得（Restrictで直近N日に絞り込み、受信トレイ全件走査を避ける）
        # ============================================================
        $cutoff = (Get-Date).AddDays(-$RetentionDays)
        # 注意: Outlook Restrict の日付書式はWindowsの地域設定に依存する。
        # "g"（一般日付・短い時刻／カレントカルチャ準拠）を使うが、現地PCの地域設定次第では
        # 一致しない可能性があるため、現地導入時は必ずDryRunの対象件数で妥当性を確認すること
        # （docs\ONSITE-SETUP.md 手順3参照）。
        $filterDateString = $cutoff.ToString("g")
        $filter = "[ReceivedTime] >= '" + $filterDateString + "'"

        $restrictedItems = $inbox.Items.Restrict($filter)

        # 処理済みEntryIDの読み込み（重複防止の主軸。ファイルが無くても仕様7の冪等性で実害なし）
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
        # 5. [レビュー反映・バグ2] 候補の評価（全件・上限なし）
        #    実機リハーサルで、上限 $MaxMailsPerRun が「受信トレイの未処理メール全体」に
        #    先に効いてしまい、新しく届いたコード一致メールが（より多数の）コード不一致の
        #    古いメールに上限を先食いされて後回しになる問題が判明した。
        #    そこで上限は「コード一致かつ保存対象添付あり」という保存作業の候補にのみ
        #    適用するよう変更する。コード不一致・添付なしのメールは保存作業を伴わない
        #    （＝上限を消費する理由が無い）ため、上限に関係なくその場で処理済みとして
        #    記録してよい。この評価（件名照合・添付列挙）自体は実測1〜3秒/100通と安価。
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

                $codeResult = Get-CodeFromSubject -Subject $subject -Patterns $SubjectPatterns
                if ($codeResult.MultipleFound) {
                    Write-Log ("複数コード検出（先頭を採用）: EntryID={0} 件名=[{1}] 採用コード={2}" -f $entryId, $subject, $codeResult.Code)
                }

                if ($null -eq $codeResult.Code) {
                    # コード不一致＝保存作業を伴わない。上限を消費せずその場で処理済み記録する。
                    [void](Add-ProcessedId -Path $ProcessedIdsPath -EntryId $entryId -DryRun:$DryRun)
                    $nonCandidateCount++
                    continue
                }

                # 添付の選別（インライン画像を除く）。HTMLBodyはメール1通につき1回だけ取得する。
                $htmlBody = $null
                if ($ExcludeInlineImages) {
                    try { $htmlBody = $mailItem.HTMLBody } catch { $htmlBody = $null }
                }

                $attachments = $mailItem.Attachments
                $targetAttachments = New-Object System.Collections.Generic.List[object]

                for ($i = 1; $i -le $attachments.Count; $i++) {
                    $att = $attachments.Item($i)
                    $shouldExclude = $false

                    if ($ExcludeInlineImages) {
                        $hiddenProp = $null
                        $contentId = $null
                        try { $hiddenProp = $att.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x7FFE000B") } catch { $hiddenProp = $null }
                        try { $contentId = $att.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F") } catch { $contentId = $null }

                        $shouldExclude = Test-InlineAttachment -HiddenProp $hiddenProp -ContentId $contentId -HtmlBody $htmlBody
                    }

                    if (-not $shouldExclude) {
                        $targetAttachments.Add($att)
                    }
                }

                if ($targetAttachments.Count -eq 0) {
                    # 添付なし＝保存作業を伴わない。上限を消費せずその場で処理済み記録する（仕様4）。
                    [void](Add-ProcessedId -Path $ProcessedIdsPath -EntryId $entryId -DryRun:$DryRun)
                    $nonCandidateCount++
                    continue
                }

                # ここまで残ったものだけが「保存作業の候補」＝上限の適用対象
                $candidates.Add([PSCustomObject]@{
                    EntryID           = $entryId
                    ReceivedTime      = $receivedTime
                    Code              = $codeResult.Code
                    TargetAttachments = $targetAttachments
                    Subject           = $subject
                })
            }
            catch {
                Write-Log ("メール評価中にエラー（該当メールをスキップ・次回リトライ）: EntryID={0} 件名=[{1}] {2}" -f $entryId, $subjectForLog, $_.Exception.Message)
            }
        }

        # ============================================================
        # 6. [レビュー反映・バグ2] 上限の適用（候補のみに対して）
        # ============================================================
        $selection = Select-MailsToProcess -MailInfos $candidates.ToArray() -Max $MaxMailsPerRun -Backfill:$Backfill

        Write-Log ("候補メール（コード一致かつ添付あり）: {0} 件（未処理合計 {1} 件中・コード不一致/添付なしで {2} 件を即記録）" -f `
            $candidates.Count, $allUnprocessedCount, $nonCandidateCount)
        Write-Log ("今回処理: {0} 件・残り: {1} 件" -f $selection.Selected.Count, $selection.Remaining)
        if ($selection.Remaining -gt 0) {
            Write-Log ("上限超過: 残り {0} 件は次回実行で処理します（-Backfill 指定で今回まとめて処理可）。" -f $selection.Remaining)
        }

        # DryRun集計用
        $dryPlannedFolders = New-Object System.Collections.Generic.HashSet[string]
        $dryPlannedFiles = 0
        $drySkipEstimate = 0

        # 通知用集計
        $savedFileCount = 0
        $touchedFolders = New-Object System.Collections.Generic.List[string]

        # ============================================================
        # 7. 候補ごとの保存処理（1通単位でtry/catch。失敗はログに残して継続、
        #    processed-idsには書かない＝次回自動リトライ）
        # ============================================================
        foreach ($candidate in $selection.Selected) {
            try {
                $folderName = ConvertTo-SafeName -Name ($candidate.Code + "_")
                $targetDir = Join-Path $SaveRoot $folderName

                # [レビュー反映・修正3] フォルダ作成はMailAttachLib.ps1のInitialize-TargetFolderに集約。
                # DryRun時はここが一切書き込まないことをtests側で検証済み。
                $folderResult = Initialize-TargetFolder -Path $targetDir -DryRun:$DryRun
                if (-not $DryRun -and $folderResult.Created) {
                    Write-Log ("フォルダ作成: {0}" -f $targetDir)
                }
                elseif ($DryRun -and -not $folderResult.Existed) {
                    [void]$dryPlannedFolders.Add($folderName)
                }

                $savedAnyForThisMail = $false

                foreach ($att in $candidate.TargetAttachments) {
                    $originalName = $att.FileName
                    $safeName = ConvertTo-SafeName -Name $originalName
                    $estimatedSize = [long]$att.Size

                    # [レビュー反映・修正1] att.Size(MAPI PR_ATTACH_SIZE由来の近似値)ではなく、
                    # 一時ファイルへ実際に書き出した実サイズ・実ハッシュで判定する。
                    # Outlook COMへの依存はこのスクリプトブロックだけに閉じ込め、
                    # 判定ロジック自体はMailAttachLib.ps1側（Outlook非依存・テスト対象）に置く。
                    # Move-Itemが失敗した場合はここで例外が投げられ、下のメール単位try/catchに
                    # 委ねられる（＝そのメールはprocessed-idsに記録されず次回自動リトライされる）。
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
                        # [レビュー反映・軽微3] $safeName（ベース名）ではなく、実際に一致した
                        # ファイル名（連番側のこともある）$result.MatchedName を表示する。
                        Write-Log ("スキップ（保存済み・実サイズ/ハッシュ一致）: {0}\{1}" -f $folderName, $result.MatchedName)
                        continue
                    }

                    if ($result.Truncated) {
                        Write-Log ("パス長切り詰め: {0}" -f $result.SavedPath)
                    }
                    Write-Log ("保存: {0}" -f $result.SavedPath)
                    $savedFileCount++
                    $savedAnyForThisMail = $true
                }

                if ($savedAnyForThisMail -and -not $touchedFolders.Contains($folderName)) {
                    $touchedFolders.Add($folderName)
                }

                # メール本体には一切書き込まない（UnRead代入禁止・Save()禁止・移動削除禁止）。
                # 添付の保存が完了した（=このメールの評価が正常に終わった）ので処理済みとして記録する。
                [void](Add-ProcessedId -Path $ProcessedIdsPath -EntryId $candidate.EntryID -DryRun:$DryRun)
            }
            catch {
                Write-Log ("エラー: EntryID={0} 件名=[{1}] {2}" -f $candidate.EntryID, $candidate.Subject, $_.Exception.Message)
            }
        }

        # ============================================================
        # 8. DryRun集計の出力
        #    [レビュー反映・バグ2] 「一致n件のうち、今回m件・残りk件」が分かる表示にする。
        # ============================================================
        if ($DryRun) {
            $summaryLines = @(
                "----- DryRun 集計 -----",
                ("対象メール（コード一致かつ添付あり）: {0} 件（今回処理: {1} 件・残り: {2} 件）" -f `
                    $candidates.Count, $selection.Selected.Count, $selection.Remaining),
                ("作成予定フォルダ: {0} 件" -f $dryPlannedFolders.Count),
                ("保存予定ファイル数: {0} 件" -f $dryPlannedFiles),
                ("スキップ見込み（保存済みと思われる件数）: {0} 件" -f $drySkipEstimate),
                "※スキップ見込みは概算です（添付のSize情報のみによる判定のため）。実行時は一時ファイルへの書き出し後の実サイズ・MD5ハッシュで正確に判定します。"
            )
            if ($selection.Remaining -gt 0) {
                $summaryLines += ("上限超過につき今回対象外: {0} 件（-Backfillで解除可）" -f $selection.Remaining)
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
        # 9. 通知（1件以上保存した実行のみ・DryRunでは出さない・設定でOFF可）
        #
        #    [レビュー反映・バグ1] 診断の経緯: 実機リハーサルで「保存4件・フォルダ2件が
        #    あったのにポップアップが出た形跡がなく、ログにも通知関連の行が一切なかった」
        #    という報告を受け、$DryRun/$EnableNotification/$savedFileCount の分岐を
        #    変数レベルで追跡したが論理的な誤りは見当たらなかった。次に、この開発機で
        #    （Outlook・Z:・タスク登録には触れない範囲で）WScript.Shell.Popup単体を
        #    -WindowStyle Hidden の子プロセスから呼ぶ実験を行い、Popupは実際に可視の
        #    トップレベルウィンドウを生成し、閉じられるまで正しくブロックすることを
        #    実測で確認した（＝hidden起動でPopup自体が機能しない、という仮説は再現しなかった）。
        #    一方で当時のコードには「表示を試みた」ことを示すログが一切無く、
        #    (a)このコードに到達していない (b)到達しPopupは呼ばれたがまだ閉じられて
        #    おらずブロック中（前面に出ず見落とされた可能性を含む）(c)失敗はしたが
        #    その失敗ログ自体の書き込みが別要因で失敗した、のいずれであるかをログから
        #    区別する手段が無かった。この「区別できないこと」自体が実質的な欠陥のため、
        #    判定ロジックをMailAttachLib.ps1のShow-SaveNotificationへ切り出してテスト可能にし
        #    （実際の表示はモック可能なスクリプトブロックとして注入）、
        #    表示直前ログの追加と、見落とし防止のためのシステムモーダル化をセットで行う。
        # ============================================================
        $notifyResult = Show-SaveNotification -SavedFileCount $savedFileCount -TouchedFolders @($touchedFolders) `
            -EnableNotification ([bool]$EnableNotification) -DryRun:$DryRun -ShowPopup {
                param($popupMessage)
                Write-Log ("通知表示: 保存 {0} 件" -f $savedFileCount)
                $shell = New-Object -ComObject WScript.Shell
                # 64=vbInformation, 4096=vbSystemModal を加算(4160)。システムモーダルにすることで
                # 他アプリの背後に隠れて見落とされる事故を防ぐ（timeout=0=手動で閉じるまで残る仕様は維持）。
                [void]$shell.Popup($popupMessage, 0, "メール添付保存", 4160)
            }

        if ($notifyResult.Action -eq "Failed") {
            Write-Log ("通知ポップアップの表示に失敗しました: {0}" -f $notifyResult.Error)
        }

        Write-Log ("===== 実行終了 (保存 {0} 件・フォルダ {1} 件) =====" -f $savedFileCount, $touchedFolders.Count)
    }
    catch {
        Write-Log ("予期しないエラーのため処理を中断しました: {0}" -f $_.Exception.Message)
    }
    finally {
        # [レビュー反映・修正1] 一時フォルダ(logs\tmp)は実行終了時にも掃除しておく（DryRunでは触らない）。
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
