#requires -Version 5.1
<#
    MailAttachLib.ps1  （mail-watch 用の複製）

    ★正本は ..\mail-attach-saver\MailAttachLib.ps1 側（本ツールと同じ階層にある想定）。
    ★改修時は両方へ反映すること（特にGet-CodeFromSubject〜Show-SaveNotification/
      Format-TruncatedListまでの共通部分。Test-SenderMatch以降はmail-watch固有の
      追加分で、mail-attach-saver側には無い＝差分があって正常）。

    Outlook添付ファイル関連ツールの「純粋ロジック」関数群。
    Outlook COM に一切依存しない（=このファイル単体で tests\Run-Tests.ps1 から
    合成データを使ってテストできる）。

    Watch-ClientMail.ps1 からは次のように読み込む:
        . (Join-Path $PSScriptRoot "MailAttachLib.ps1")

    ここに書いてよいこと:
        - 文字列処理・正規表現・パス計算・配列の選別など、入力と出力だけで完結する関数
    ここに書いてはいけないこと:
        - Outlook COM オブジェクトの生成・参照
        - ネットワークドライブ／実際のメールボックスへのアクセス
#>

# ------------------------------------------------------------
# Get-CodeFromSubject
#   件名から管理コードを抽出する。mail-watch本体では現状使用しないが、
#   mail-attach-saver側との複製の一致を保つため残してある。
#   - $Patterns は正規表現の配列（大文字小文字は無視して照合する）
#   - 複数パターン・同一パターンの複数マッチをすべて集め、件名内で最も先頭に
#     現れたものを採用する（"先頭のコードを採用"仕様に対応）
#   - 2件以上マッチした場合は MultipleFound = $true を立てる
#   戻り値: PSCustomObject @{ Code = <string|$null>; MultipleFound = <bool> }
# ------------------------------------------------------------
function Get-CodeFromSubject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Subject,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string[]]$Patterns
    )

    if ($null -eq $Subject) { $Subject = "" }
    if ($null -eq $Patterns) { $Patterns = @() }

    $foundList = New-Object System.Collections.Generic.List[object]

    foreach ($pattern in $Patterns) {
        if ([string]::IsNullOrWhiteSpace($pattern)) { continue }

        $regexMatches = [System.Text.RegularExpressions.Regex]::Matches(
            $Subject,
            $pattern,
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
        )

        foreach ($m in $regexMatches) {
            $foundList.Add([PSCustomObject]@{
                Index = $m.Index
                Value = $m.Value.ToUpperInvariant()
            })
        }
    }

    if ($foundList.Count -eq 0) {
        return [PSCustomObject]@{
            Code          = $null
            MultipleFound = $false
        }
    }

    # 件名内での出現位置（Index）昇順で並べ、先頭に現れたものを採用
    $sorted = $foundList | Sort-Object Index

    return [PSCustomObject]@{
        Code          = ($sorted | Select-Object -First 1).Value
        MultipleFound = ($sorted.Count -gt 1)
    }
}

# ------------------------------------------------------------
# ConvertTo-SafeName
#   Windows のファイル名・フォルダ名として使えない文字を "_" に置換し、
#   末尾のドット・空白を除去する。結果が空文字になった場合は "attachment" とする。
# ------------------------------------------------------------
function ConvertTo-SafeName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Name
    )

    if ($null -eq $Name) { $Name = "" }

    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    $sb = New-Object System.Text.StringBuilder

    foreach ($ch in $Name.ToCharArray()) {
        if ($invalidChars -contains $ch) {
            [void]$sb.Append('_')
        }
        else {
            [void]$sb.Append($ch)
        }
    }

    $result = $sb.ToString()

    # 末尾のドット・空白を除去（Windowsはこれらで終わる名前を作れないため）
    $result = $result.TrimEnd('.', ' ')

    if ([string]::IsNullOrEmpty($result)) {
        $result = "attachment"
    }

    return $result
}

# ------------------------------------------------------------
# Get-SaveFileName
#   保存先フォルダ内の実ファイルを確認し、実際に使うファイル名を決定する。
#   （冪等性＝事故復旧の要。同名・同実サイズ・同ハッシュなら「保存済み」とみなしスキップする）
#
#   COMの Attachment.Size は MAPI PR_ATTACH_SIZE 由来で、実際の添付内容（保存後の
#   ファイルの実バイト数）より大きいことがあるとMicrosoft公式にも明記されている
#   （S/MIME等）。そのため本関数は att.Size のような近似値ではなく、「実際に
#   書き出したファイルの実サイズ・実ハッシュ」を呼び出し側から渡してもらう前提。
#   呼び出し側（Save-AttachmentToTarget）は一時ファイルへ保存してからその実サイズ・
#   実ハッシュを計算して本関数に渡す。
#
#   戻り値:
#     - 同名で実サイズも一致するファイルが既にある
#         → 実ハッシュ(MD5)も一致すれば $null （スキップ指示。保存済みとみなす）
#         → ハッシュが不一致なら「別内容がたまたま同サイズ」とみなし連番探索へ（安全側）
#     - 同名だが実サイズが異なるファイルがある     → "name (2).ext" 形式の連番名
#         （(2) も使用中なら (3)、(3) も使用中なら (4)…と繰り上げる。
#           その連番名について「既存・同実サイズ・同ハッシュ」であれば、そこもスキップ扱い
#           ＝過去に同じ内容を連番名で保存済みだった場合の再実行にも対応する）
#     - 同名のファイルが存在しない               → 渡された名前をそのまま返す
#
#   スキップ（$null）の場合、実際に一致した既存ファイル名（ベース名のこともあれば
#   "name (2).ext" 等の連番名のこともある）を呼び出し側へ伝えるための任意の
#   [ref]$MatchedName パラメータがある。省略可能。
#
#   $Folder はテスト時には一時ディレクトリを渡す想定。実フォルダを見る実装。
# ------------------------------------------------------------
function Get-SaveFileName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Folder,

        [Parameter(Mandatory = $true)]
        [string]$FileName,

        [Parameter(Mandatory = $true)]
        [long]$ActualSize,

        [Parameter(Mandatory = $true)]
        [string]$ActualHash,

        [Parameter(Mandatory = $false)]
        [ref]$MatchedName
    )

    function Test-SameExistingFile {
        param([string]$Path)
        $existingItem = Get-Item -LiteralPath $Path
        if ($existingItem.Length -ne $ActualSize) { return $false }
        $existingHash = (Get-FileHash -LiteralPath $Path -Algorithm MD5).Hash
        return ($existingHash -eq $ActualHash)
    }

    $fullPath = Join-Path $Folder $FileName
    if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
        return $FileName
    }
    if (Test-SameExistingFile -Path $fullPath) {
        if ($PSBoundParameters.ContainsKey('MatchedName')) { $MatchedName.Value = $FileName }
        return $null
    }

    $lastDot = $FileName.LastIndexOf('.')
    if ($lastDot -gt 0) {
        $baseName = $FileName.Substring(0, $lastDot)
        $ext = $FileName.Substring($lastDot)
    }
    else {
        $baseName = $FileName
        $ext = ""
    }

    $n = 2
    while ($true) {
        $candidateName = "{0} ({1}){2}" -f $baseName, $n, $ext
        $candidatePath = Join-Path $Folder $candidateName

        if (-not (Test-Path -LiteralPath $candidatePath -PathType Leaf)) {
            return $candidateName
        }
        if (Test-SameExistingFile -Path $candidatePath) {
            if ($PSBoundParameters.ContainsKey('MatchedName')) { $MatchedName.Value = $candidateName }
            return $null
        }

        $n++
    }
}

# ------------------------------------------------------------
# Limit-PathLength
#   フルパスが $Max 文字を超える場合、拡張子を保持したままファイル名部分だけを
#   切り詰める。ディレクトリ部分は変更しない。
#   戻り値: PSCustomObject @{ Path = <string>; Truncated = <bool> }
# ------------------------------------------------------------
function Limit-PathLength {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FullPath,

        [Parameter(Mandatory = $false)]
        [int]$Max = 250
    )

    if ($FullPath.Length -le $Max) {
        return [PSCustomObject]@{
            Path      = $FullPath
            Truncated = $false
        }
    }

    # 注意: [System.IO.Path]::GetDirectoryName 等（.NET Framework実装）は
    # 260文字前後を超える文字列を渡すと PathTooLongException を投げてしまう。
    # このメソッドの目的自体が「長すぎるパスの是正」なので、Path静的メソッドには
    # 頼らず素の文字列操作（LastIndexOf/Substring）で分解する。
    $lastSep = $FullPath.LastIndexOf('\')
    if ($lastSep -ge 0) {
        $dir = $FullPath.Substring(0, $lastSep)
        $fileName = $FullPath.Substring($lastSep + 1)
    }
    else {
        $dir = ""
        $fileName = $FullPath
    }

    $lastDot = $fileName.LastIndexOf('.')
    if ($lastDot -gt 0) {
        # 先頭が"."のファイル（隠しファイル等）は拡張子なし扱いにする
        $nameOnly = $fileName.Substring(0, $lastDot)
        $ext = $fileName.Substring($lastDot)
    }
    else {
        $nameOnly = $fileName
        $ext = ""
    }

    # ディレクトリ + 区切り文字(\) + 拡張子 の分だけファイル名本体に使える長さを計算
    $reserved = $dir.Length + 1 + $ext.Length
    $allowedNameLength = $Max - $reserved

    if ($allowedNameLength -lt 1) {
        # ディレクトリパス自体が長すぎる極端なケースの安全策（最低1文字は確保する）
        $allowedNameLength = 1
    }

    if ($nameOnly.Length -gt $allowedNameLength) {
        $nameOnly = $nameOnly.Substring(0, $allowedNameLength)
    }

    $newPath = $dir + '\' + $nameOnly + $ext

    return [PSCustomObject]@{
        Path      = $newPath
        Truncated = $true
    }
}

# ------------------------------------------------------------
# Select-MailsToProcess
#   受信日時の古い順に並べ替えたうえで、1回の実行で処理する件数の上限を適用する。
#   -Backfill 指定時は上限を解除し全件を対象にする。
#
#   mail-watchでは「候補（差出人ドメイン一致かつ保存対象添付あり）」に対して
#   この上限を適用する（mail-attach-saverのバグ2修正と同じ3フェーズ思想）。
#
#   $MailInfos は COM に依存しない PSCustomObject の配列（各要素は最低限
#   ReceivedTime [datetime] プロパティを持つこと。EntryID 等は本体側で付与）。
#
#   戻り値: PSCustomObject @{ Selected = <object[]>; Remaining = <int> }
# ------------------------------------------------------------
function Select-MailsToProcess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$MailInfos,

        [Parameter(Mandatory = $false)]
        [int]$Max = 50,

        [Parameter(Mandatory = $false)]
        [switch]$Backfill
    )

    if ($null -eq $MailInfos) { $MailInfos = @() }

    $sorted = @($MailInfos | Sort-Object ReceivedTime)

    if ($Backfill -or $sorted.Count -le $Max) {
        return [PSCustomObject]@{
            Selected  = $sorted
            Remaining = 0
        }
    }

    $selected = @($sorted | Select-Object -First $Max)
    $remaining = $sorted.Count - $Max

    return [PSCustomObject]@{
        Selected  = $selected
        Remaining = $remaining
    }
}

# ------------------------------------------------------------
# Test-InlineAttachment
#   添付ファイルが「インライン画像（署名画像等）」であり除外すべきかどうかを判定する。
#   COM プロパティの取得自体は呼び出し側（本体）が try/catch で行い、
#   取得できなかった値は $null としてこの関数に渡す（＝ここでは常にCOM非依存）。
#
#   判定1: $HiddenProp（PR_ATTACHMENT_HIDDEN の値）が $true → 除外
#   判定2: $ContentId（PR_ATTACH_CONTENT_ID）が取得できており、
#          $HtmlBody 内に "cid:<ContentId>" への参照があれば → 除外
#   どちらも判定不能（$null 等）なら安全側に倒して除外しない（$false）。
# ------------------------------------------------------------
function Test-InlineAttachment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        $HiddenProp,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$ContentId,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$HtmlBody
    )

    # 判定1: PR_ATTACHMENT_HIDDEN
    if ($HiddenProp -eq $true) {
        return $true
    }

    # 判定2: PR_ATTACH_CONTENT_ID を HTMLBody 内の cid: 参照と突き合わせる
    if (-not [string]::IsNullOrEmpty($ContentId)) {
        $body = $HtmlBody
        if ($null -eq $body) { $body = "" }

        $needle = "cid:" + $ContentId
        if ($body.IndexOf($needle, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
            return $true
        }
    }

    # 判定不能、または該当なし → 安全側（保存する＝除外しない）
    return $false
}

# ------------------------------------------------------------
# 書き込み系統の集約
#   フォルダ作成／添付保存／processed-ids追記／(mail-watch追加分)イベント記録追記
#   の「実際に書き込む」処理をここに集約し、いずれも -DryRun 指定時はファイル
#   システムへ一切書き込まず計画情報だけを返す構造にする。本体はこれらの関数を
#   呼ぶだけにし、DryRunかどうかの分岐をあちこちに書かない。
#
#   Save-AttachmentToTarget だけは実際の保存という性質上「添付の中身をどこかに
#   書き出す処理」が要るが、Outlook COMを直接参照させないために -WriteTempFile に
#   スクリプトブロックとして注入してもらう方式にした（本体側は { param($tmpPath)
#   $att.SaveAsFile($tmpPath) } を渡す。テスト側はOutlook非依存の偽の書き込み
#   （Set-Content等）を渡せる）。これによりこのファイル自体はOutlook非依存のまま保てる。
# ------------------------------------------------------------

# ------------------------------------------------------------
# Initialize-TargetFolder
#   保存先フォルダの有無を確認し、無ければ作成する（フォルダ作成の書き込み系統）。
#   -DryRun 指定時は作成せず、既存有無だけを返す。
#   戻り値: PSCustomObject @{ Existed = <bool>; Created = <bool> }
# ------------------------------------------------------------
function Initialize-TargetFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [switch]$DryRun
    )

    $existed = Test-Path -LiteralPath $Path -PathType Container

    if ($DryRun) {
        return [PSCustomObject]@{ Existed = $existed; Created = $false }
    }

    if (-not $existed) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
        return [PSCustomObject]@{ Existed = $false; Created = $true }
    }

    return [PSCustomObject]@{ Existed = $true; Created = $false }
}

# ------------------------------------------------------------
# Add-ProcessedId
#   処理済みメールのEntryIDを processed-ids.txt に追記する（processed-ids追記の書き込み系統）。
#   -DryRun 指定時は追記せず、Recorded=$false を返すだけにする。
#   戻り値: PSCustomObject @{ Recorded = <bool> }
# ------------------------------------------------------------
function Add-ProcessedId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$EntryId,

        [switch]$DryRun
    )

    if ($DryRun) {
        return [PSCustomObject]@{ Recorded = $false }
    }

    Add-Content -LiteralPath $Path -Value $EntryId -Encoding UTF8
    return [PSCustomObject]@{ Recorded = $true }
}

# ------------------------------------------------------------
# Save-AttachmentToTarget
#   添付ファイルの保存（添付保存の書き込み系統）。
#
#   本実行時は次の流れで「実サイズ・実ハッシュ」に基づいて判定する:
#     1. $WriteTempFile スクリプトブロックで $TempFolder 配下の一時ファイルへ書き出す
#        （$TempFolder はローカルの一時フォルダを渡す想定。
#          Move-Itemがボリューム跨ぎで実質コピーになっても問題ない）
#     2. 書き出された一時ファイルの実際の Length と MD5 ハッシュを計算する
#     3. Get-SaveFileName にその実サイズ・実ハッシュを渡して判定する
#        （同名同実サイズ同ハッシュ→一時ファイル破棄でスキップ／
#          同名別実サイズ・別ハッシュ→連番名でMove-Item／未存在→そのままMove-Item）
#     4. Move-Item に失敗した場合は例外をそのまま投げる
#        （呼び出し側のメール単位try/catchで「そのメールを失敗扱い」にし、
#          processed-idsに記録しない＝次回自動リトライさせるため、ここでは握りつぶさない）
#
#   -DryRun 指定時は一時ファイルすら作らない。SaveAsFileできないため、
#   $EstimatedSize（att.Sizeなど近似値）と既存ファイルの実Lengthとの比較のみによる
#   概算判定にとどめる（MAPIのPR_ATTACH_SIZEは実サイズと厳密一致しない場合があるため、
#   ここでの判定はあくまで概算であり、本実行時の実サイズ・実ハッシュ判定より精度が低いことに注意）。
#
#   戻り値: PSCustomObject @{ Action = <string>; SavedPath = <string|$null>; Truncated = <bool>; MatchedName = <string|$null> }
#     Action: "Saved" | "Skip" | "SaveEstimate"(DryRunの保存予定) | "SkipEstimate"(DryRunのスキップ見込み)
# ------------------------------------------------------------
function Save-AttachmentToTarget {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetFolder,

        [Parameter(Mandatory = $true)]
        [string]$SafeFileName,

        [Parameter(Mandatory = $true)]
        [string]$TempFolder,

        [Parameter(Mandatory = $true)]
        [scriptblock]$WriteTempFile,

        [Parameter(Mandatory = $false)]
        [long]$EstimatedSize = 0,

        [switch]$DryRun
    )

    if ($DryRun) {
        # 実ファイル操作は一切行わない（一時ファイルすら作らない・$WriteTempFileも呼ばない）。
        $fullPath = Join-Path $TargetFolder $SafeFileName
        $approxSkip = $false
        try {
            if (Test-Path -LiteralPath $fullPath -PathType Leaf) {
                $existingLen = (Get-Item -LiteralPath $fullPath).Length
                if ($existingLen -eq $EstimatedSize) { $approxSkip = $true }
            }
        }
        catch {
            $approxSkip = $false
        }

        return [PSCustomObject]@{
            Action      = if ($approxSkip) { "SkipEstimate" } else { "SaveEstimate" }
            SavedPath   = $null
            Truncated   = $false
            MatchedName = $null
        }
    }

    if (-not (Test-Path -LiteralPath $TempFolder)) {
        New-Item -ItemType Directory -Path $TempFolder -Force | Out-Null
    }

    $tmpName = [guid]::NewGuid().ToString("N") + "_" + $SafeFileName
    $tmpPath = Join-Path $TempFolder $tmpName

    & $WriteTempFile $tmpPath

    $actualSize = (Get-Item -LiteralPath $tmpPath).Length
    $actualHash = (Get-FileHash -LiteralPath $tmpPath -Algorithm MD5).Hash

    $matchedNameRef = [ref]$null
    $saveName = Get-SaveFileName -Folder $TargetFolder -FileName $SafeFileName -ActualSize $actualSize -ActualHash $actualHash -MatchedName $matchedNameRef

    if ($null -eq $saveName) {
        Remove-Item -LiteralPath $tmpPath -Force -ErrorAction SilentlyContinue
        return [PSCustomObject]@{
            Action      = "Skip"
            SavedPath   = $null
            Truncated   = $false
            MatchedName = $matchedNameRef.Value
        }
    }

    $rawFullPath = Join-Path $TargetFolder $saveName
    $limited = Limit-PathLength -FullPath $rawFullPath -Max 250

    try {
        Move-Item -LiteralPath $tmpPath -Destination $limited.Path -Force
    }
    catch {
        throw ("添付ファイルの移動に失敗しました ({0}): {1}" -f $SafeFileName, $_.Exception.Message)
    }

    return [PSCustomObject]@{
        Action      = "Saved"
        SavedPath   = $limited.Path
        Truncated   = $limited.Truncated
        MatchedName = $null
    }
}

# ------------------------------------------------------------
# Show-SaveNotification
#   保存完了ポップアップの「出すか出さないか」の判定ロジックだけを持つ。
#   メッセージ文言はツールごとに違うため、呼び出し側が -Message で組み立てて渡す。
#   実際にポップアップを表示する処理（Outlook非依存にできないCOM呼び出し）は
#   -ShowPopup スクリプトブロックとして注入させる。これにより判定ロジック自体は
#   Outlook無しでテストでき、ShowPopup側はテストでモックに差し替えられる。
#
#   戻り値: PSCustomObject @{ Action = <string>; Message = <string|$null>; Error = <string|$null> }
#     Action: "Shown"（表示を試みた=ShowPopupを呼んだ) | "Failed"（ShowPopupが例外を投げた) |
#             "Skipped"（DryRun／$EnableNotification=$false／保存0件のいずれかで非表示）
# ------------------------------------------------------------
function Show-SaveNotification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$SavedFileCount,

        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $true)]
        [bool]$EnableNotification,

        [Parameter(Mandatory = $true)]
        [scriptblock]$ShowPopup,

        [switch]$DryRun
    )

    if ($DryRun -or -not $EnableNotification -or $SavedFileCount -le 0) {
        return [PSCustomObject]@{
            Action  = "Skipped"
            Message = $null
            Error   = $null
        }
    }

    try {
        & $ShowPopup $Message
        return [PSCustomObject]@{
            Action  = "Shown"
            Message = $Message
            Error   = $null
        }
    }
    catch {
        return [PSCustomObject]@{
            Action  = "Failed"
            Message = $Message
            Error   = $_.Exception.Message
        }
    }
}

# ------------------------------------------------------------
# Format-TruncatedList
#   一覧を先頭N件＋「他n件」に要約する汎用ヘルパー。
#   $Items.Count が $Threshold 以下ならそのまま返す。超える場合は
#   先頭 $ShowCount 件 + $OthersFormat（{0}に残数が入る）を1件追加して返す。
# ------------------------------------------------------------
function Format-TruncatedList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string[]]$Items,

        [Parameter(Mandatory = $false)]
        [int]$Threshold = 6,

        [Parameter(Mandatory = $false)]
        [int]$ShowCount = 5,

        [Parameter(Mandatory = $false)]
        [string]$OthersFormat = "他 {0} 件"
    )

    # 注意: PowerShellでは @($null) は「$null が1個入った配列」(Count=1) になり、
    # 空配列にはならない（実測確認済み）。$null は明示的に空配列として扱う。
    if ($null -eq $Items) {
        $list = @()
    }
    else {
        $list = @($Items)
    }

    if ($list.Count -le $Threshold) {
        return $list
    }

    $shown = @($list | Select-Object -First $ShowCount)
    $shown += ($OthersFormat -f ($list.Count - $ShowCount))
    return $shown
}

# ==============================================================
# ここから下は mail-watch 固有の追加関数。
# mail-attach-saver 側の MailAttachLib.ps1 には存在しない（差分があって正常）。
# ==============================================================

# ------------------------------------------------------------
# Test-SenderMatch
#   差出人メールアドレスが監視対象ドメインのいずれかに一致するかを判定する。
#   - $SenderEmailAddress は既に解決済みの生SMTPアドレス文字列を渡すこと
#     （Exchange DN形式("/O="始まり)の解決自体はCOM依存のため本体側の責務。
#      解決できなかった場合は $null または未解決の生文字列を渡せばよく、
#      いずれの場合もこの関数は「一致しない」を返す）
#   - 小文字化したうえで、各ドメインについて "@"+ドメイン または "."+ドメイン
#     で終わるか（後方一致）を見る。"."を使うことで "mail.example.co.jp" が
#     "example.co.jp" 指定でもサブドメインとして一致する一方、
#     "notexample.co.jp" のような紛らわしい別ドメインを誤って一致させない
#     （区切り文字を挟んだ後方一致であることが安全性の要）。
# ------------------------------------------------------------
function Test-SenderMatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$SenderEmailAddress,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string[]]$AllowedDomains
    )

    if ([string]::IsNullOrWhiteSpace($SenderEmailAddress)) { return $false }
    if ($null -eq $AllowedDomains) { return $false }

    $addr = $SenderEmailAddress.ToLowerInvariant()

    foreach ($domain in $AllowedDomains) {
        if ([string]::IsNullOrWhiteSpace($domain)) { continue }

        $d = $domain.Trim().ToLowerInvariant().TrimStart('@')
        if ([string]::IsNullOrEmpty($d)) { continue }

        if ($addr.EndsWith("@" + $d, [System.StringComparison]::Ordinal) -or `
            $addr.EndsWith("." + $d, [System.StringComparison]::Ordinal)) {
            return $true
        }
    }

    return $false
}

# ------------------------------------------------------------
# Get-EventFolderName
#   受信メール1通ぶんの保存先フォルダ名を決定する。
#   命名規則: "yyyy-MM-dd_HHmm_<サニタイズ済み件名先頭40字>"
#   $ParentPath 配下に同名フォルダが既にあれば " (2)"..." (3)"... と連番を振る
#   （実フォルダの有無を見るだけの読み取り専用。作成はしない＝Initialize-TargetFolderの役目）。
# ------------------------------------------------------------
function Get-EventFolderName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ParentPath,

        [Parameter(Mandatory = $true)]
        [datetime]$ReceivedTime,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Subject
    )

    $safeSubject = ConvertTo-SafeName -Name $Subject
    if ($safeSubject.Length -gt 40) {
        $safeSubject = $safeSubject.Substring(0, 40)
    }
    # 40字カット後に末尾がドット/空白になったり空になったりし得るため再度整える
    $safeSubject = ConvertTo-SafeName -Name $safeSubject

    $baseName = "{0}_{1}" -f $ReceivedTime.ToString("yyyy-MM-dd_HHmm"), $safeSubject

    $candidateName = $baseName
    $n = 2
    while (Test-Path -LiteralPath (Join-Path $ParentPath $candidateName)) {
        $candidateName = "{0} ({1})" -f $baseName, $n
        $n++
    }

    return $candidateName
}

# ------------------------------------------------------------
# New-MailEventLine
#   着信記録(JSONL)の1行分を組み立てる。ConvertTo-Json -Compress は既定で
#   非ASCII文字を \uXXXX にエスケープするため、Claude・人間どちらが直接
#   ファイルを開いても読みやすいよう日本語部分は実際のUTF-8文字へ戻して返す
#   （\uXXXX のままでもJSONとしては正しく壊れないが、可読性のための後処理）。
#   戻り値: JSONL用の1行分の文字列（末尾に改行は含まない）
# ------------------------------------------------------------
function New-MailEventLine {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [datetime]$ReceivedTime,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$FromDomain,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$FromName,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Subject,

        [Parameter(Mandatory = $true)]
        [string]$Folder,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string[]]$Files,

        [Parameter(Mandatory = $true)]
        [int]$SavedCount,

        [Parameter(Mandatory = $false)]
        [datetime]$Timestamp = (Get-Date)
    )

    if ($null -eq $FromName) { $FromName = "" }
    if ($null -eq $Files) { $Files = @() }

    $record = [ordered]@{
        receivedTime = $ReceivedTime.ToString("o")
        fromDomain   = $FromDomain
        fromName     = $FromName
        subject      = $Subject
        folder       = $Folder
        files        = @($Files)
        savedCount   = $SavedCount
        ts           = $Timestamp.ToString("o")
    }

    $json = $record | ConvertTo-Json -Compress

    # \uXXXX エスケープを実文字へ戻す（JSONとしての正しさには影響しない可読性のための処理）
    $json = [regex]::Replace($json, '\\u([0-9a-fA-F]{4})', {
        param($m)
        [string][char]([Convert]::ToInt32($m.Groups[1].Value, 16))
    })

    return $json
}

# ------------------------------------------------------------
# Add-MailEventRecord
#   New-MailEventLine が組み立てたJSON行を $Path（着信記録JSONL）へ追記する
#   （イベント記録追記の書き込み系統）。
#   仕様により本ファイルはBOM無しUTF-8で書く（.ps1のBOM付き要件とは別）。
#   .NET の UTF8Encoding($false) は「BOMを出さない」設定であり、新規作成時も
#   追記時もBOMを一切書き出さないことを利用する。
#   -DryRun 指定時は追記せず、Recorded=$false を返すだけにする。
#   戻り値: PSCustomObject @{ Recorded = <bool> }
# ------------------------------------------------------------
function Add-MailEventRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$JsonLine,

        [switch]$DryRun
    )

    if ($DryRun) {
        return [PSCustomObject]@{ Recorded = $false }
    }

    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    $noBomUtf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::AppendAllText($Path, $JsonLine + "`r`n", $noBomUtf8)

    return [PSCustomObject]@{ Recorded = $true }
}
