# mail-watch — 法人案件Aメール個人用ウォッチャー

開発者本人のPCで使う個人ツール。**クライアント納品物ではない。**
特定クライアント（本ドキュメント内では呼称「法人案件A」を使用。実社名・実ドメインは
`config.local.ps1` にのみ本人が設定し、このリポジトリには書かない）からの
添付付きメールが届いたら、添付をローカル保存し、ポップアップで気づけるようにし、
Claudeが後で読める着信記録（JSONL）を残す。

想定運用: メールが届く →（当面は手動で）本ツールを実行 → ポップアップで気づく →
Claudeに「メール確認」と言う → Claudeが `logs\inbox-events.jsonl` を読んで
次のたたき台準備に入る。

## 概要

- Outlookは「起動中のインスタンス」にのみ接続する（未起動なら何もせず終了）
- 受信トレイ直近N日分を対象に、**差出人ドメインが一致し、かつ保存対象添付が1件以上ある**
  メールだけを検知する（挨拶メール等の添付なしメールは無視）
- 該当メールの添付を `received\YYYY-MM-DD_HHmm_<件名先頭40字>\` へ保存する
- 保存したら着信記録を `logs\inbox-events.jsonl` に1メール1行で追記する
- 保存があった実行のみ、最前面表示・手動で閉じるまで消えないポップアップで知らせる
- 当面は手動実行。タスクスケジューラ登録スクリプトは同梱するが、
  セキュリティ確認ダイアログの問題（後述）が解消するまでは登録しないこと

## ファイル構成

```
mail-watch/
├── Watch-ClientMail.ps1        本体
├── MailAttachLib.ps1           純粋ロジック関数群（mail-attach-saverからの複製）
├── config.sample.ps1           設定サンプル（全項目コメント付き）
├── config.local.ps1            実際に使う設定の正本（Git管理対象外・プレースホルダで管理）
├── Register-Watch-Task.ps1     タスクスケジューラへの登録（5分毎。実行はまだしないこと）
├── Unregister-Watch-Task.ps1   タスクスケジューラからの登録解除
├── tests/
│   └── Run-Tests.ps1           Outlook不要の合成テスト
├── received/                   保存された添付（Git管理対象外）
├── logs/                       実行ログ・processed-ids.txt・inbox-events.jsonl（Git管理対象外）
└── .gitignore
```

`MailAttachLib.ps1` は `..\mail-attach-saver\MailAttachLib.ps1`（クライアント納品キット側）
の複製。**正本は mail-attach-saver 側**。共通関数（Get-CodeFromSubject〜
Show-SaveNotification・Format-TruncatedList）を改修する場合は両方へ反映すること。
`Test-SenderMatch`・`Get-EventFolderName`・`New-MailEventLine`・`Add-MailEventRecord`
の4関数はmail-watch専用の追加分で、mail-attach-saver側には存在しない（差分があって正常）。

## 設定一覧（config.local.ps1 / config.sample.ps1）

| 設定名 | 既定値 | 内容 |
|--------|--------|------|
| `$WatchSenderDomains` | プレースホルダ | 監視対象の差出人ドメイン（"@"より後ろ・配列で複数可） |
| `$RequireAttachment` | `$true` | 保存対象添付が1件以上あることを候補条件にするか |
| `$RetentionDays` | 7 | 受信トレイの遡及日数 |
| `$MaxMailsPerRun` | 20 | 1回の実行で処理する候補メール数の上限 |
| `$EnableNotification` | `$true` | 保存完了時のポップアップ通知を出すか |
| `$SaveRoot` | `<このフォルダ>\received` | 添付の保存先ルート |
| `$EventLogPath` | `logs\inbox-events.jsonl` | 着信記録(JSONL)の保存先 |
| `$ProcessedIdsPath` | `logs\processed-ids.txt` | 処理済みEntryIDの記録先 |
| `$LogFolder` | `logs\` | 実行ログの保存先フォルダ |

### 使い方（最短ルート）

1. `config.sample.ps1` を見ながら `config.local.ps1` の **`$WatchSenderDomains` の1行だけ**
   実際のドメインに書き換える（例: `$WatchSenderDomains = @("example.co.jp")`）
2. まず安全確認としてDryRunを実行する:
   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Watch-ClientMail.ps1 -DryRun
   ```
3. 問題なければ通常実行する:
   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Watch-ClientMail.ps1
   ```
4. ポップアップが出たら、Claudeに「メール確認」と伝える
   （Claude側は `logs\inbox-events.jsonl` の末尾（前回チェック以降）を読んで、
   件名・保存先フォルダ・ファイル一覧から次の作業に入れる）

## 着信記録(JSONL)の見方

`logs\inbox-events.jsonl` は1行1メールのJSON Lines形式。例:

```json
{"receivedTime":"2026-08-29T18:52:52+09:00","fromDomain":"example.co.jp","fromName":"山田太郎","subject":"至急のご確認をお願いいたします","folder":"received\\2026-08-29_1852_至急のご確認をお願いいたします","files":["invoice.pdf"],"savedCount":1,"ts":"2026-08-29T18:53:01+09:00"}
```

- `receivedTime`: メールの受信日時（ISO8601）
- `fromDomain` / `fromName`: 差出人のドメイン・表示名
- `subject`: 件名
- `folder`: 保存先フォルダの相対パス（`$PSScriptRoot` 基準）
- `files`: 今回新規に保存したファイル名の一覧
- `savedCount`: 今回新規に保存した件数
- `ts`: この行を記録した時刻（`receivedTime`とは別。実行のタイミング）

新規保存が0件（全て冪等スキップ）だったメールは記録しない（Claudeに伝えるべき新しい
情報が無いため）。DryRun中は記録しない。

**【重要・エンコーディングの注意】** このファイルは仕様によりBOM無しUTF-8で書かれる。
Windows PowerShellで直接読む場合は `Get-Content -Encoding UTF8` のように
**必ず -Encoding UTF8 を明示する**こと。指定を省略すると、BOMが無いことで
Windows PowerShell 5.1がシステム既定のコードページ（日本語環境ではCP932）と誤認し、
日本語部分が文字化けする（開発時に実際に確認済みの不具合パターン。`.ps1`をBOM付きに
する必要がある理由と表裏の関係）。Claudeが直接ファイルを読む分にはこの問題は起きない。

## セキュリティ確認ダイアログについて（常駐化前に必ず確認）

Outlookには「プログラムによるアクセスの保護」（Object Model Guard）という標準機能があり、
ウイルス対策ソフトが有効と認識されていないPCでは、本ツールの実行のたびに
『Outlook内に保存されている電子メールアドレス情報が、以下のプログラムによって
アクセスされようとしています。』という警告ダイアログが表示され、誰かが手動で
「許可する」を押すまで処理が止まる（姉妹ツールmail-attach-saverの実機リハーサルで
判明した既知の挙動）。

- **当面（手動実行）は問題ない**: ダイアログが出たらその場で許可すればよい
- **タスクスケジューラに登録して常駐化する場合は非実用**: 5分毎に無人でダイアログが
  出ては止まるだけになる。登録前に、PCのウイルス対策ソフトの状態を正常化し
  （Windowsセキュリティで警告が出ていない状態にする）、ダイアログが出なくなることを
  手動実行で確認してから `Register-Watch-Task.ps1` を実行すること

## 再発防止・設計上の工夫（姉妹ツールmail-attach-saverから踏襲）

| 工夫 | 内容 |
|------|------|
| 処理件数上限（候補にのみ適用） | ドメイン不一致・添付なしのメールに上限を先食いされない。上限は「ドメイン一致かつ添付あり」の候補にのみ適用する |
| 冪等スキップ（実サイズ・ハッシュ一致） | 添付を一時ファイルへ書き出した実際のサイズ・MD5ハッシュで既存ファイルと比較するため、同じメールに何度実行しても重複ファイルが増えない |
| `-DryRun` | 実際の保存・フォルダ作成・記録・通知を一切行わず、対象件数を事前に確認できる |
| メール本体の不可侵 | `UnRead`への代入・`.Save()`・移動・削除を一切行わない |
| 単一起動の保証（ミューテックス） | 多重起動で処理が重複するのを防ぐ（kitとは別名のミューテックスを使用） |
| メール1通単位のエラー処理 | 1通の処理に失敗しても他のメールの処理を止めず、失敗したメールだけ次回自動リトライする |
| 通知の見落とし対策 | 表示直前に必ずログを残し、システムモーダル表示にする |
| Outlook接続の3段構え（`Connect-Outlook`） | 実機で「Outlookが起動・応答しているのにROT未登録で`GetActiveObject`が失敗する」現象が確認されたため、失敗時はOUTLOOK.EXEプロセスの存在を確認した上でのみ相乗り接続（`New-Object`）を試す。プロセスが無い状態でこの経路に入ると誤ってOutlookを起動してしまうため、プロセスチェックを必須のゲートにしている（mail-attach-saverと同一実装） |

## トラブルシュート: Outlookが起動しているのに「起動していない」と出る／接続できない

実機でOutlookが起動・応答・受信トレイ表示中でもROT（Running Object Table）への登録が
外れる現象が確認されている（メインウィンドウを閉じて通知領域の常駐だけ残した後に
再表示した等の経緯で発生）。本ツールは自動的に「相乗り接続」（起動中のOUTLOOK.EXE
プロセスに対して別経路で接続し直す）を試みるため、ログに
`Outlook接続: 相乗り(ROT未登録を回避)` と出ていれば自動的に回避できている（対処不要）。
それでも `Outlookは起動していますが接続できませんでした（Outlook再起動で解消する
見込み）` が出る場合は、Outlookを一度完全に終了してから再起動すること。

## 開発・検証について

この開発機にはOutlookプロファイルが無いため、`Watch-ClientMail.ps1` 本体・
`Register-Watch-Task.ps1`・`Unregister-Watch-Task.ps1` は実機での動作確認ができない。

- 開発機で実施済み: `MailAttachLib.ps1` の純粋ロジックに対する合成テスト
  （`tests\Run-Tests.ps1`、全PASS）、全 `.ps1` の静的構文チェック
- 開発機では実施できない（本人のPCでのみ可能）: Outlook COM接続・実際の添付保存・
  タスクスケジューラ登録・ポップアップ通知の実機確認

## ライブテスト時に確認してほしいポイント

（詳細は今回の報告メッセージ内の「ライブテスト時に確認すべきポイント一覧」を参照）

## 保守者向け注意

`MailAttachLib.ps1` の共通部分（Get-CodeFromSubject〜Format-TruncatedList）を
mail-attach-saver側で改修した場合は、このファイルにも同じ変更を反映すること。
逆にmail-watch固有の4関数（Test-SenderMatch等）は複製先には含めなくてよい。
