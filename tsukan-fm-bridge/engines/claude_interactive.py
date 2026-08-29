"""
engines/claude_interactive.py — 対話型エンジンの差替口（v0: スタブのみ・未実装）。

機密ポリシー NFR-1-EX（本プロジェクトの厳守事項）:
  次の3条件が「すべて」成立する場合に限り、OCRなどで得た書類テキストを
  対話セッションに渡して抽出候補を生成する、という導線を将来実装してよい。
    1. 書面承諾  … 対象案件について機密情報取り扱いの書面承諾を得ていること
    2. interactive … バッチ/自動実行ではなく、人が同席する対話セッションであること
    3. APIキー不在 … Anthropic API を直接呼び出すコード・APIキーの取り扱いを
                     一切追加しないこと（本プロジェクトは機密ポリシーにより
                     APIキー不使用が確定済み）

本ファイルは v0 時点では上記の導線（コメント・docstring）のみを記述し、
条件判定処理・対話セッションへの受け渡し処理・エンジンとしての実装の
いずれも行わない。Anthropic API を呼ぶコードやAPIキーを扱うコードは
本ファイルは元より本プロジェクト全体に一切含まれない。
"""
from __future__ import annotations

from .base import Candidate, DocText, ExtractionEngine, FieldSpec


class ClaudeInteractiveEngine(ExtractionEngine):
    """NFR-1-EX の3条件が成立した場合にのみ有効化を検討する、v0未実装のスタブ。"""

    def extract_fields(
        self, case_id: str, docs: list[DocText], fields: list[FieldSpec]
    ) -> list[Candidate]:
        raise NotImplementedError(
            "claude_interactive エンジンは v0 では未実装のスタブです。"
            "NFR-1-EX の3条件（書面承諾／interactive／APIキー不在）が"
            "すべて成立した場合に限り、別途実装を検討してください。"
        )
