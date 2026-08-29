"""
engines/rules.py — v0実装（規則ベース＝正規表現エンジン）。

fields.yaml の regex_hints を優先順位どおりに適用し、
案件内の書類テキスト（要OCR・エラーの書類は除外済み）から
候補を1件ずつ拾う、もっとも単純な実装。

探索順:
  1. field.source_docs に一致する doc_kind の書類を、ファイル名の昇順で先に探す
  2. 見つからなければ、それ以外の doc_kind の書類にもフォールバックする
     （フォールバックで見つかった場合は confidence="low"）
  3. 同一書類内では行の出現順（上から下）に探す
  4. regex_hints は上から順に試し、最初にヒットしたパターン・行を採用する
"""
from __future__ import annotations

import re

from .base import Candidate, DocText, ExtractionEngine, FieldSpec


class RuleBasedEngine(ExtractionEngine):
    """v0: fields.yaml の regex_hints による規則ベース抽出エンジン。"""

    def extract_fields(
        self, case_id: str, docs: list[DocText], fields: list[FieldSpec]
    ) -> list[Candidate]:
        usable_docs = [d for d in docs if not d.needs_ocr and d.error is None]
        candidates: list[Candidate] = []
        for f in fields:
            cand = self._find_for_field(f, usable_docs)
            if cand is not None:
                candidates.append(cand)
        return candidates

    @staticmethod
    def _order_docs(f: FieldSpec, docs: list[DocText]) -> list[DocText]:
        preferred = [d for d in docs if d.doc_kind in f.source_docs]
        others = [d for d in docs if d.doc_kind not in f.source_docs]
        preferred.sort(key=lambda d: d.file_name)
        others.sort(key=lambda d: d.file_name)
        return preferred + others

    def _find_for_field(self, f: FieldSpec, docs: list[DocText]) -> Candidate | None:
        if not f.v0_target or not f.regex_hints:
            return None

        ordered_docs = self._order_docs(f, docs)

        for pattern_index, pattern in enumerate(f.regex_hints):
            try:
                rx = re.compile(pattern, re.IGNORECASE)
            except re.error:
                # 壊れた正規表現はスキップ（fields.yaml側の記述ミス）。
                # extract.py 側のエラーサマリに件数として出す。
                continue

            for doc in ordered_docs:
                for line in doc.lines:
                    m = rx.search(line.text)
                    if not m:
                        continue
                    value = m.groupdict().get("value")
                    if value is None:
                        value = m.group(0)
                    value = value.strip()
                    if not value:
                        continue
                    confidence = "high" if doc.doc_kind in f.source_docs else "low"
                    return Candidate(
                        field_id=f.id,
                        value=value,
                        source_file=doc.file_name,
                        source_location=line.location,
                        match_start=m.start(),
                        match_end=m.end(),
                        pattern_index=pattern_index,
                        matched_doc_kind=doc.doc_kind,
                        confidence=confidence,
                    )
        return None
