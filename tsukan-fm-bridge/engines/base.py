"""
エンジン差替口の共通インターフェース定義。

- FieldSpec  : fields.yaml の1フィールド分の仕様
- DocLine    : 1書類中の「1行」（行番号相当の位置情報つきテキスト）
- DocText    : 1書類から抽出した行の並び＋分類結果＋要OCR判定
- Candidate  : 1フィールドについて見つかった抽出候補（根拠つき）
- ExtractionEngine : 候補生成を行うエンジンの共通インターフェース

このファイルはデータ構造とインターフェースのみを定義する。
実際の抽出ロジックは engines/rules.py（v0実装）・
engines/claude_interactive.py（スタブ）側に置く。
"""
from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field as dc_field


@dataclass
class FieldSpec:
    id: str
    label_ja: str
    ida_item_no: str
    official_symbol: str
    type: str
    source_docs: list[str]
    v0_target: bool
    regex_hints: list[str]
    notes: str = ""


@dataclass
class DocLine:
    """1行分のテキストと、その行の位置情報。

    location の書式例:
      PDF  : "p3L12"   (3ページ目・そのページ内12行目)
      xlsx : "SheetName!R5" (シート名・5行目。セルはタブ区切りで連結済み)
    """
    location: str
    text: str


@dataclass
class DocText:
    file_name: str
    doc_kind: str
    needs_ocr: bool
    lines: list[DocLine] = dc_field(default_factory=list)
    char_count: int = 0
    error: str | None = None


@dataclass
class Candidate:
    field_id: str
    value: str
    source_file: str
    source_location: str
    match_start: int
    match_end: int
    pattern_index: int
    matched_doc_kind: str
    confidence: str  # "high"（想定doc_kindで一致）/ "low"（フォールバック一致）


class ExtractionEngine(ABC):
    """フィールド抽出エンジンの共通インターフェース。

    差し替え可能にするため、実装は「案件内の書類テキスト群 (docs) と
    フィールド仕様 (fields) を受け取り、候補のリストを返す」という
    シグネチャだけを守ればよい。内部で regex を使うか、対話セッションに
    委ねるかはエンジン側の自由。
    """

    @abstractmethod
    def extract_fields(
        self, case_id: str, docs: list[DocText], fields: list[FieldSpec]
    ) -> list[Candidate]:
        """docs から fields の候補を抽出して返す。"""
        raise NotImplementedError
