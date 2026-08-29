"""
共通ヘルパー群。

含むもの:
  - load_config / load_fields : config.json(+config.local.json) と fields.yaml の読込
  - build_doc_text            : 1書類(PDF/xlsx)からテキスト行を抽出（要OCR判定つき）
  - classify_doc_kind         : 書類種の自動分類（ファイル名手がかり＋内容キーワード）
  - is_answer_key             : 人手完成形xlsx（評価管理表）かどうかの判定
  - normalize_value           : 突合用の値正規化（全角/半角・空白吸収）
  - write_csv                 : CSV書き出し（fm_import.csv はBOM付きUTF-8で使う）

★機密上の注意★
  ここで実際の書類テキストを扱う関数（build_doc_text 等）は、抽出した値を
  「ファイルに書く」ことは想定しているが、呼び出し側（extract.py/compare.py）が
  デバッグ目的で標準出力に流してよいのは件数・文字数・真偽値・型だけ。
  実際の抽出値・書類本文を print してはならない（README/プロジェクト規約）。
  例外発生時も、生の例外メッセージ（ファイルパスを含み得る）はそのまま
  ログ・print しない。例外クラス名だけを使うこと（sanitize_error参照）。
"""
from __future__ import annotations

import csv
import json
import os
import re
import unicodedata
from pathlib import Path

from engines.base import DocLine, DocText, FieldSpec


class ConfigError(RuntimeError):
    pass


# ---------------------------------------------------------------------------
# config / fields.yaml 読込
# ---------------------------------------------------------------------------

def load_config(project_dir: Path) -> dict:
    """config.json（安全なテンプレート）+ config.local.json（実パス・gitignore対象）
    + 環境変数 TSUKAN_CASES_ROOT を合成して設定を返す。"""
    config_path = project_dir / "config.json"
    with open(config_path, encoding="utf-8") as f:
        cfg = json.load(f)

    local_path = project_dir / "config.local.json"
    case_map: dict[str, str] = {}
    cases_root = os.environ.get("TSUKAN_CASES_ROOT")

    if local_path.exists():
        with open(local_path, encoding="utf-8") as f:
            local_cfg = json.load(f)
        if not cases_root:
            cases_root = local_cfg.get("cases_root")
        case_map = local_cfg.get("case_map", {}) or {}

    if not cases_root:
        raise ConfigError(
            "cases_root が未設定です。config.local.json を作成する"
            "（ひな形: config.local.example.json をコピーして実値を入れる）か、"
            "環境変数 TSUKAN_CASES_ROOT を設定してから実行してください。"
        )
    if not case_map:
        raise ConfigError(
            "case_map が空です。config.local.json に案件ID→実フォルダ名の対応を"
            "設定してください（例: {\"case01\": \"実フォルダ名\"}）。"
        )

    cfg["cases_root"] = cases_root
    cfg["case_map"] = case_map
    return cfg


def load_fields(project_dir: Path) -> tuple[list[FieldSpec], list[str]]:
    """fields.yaml を読み込み、(FieldSpecのリスト, doc_kindsのリスト) を返す。"""
    import yaml

    with open(project_dir / "fields.yaml", encoding="utf-8") as f:
        data = yaml.safe_load(f)

    doc_kinds = list(data.get("doc_kinds", []))
    fields: list[FieldSpec] = []
    for item in data["fields"]:
        fields.append(
            FieldSpec(
                id=item["id"],
                label_ja=item.get("label_ja", item["id"]),
                ida_item_no=str(item.get("ida_item_no", "")),
                official_symbol=str(item.get("official_symbol", "")),
                type=str(item.get("type", "")),
                source_docs=list(item.get("source_docs", []) or []),
                v0_target=bool(item.get("v0_target", False)),
                regex_hints=list(item.get("regex_hints", []) or []),
                notes=str(item.get("notes", "") or ""),
            )
        )
    return fields, doc_kinds


# ---------------------------------------------------------------------------
# 書類の種別判定
# ---------------------------------------------------------------------------

_FILENAME_STRONG = {
    "shipping_doc": re.compile(r"出荷書類"),
    "inspection_cert": re.compile(r"^INS[\s〇]", re.IGNORECASE),
}
_FILENAME_WEAK = {
    "invoice_packing_list": re.compile(r"inv[,\s._-]*pl|packing[\s_-]*list|\binv\b|invoice", re.IGNORECASE),
    "evaluation_doc": re.compile(r"評価書類|COMMISSION|^EDWSC-?\d|^EDW\d", re.IGNORECASE),
}
_CONTENT_KEYWORDS = {
    "invoice_packing_list": ["INVOICE", "COMMERCIAL INVOICE", "PACKING LIST", "P/L", "SELLER", "BUYER", "NET WEIGHT"],
    "shipping_doc": ["BILL OF LADING", "B/L", "ARRIVAL NOTICE", "SHIPPER", "CONSIGNEE", "VESSEL", "FREIGHT", "NOTIFY PARTY"],
    "evaluation_doc": ["COMMISSION", "評価"],
    "inspection_cert": ["INSPECTION", "INSURANCE", "CERTIFICATE"],
}


def classify_doc_kind(file_name: str, lines: list[DocLine]) -> str:
    """ファイル名の強い手がかり → 内容キーワードのスコアリング → 弱い手がかり、の順で判定。
    document_classification.yaml の「2件以上ヒットで確信」方式を踏襲。"""
    for kind, pat in _FILENAME_STRONG.items():
        if pat.search(file_name):
            return kind

    if lines:
        blob = "\n".join(l.text for l in lines).upper()
        scores = {
            kind: sum(blob.count(kw.upper()) for kw in kws)
            for kind, kws in _CONTENT_KEYWORDS.items()
        }
        best_kind, best_score = max(scores.items(), key=lambda kv: kv[1])
        if best_score >= 2:
            return best_kind

    for kind, pat in _FILENAME_WEAK.items():
        if pat.search(file_name):
            return kind

    return "unclassified"


def is_answer_key(file_name: str, pattern: str) -> bool:
    return file_name.lower().endswith(".xlsx") and re.search(pattern, file_name) is not None


# ---------------------------------------------------------------------------
# エラーの安全なログ表現（生の例外メッセージはファイルパスを含み得るため使わない）
# ---------------------------------------------------------------------------

def sanitize_error(exc: BaseException) -> str:
    return type(exc).__name__


# ---------------------------------------------------------------------------
# PDF / xlsx テキスト抽出
# ---------------------------------------------------------------------------

def _extract_pdf_lines_pdfplumber(path: Path) -> tuple[list[DocLine], int, int]:
    import pdfplumber

    lines: list[DocLine] = []
    total_chars = 0
    with pdfplumber.open(str(path)) as pdf:
        page_count = len(pdf.pages)
        for pno, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            total_chars += len(text.strip())
            for lno, raw_line in enumerate(text.splitlines(), start=1):
                if raw_line.strip():
                    lines.append(DocLine(location=f"p{pno}L{lno}", text=raw_line))
    return lines, total_chars, page_count


def _extract_pdf_lines_pypdf(path: Path) -> tuple[list[DocLine], int, int]:
    from pypdf import PdfReader

    lines: list[DocLine] = []
    total_chars = 0
    reader = PdfReader(str(path))
    page_count = len(reader.pages)
    for pno, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        total_chars += len(text.strip())
        for lno, raw_line in enumerate(text.splitlines(), start=1):
            if raw_line.strip():
                lines.append(DocLine(location=f"p{pno}L{lno}", text=raw_line))
    return lines, total_chars, page_count


def extract_pdf_text(
    path: Path, ocr_min_chars_per_page: int
) -> tuple[list[DocLine], int, bool, str | None]:
    """PDFのテキスト層を抽出する。pdfplumberを優先し、失敗時はpypdfへフォールバック。
    戻り値: (行のリスト, 総文字数, 要OCRフラグ, エラー種別名 or None)"""
    try:
        lines, total_chars, page_count = _extract_pdf_lines_pdfplumber(path)
    except Exception as e1:
        try:
            lines, total_chars, page_count = _extract_pdf_lines_pypdf(path)
        except Exception as e2:
            return [], 0, False, f"{sanitize_error(e1)}/{sanitize_error(e2)}"

    if page_count == 0:
        return lines, total_chars, True, None

    avg_chars_per_page = total_chars / page_count
    needs_ocr = avg_chars_per_page < ocr_min_chars_per_page
    return lines, total_chars, needs_ocr, None


def extract_xlsx_text(path: Path) -> tuple[list[DocLine], int, str | None]:
    """xlsxの各シートを行単位（セルをタブ連結）でテキスト化する。
    戻り値: (行のリスト, 総文字数, エラー種別名 or None)
    既知の限界: 結合セルは左上セル以外は空扱い（openpyxl標準動作のまま）。"""
    import openpyxl

    lines: list[DocLine] = []
    total_chars = 0
    try:
        wb = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
        try:
            for ws in wb.worksheets:
                for rno, row in enumerate(ws.iter_rows(), start=1):
                    cells = [str(c.value) for c in row if c.value is not None]
                    if not cells:
                        continue
                    line_text = "\t".join(cells)
                    total_chars += len(line_text)
                    lines.append(DocLine(location=f"{ws.title}!R{rno}", text=line_text))
        finally:
            wb.close()
    except Exception as e:
        return [], 0, sanitize_error(e)
    return lines, total_chars, None


def build_doc_text(path: Path, ocr_min_chars_per_page: int) -> DocText:
    """1書類ファイルからDocTextを構築する（PDF/xlsxのみ対応。それ以外は unsupported）。"""
    ext = path.suffix.lower()
    if ext == ".pdf":
        lines, char_count, needs_ocr, err = extract_pdf_text(path, ocr_min_chars_per_page)
        doc_kind = "unclassified" if err else classify_doc_kind(path.name, lines)
        return DocText(
            file_name=path.name,
            doc_kind=doc_kind,
            needs_ocr=needs_ocr,
            lines=lines,
            char_count=char_count,
            error=err,
        )
    if ext == ".xlsx":
        lines, char_count, err = extract_xlsx_text(path)
        doc_kind = "unclassified" if err else classify_doc_kind(path.name, lines)
        return DocText(
            file_name=path.name,
            doc_kind=doc_kind,
            needs_ocr=False,
            lines=lines,
            char_count=char_count,
            error=err,
        )
    return DocText(
        file_name=path.name,
        doc_kind="unclassified",
        needs_ocr=False,
        lines=[],
        char_count=0,
        error="UNSUPPORTED_EXT",
    )


# ---------------------------------------------------------------------------
# 値の正規化・CSV出力
# ---------------------------------------------------------------------------

def normalize_value(s) -> str:
    """突合用の正規化: NFKC正規化（全角/半角統一）→ 前後空白除去 → 大文字化。"""
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    return s.strip().upper()


def write_csv(path: Path, header: list[str], rows: list[list], encoding: str = "utf-8") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding=encoding, newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def list_case_files(cases_root: str, real_folder_name: str) -> list[Path]:
    folder = Path(cases_root) / real_folder_name
    return sorted(p for p in folder.iterdir() if p.is_file())
