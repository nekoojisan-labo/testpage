#!/usr/bin/env python
"""
extract.py — たたき台v0: 案件フォルダを走査し、規則ベースで各フィールド候補を抽出する。

使い方:
  python extract.py                 # config.local.json の case_map 全件を処理
  python extract.py --limit 3       # 先頭3案件のみ処理（動作確認用）
  python extract.py --cases case01,case03   # 案件IDを明示指定

出力（すべて out/ 配下・.gitignore対象）:
  out/<case_id>/candidates.csv … 案件ごとの全フィールドの抽出結果
                                  （found=見つかった / not_found=v0対象だが未検出 /
                                    skipped_v0=このデータセットではv0対象外）
  out/fm_import.csv            … 全案件まとめ、FM取込用（UTF-8 BOM・1案件1行・fields.yaml列順）
  out/file_inventory.csv       … 全案件まとめ、書類ごとの分類・要OCR・文字数の一覧（検証用）

★実データの値は標準出力に出さない。件数・文字数・真偽値・型のみ表示する★
（out/ 配下のCSVには実際の抽出値を書く。これはこのツールの成果物そのものであり、
  gitignore対象・元書類のコピーではない＝機密ポリシー3の対象外）
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_DIR))

from engines.rules import RuleBasedEngine  # noqa: E402
from lib.common import (  # noqa: E402
    ConfigError,
    build_doc_text,
    is_answer_key,
    list_case_files,
    load_config,
    load_fields,
    write_csv,
)


def parse_args():
    p = argparse.ArgumentParser(description="案件フォルダから規則ベースでフィールド候補を抽出する")
    p.add_argument("--limit", type=int, default=None, help="先頭N案件のみ処理（動作確認用）")
    p.add_argument(
        "--cases", type=str, default=None,
        help="処理する案件IDをカンマ区切りで指定（例: case01,case03）",
    )
    return p.parse_args()


def select_case_ids(case_map: dict, args) -> list[str]:
    ids = sorted(case_map.keys())
    if args.cases:
        wanted = {c.strip() for c in args.cases.split(",") if c.strip()}
        ids = [c for c in ids if c in wanted]
    if args.limit is not None:
        ids = ids[: args.limit]
    return ids


def process_case(case_id: str, real_folder: str, cfg: dict, fields, engine):
    files = list_case_files(cfg["cases_root"], real_folder)
    answer_key_pattern = cfg["answer_key_filename_pattern"]
    ocr_min = cfg["ocr_min_chars_per_page"]

    doc_texts = []
    answer_key_files = []
    unsupported = []
    inventory_rows = []

    for path in files:
        if is_answer_key(path.name, answer_key_pattern):
            answer_key_files.append(path.name)
            inventory_rows.append([case_id, path.name, path.suffix.lower(), "manual_answer_key", "", "", ""])
            continue
        if path.suffix.lower() not in (".pdf", ".xlsx"):
            unsupported.append(path.name)
            inventory_rows.append([case_id, path.name, path.suffix.lower(), "unsupported", "", "", ""])
            continue

        doc = build_doc_text(path, ocr_min)
        doc_texts.append(doc)
        inventory_rows.append(
            [case_id, doc.file_name, path.suffix.lower(), doc.doc_kind,
             str(doc.needs_ocr), str(doc.char_count), doc.error or ""]
        )

    candidates = engine.extract_fields(case_id, doc_texts, fields)
    candidates_by_field = {c.field_id: c for c in candidates}

    candidate_rows = []
    for f in fields:
        c = candidates_by_field.get(f.id)
        if c is not None:
            candidate_rows.append(
                [f.id, f.label_ja, str(f.v0_target), "found", ";".join(f.source_docs),
                 c.matched_doc_kind, c.confidence, c.source_file, c.source_location,
                 c.match_start, c.match_end, c.pattern_index, c.value]
            )
        else:
            status = "not_found" if f.v0_target else "skipped_v0"
            candidate_rows.append(
                [f.id, f.label_ja, str(f.v0_target), status, ";".join(f.source_docs),
                 "", "", "", "", "", "", "", ""]
            )

    stats = {
        "case_id": case_id,
        "file_count": len(files),
        "answer_key_count": len(answer_key_files),
        "unsupported_count": len(unsupported),
        "doc_count": len(doc_texts),
        "needs_ocr_count": sum(1 for d in doc_texts if d.needs_ocr),
        "error_count": sum(1 for d in doc_texts if d.error),
        "text_ok_count": sum(1 for d in doc_texts if not d.needs_ocr and not d.error),
        "found_count": sum(1 for r in candidate_rows if r[3] == "found"),
        "not_found_count": sum(1 for r in candidate_rows if r[3] == "not_found"),
        "skipped_v0_count": sum(1 for r in candidate_rows if r[3] == "skipped_v0"),
    }
    return candidate_rows, stats, inventory_rows, candidates_by_field


def main():
    args = parse_args()
    try:
        cfg = load_config(PROJECT_DIR)
    except ConfigError as e:
        print(f"[設定エラー] {e}")
        sys.exit(1)

    fields, _doc_kinds = load_fields(PROJECT_DIR)
    engine = RuleBasedEngine()

    case_ids = select_case_ids(cfg["case_map"], args)
    if not case_ids:
        print("処理対象の案件がありません（--cases / --limit / case_map を確認してください）")
        sys.exit(1)

    out_dir = PROJECT_DIR / "out"
    all_inventory_rows = []
    fm_rows = []
    all_stats = []
    run_errors = []

    for case_id in case_ids:
        real_folder = cfg["case_map"][case_id]
        try:
            candidate_rows, stats, inventory_rows, candidates_by_field = process_case(
                case_id, real_folder, cfg, fields, engine
            )
        except Exception as e:
            run_errors.append(f"{case_id}: {type(e).__name__}")
            print(f"[エラー] {case_id} の処理中に {type(e).__name__} が発生しました。スキップします。")
            continue

        write_csv(
            out_dir / case_id / "candidates.csv",
            ["field_id", "label_ja", "v0_target", "status", "doc_kind_expected",
             "matched_doc_kind", "confidence", "source_file", "source_location",
             "match_start", "match_end", "pattern_index", "value"],
            candidate_rows,
        )
        all_inventory_rows.extend(inventory_rows)
        all_stats.append(stats)

        fm_row = [case_id] + [
            (candidates_by_field[f.id].value if f.id in candidates_by_field else "") for f in fields
        ]
        fm_rows.append(fm_row)

        print(
            f"[{case_id}] 書類{stats['file_count']}件 "
            f"(解析対象{stats['doc_count']} / 要OCR{stats['needs_ocr_count']} / "
            f"抽出エラー{stats['error_count']} / 回答キー{stats['answer_key_count']}) "
            f"フィールド found={stats['found_count']} not_found={stats['not_found_count']} "
            f"skipped(v0対象外)={stats['skipped_v0_count']}"
        )

    write_csv(
        out_dir / "file_inventory.csv",
        ["case_id", "file_name", "ext", "doc_kind", "needs_ocr", "char_count", "error"],
        all_inventory_rows,
    )
    write_csv(
        out_dir / cfg["fm_import_csv_name"],
        ["case_id"] + [f.id for f in fields],
        fm_rows,
        encoding=cfg["fm_import_csv_encoding"],
    )

    total_files = sum(s["file_count"] for s in all_stats)
    total_docs = sum(s["doc_count"] for s in all_stats)
    total_ocr = sum(s["needs_ocr_count"] for s in all_stats)
    total_err = sum(s["error_count"] for s in all_stats)
    total_found = sum(s["found_count"] for s in all_stats)
    total_not_found = sum(s["not_found_count"] for s in all_stats)

    print("=" * 60)
    print(
        f"処理案件数: {len(all_stats)} / 総書類数: {total_files} "
        f"(解析対象{total_docs} / 要OCR{total_ocr} / 抽出エラー{total_err})"
    )
    print(
        f"フィールド抽出合計: found={total_found} not_found={total_not_found} "
        f"(v0対象フィールド={sum(1 for f in fields if f.v0_target)}/{len(fields)})"
    )
    if run_errors:
        print(f"案件単位のエラー: {len(run_errors)}件 → {run_errors}")
    print(f"出力先: {out_dir}")


if __name__ == "__main__":
    main()
