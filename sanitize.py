from __future__ import annotations

#!/usr/bin/env python3
"""
sanitize_excel_run.py
=====================
Excelファイル内の個人情報をサニタイズするスクリプト。SpaCy動作のためPython 3.12推奨。
1. VS Code でこのファイルを開く
2. CONFIG ブロックを書き換える
3. ツールバーの ▶Run  (または F5) を押す

すると:
    指定の Excel から PII を疑似値に置換し
    <出力ファイル> を保存 (置換マッピング JSON も任意で出力)

依存ライブラリや spaCy 日本語モデルは
最初に自動インストールされるので何もする必要はありません。
"""

# --------------------------------------------------------------------------- #
# CONFIG ――― ここを書き換えるだけで OK
# --------------------------------------------------------------------------- #
CONFIG = {
    # 処理したい Excel ファイル (.xlsx)
    "input_path": r"（インプットファイルパス）",

    # 処理対象シートを列挙。None なら全シート
    "sheets": ["（シート名）"],  # 例: ["社員一覧", "2025-Logs"]

    # 出力先 .xlsx (None なら <input>_sanitized.xlsx)
    "output_path": r"（アウトプットファイルパス）"

    # 置換マッピングを JSON 保存したい場合だけパスを指定
    "mapping_json_path": None,  # 例: r"./mapping.json"
}
# --------------------------------------------------------------------------- #


import dataclasses as _dc
import importlib
import json
import os
import re
import subprocess
import sys
from pathlib import Path
from typing import Dict, List, Tuple

# ----------------------- 自動インストール ---------------------------------- #
_PKGS = ("openpyxl", "faker", "spacy")

def _ensure_pkg(name: str) -> None:
    try:
        importlib.import_module(name)
    except ModuleNotFoundError:
        print(f"[INFO] Installing '{name}' ...")
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "--user", name],
            stdout=subprocess.DEVNULL,
        )

for _p in _PKGS:
    _ensure_pkg(_p)

import spacy
try:
    spacy.load("ja_core_news_lg")
except OSError:
    print("[INFO] Downloading spaCy model 'ja_core_news_lg' ...")
    import spacy.cli

    spacy.cli.download("ja_core_news_lg")

import faker
import openpyxl
_NLP = spacy.load("ja_core_news_lg")

# ---------------------- PII 検出・置換 ------------------------------------- #
import unicodedata                          # ← 追加 (全角→半角統一用)

_REGEX_PATTERNS: dict[str, str] = {
    # 国内携帯 & 固定電話 (ハイフン有無どちらも OK)
    "PHONE": r"""(?<!\d)(?:                       # 直前が数字でない
         0(?:70|80|90)(?:[-‐‑]?\d{4}){2}          # 090-1234-5678 / 09012345678
       | 0\d{1,4}[-‐‑]?\d{1,4}[-‐‑]?\d{3,4}       # 03-1234-5678 / 0312345678
       | 0(?:70|80|90)\d{8}                       # 09012345678
       | 0\d{9,10}                                # 0312345678 など
    )(?!\d)""",                                   # 直後が数字でない

    "EMAIL": r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b",

    "ZIP":   r"(?:〒\s*)?\b\d{3}[-‐‑]\d{4}\b",

    "IP":          r"\b(?:\d{1,3}\.){3}\d{1,3}\b",

    "CREDIT_CARD": r"\b(?:\d[ -]*?){13,16}\b",

    # 住所 (県市区町村名 + 番地数字 or “丁目”)
    "ADDR": r"""(?<![一-龥0-9])(?:                 # 手前が漢字・数字でない
         (?:                                      # ① 都道府県入り正式住所
             (?:北海道|(?:京都|大阪)府|.{2,3}県)?  # 都道府県 (任意)
             [一-龥]{1,8}(?:市|区|町|村)           # 市区町村
             [一-龥]{1,8}                          # 町域
             [\-‑‐]?\d{1,3}(?:丁目?)?              # 番地
         )
       |                                          # ② 簡易形「東京1」「大阪府2大阪市」
         [一-龥]{2,10}\d{1,3}(?:[一-龥])?           # 後ろに 1 文字漢字が来ても OK
    )""",
}

_SPACY_LABELS = {"PERSON", "GPE", "LOC", "DATE", "ORG","ADDR"}
_FAKE = faker.Faker("ja_JP")

@_dc.dataclass(slots=True)
class Replacement:
    span: Tuple[int, int]
    label: str
    fake: str

def _fake(label: str) -> str:
    fixed = {
        "PERSON":       "namexxx",
        "EMAIL":        "mailxxx",
        "PHONE":        "phonexxx",
        "ZIP":          "zipxxx",
        "IP":           "ipxxx",
        "CREDIT_CARD":  "cardxxx",
        "GPE":          "prefxxx",
        "LOC":          "cityxxx",
        "DATE":         "datexxx",
        "ORG":          "orgxxx",
        "ADDR":         "addrxxx",
    }
    return fixed.get(label, f"<{label}hogehoge>")


def _detect_regex(txt: str) -> List[Replacement]:
    flags = re.I | re.UNICODE | re.VERBOSE          # ★ VERBOSE を追加
    out: List[Replacement] = []
    for lab, pat in _REGEX_PATTERNS.items():
        for m in re.finditer(pat, txt, flags=flags):  # ★ flags を使う
            out.append(Replacement(m.span(), lab, _fake(lab)))
    return out

def _detect_spacy(txt: str) -> List[Replacement]:
    doc = _NLP(txt)
    return [
        Replacement((e.start_char, e.end_char), e.label_, _fake(e.label_))
        for e in doc.ents if e.label_ in _SPACY_LABELS
    ]

def sanitize_text(txt: str) -> Tuple[str, Dict[str, str]]:
    repls = _detect_regex(txt) + _detect_spacy(txt)
    repls.sort(key=lambda r: (r.span[0], -r.span[1]))

    merged, end = [], -1
    for r in repls:
        if r.span[0] >= end:
            merged.append(r); end = r.span[1]

    mapping, out = {}, txt
    for r in reversed(merged):
        orig = out[r.span[0]:r.span[1]]
        out = out[:r.span[0]] + r.fake + out[r.span[1]:]
        mapping[orig] = r.fake
    return out, mapping

# ----------------------- Excel 処理 ---------------------------------------- #
def sanitize_workbook(in_path: Path, sheets: List[str] | None):
    wb = openpyxl.load_workbook(in_path, data_only=False, keep_links=True)
    targets = sheets or wb.sheetnames
    all_map = {}

    for name in targets:
        if name not in wb.sheetnames:
            print(f"[WARN] Sheet '{name}' not found; skip.")
            continue
        ws, sheet_map = wb[name], {}
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and not cell.value.startswith("="):
                    san, mp = sanitize_text(cell.value)
                    if mp:
                        cell.value = san; sheet_map.update(mp)
        all_map[name] = sheet_map
        print(f"[INFO] {name}: {len(sheet_map)} replacements")
    return wb, all_map

# ----------------------- MAIN --------------------------------------------- #
def main():
    in_path  = Path(CONFIG["input_path"]).expanduser()
    out_path = (Path(CONFIG["output_path"])
                if CONFIG["output_path"] else in_path.with_stem(in_path.stem + "_sanitized"))
    sheets   = CONFIG["sheets"]
    map_json = CONFIG["mapping_json_path"]

    if not in_path.is_file():
        sys.exit(f"Input not found: {in_path}")

    wb, mapping = sanitize_workbook(in_path, sheets)
    wb.save(out_path)
    print(f"[DONE] Saved: {out_path}")

    if map_json:
        Path(map_json).write_text(json.dumps(mapping, ensure_ascii=False, indent=2),
                                  encoding="utf-8")
        print(f"[INFO] Mapping JSON: {map_json}")

if __name__ == "__main__":
    if os.name == "nt":          # Windows の文字化け対策
        os.environ.setdefault("PYTHONUTF8", "1")
    main()
