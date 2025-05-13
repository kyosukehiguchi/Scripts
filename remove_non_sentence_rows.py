#!/usr/bin/env python3
"""
remove_non_sentence_rows.py
Excelファイルの指定したカラムに自然文が含まれていない場合、行を削除するスクリプト。

使い方:
    # 何も指定しない場合はデフォルト値で実行
    python remove_non_sentence_rows.py

    # 任意に上書きしたいとき
    python remove_non_sentence_rows.py INPUT.xlsx OUTPUT.xlsx \
        --sheet Sheet1 --column コメント
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import pandas as pd

# ------------------------------------------------------------
# デフォルト値（環境に合わせて書き換えてください）
# ------------------------------------------------------------
DEFAULT_INPUT = (
    r"（インプットファイルパス）"
)
DEFAULT_OUTPUT = r"（アウトプットファイルパス）"
DEFAULT_SHEET = "（シート名）"
DEFAULT_COLUMN = "（カラム名）"


def is_sentence(text: str) -> bool:
    """シンプルなヒューリスティックで「自然文らしいか」を判定。"""
    if pd.isna(text) or not isinstance(text, str):
        return False

    s = text.strip()

    # 半角または全角スペースで分割して 2 語以上なら自然文とみなす
    if len(re.split(r"[ 　]", s)) >= 2:
        return True

    # 句読点・終止符が含まれているか
    if re.search(r"[。．\.!?？！]", s):
        return True

    return False


def main() -> None:
    parser = argparse.ArgumentParser(
        description="指定カラムが自然文でない行を削除して別ファイルに保存するスクリプト"
    )
    parser.add_argument(
        "input",
        nargs="?",
        default=DEFAULT_INPUT,
        help=f"入力 Excel ファイルパス (デフォルト: {DEFAULT_INPUT})",
    )
    parser.add_argument(
        "output",
        nargs="?",
        default=DEFAULT_OUTPUT,
        help=f"出力 Excel ファイルパス (デフォルト: {DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "--sheet",
        default=DEFAULT_SHEET,
        help=f"対象シート名 (デフォルト: {DEFAULT_SHEET})",
    )
    parser.add_argument(
        "--column",
        default=DEFAULT_COLUMN,
        help=f"対象カラム名 (デフォルト: {DEFAULT_COLUMN})",
    )
    args = parser.parse_args()

    in_path = Path(args.input)
    out_path = Path(args.output)

    # 入力ファイル存在チェック
    if not in_path.exists():
        sys.exit(f"入力ファイルが見つかりません: {in_path}")

    # DataFrame 読み込み
    try:
        df = pd.read_excel(in_path, sheet_name=args.sheet)
    except ValueError as e:
        sys.exit(f"シート '{args.sheet}' が見つかりません\n{e}")

    if args.column not in df.columns:
        sys.exit(f"カラム '{args.column}' が見つかりません。列名を確認してください。")

    # 自然文判定マスク
    mask_sentence = df[args.column].apply(is_sentence)
    removed_count = (~mask_sentence).sum()

    # 行削除
    df_clean = df[mask_sentence].copy()

    # 出力先ディレクトリがなければ作成
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # 保存
    df_clean.to_excel(out_path, index=False)

    print(f"完了: {removed_count} 行を削除しました。保存先: {out_path.resolve()}")


if __name__ == "__main__":
    main()
