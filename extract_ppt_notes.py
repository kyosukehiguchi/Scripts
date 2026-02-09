# スライドからノートを抽出してテキストファイルに保存するスクリプト
# extract_ppt_notes_fixed_path.py
# -*- coding: utf-8 -*-
#
# Windows + Microsoft PowerPoint + pywin32 前提
# pip install pywin32

from pathlib import Path
import sys
import win32com.client


# ===== ここだけ編集すればOK =====
PPTX_PATH = r"スライドファイルのフルパスをここに書く.pptx"  # 例: r"C:\work\example.pptx"
# 出力先を明示したい場合は指定（Noneなら <pptx名>_notes.txt を同じフォルダに作る）
OUTPUT_TXT_PATH = None  # 例: r"C:\work\out\example_notes.txt"
# ==============================


def normalize_newlines(s: str) -> str:
    return s.replace("\r\n", "\n").replace("\r", "\n")


def extract_notes_from_pptx(pptx_path: Path) -> list[tuple[int, str]]:
    """
    Returns: list of (slide_index_1based, notes_text)
    """
    powerpoint = None
    presentation = None

    try:
        # 既存PowerPointプロセスへの接続ではなく、新規インスタンス起動で安定化
        powerpoint = win32com.client.DispatchEx("PowerPoint.Application")

        # PowerPointは環境によって「非表示(Visible=0)」ができないため、最小化で運用する
        try:
            powerpoint.Visible = True
        except Exception:
            pass

        # ppWindowMinimized = 2
        try:
            powerpoint.WindowState = 2
        except Exception:
            pass

        # WithWindow=False でプレゼン自体のウィンドウを出さずに開く（環境差あり）
        presentation = powerpoint.Presentations.Open(
            str(pptx_path),
            WithWindow=False,
            ReadOnly=True
        )

        results: list[tuple[int, str]] = []
        slides = presentation.Slides

        for i in range(1, slides.Count + 1):
            slide = slides.Item(i)

            notes_text = ""
            # ノートは NotesPage 側に入っている
            # NotesPage.Shapes のうち、TextFrame を持つものを連結する
            try:
                notes_page = slide.NotesPage
                parts: list[str] = []

                for si in range(1, notes_page.Shapes.Count + 1):
                    shp = notes_page.Shapes.Item(si)
                    try:
                        if shp.HasTextFrame and shp.TextFrame.HasText:
                            t = shp.TextFrame.TextRange.Text
                            t = normalize_newlines(t).strip()
                            if t:
                                parts.append(t)
                    except Exception:
                        # shapeごとの取得失敗は無視して続行
                        pass

                notes_text = "\n".join(parts).strip()
            except Exception:
                notes_text = ""

            results.append((i, notes_text))

        return results

    finally:
        # 後始末（COMは残ると厄介なので確実に閉じる）
        try:
            if presentation is not None:
                presentation.Close()
        except Exception:
            pass
        try:
            if powerpoint is not None:
                powerpoint.Quit()
        except Exception:
            pass


def write_txt(output_path: Path, notes: list[tuple[int, str]]) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8", newline="\n") as f:
        for slide_no, text in notes:
            f.write(f"=== Slide {slide_no} ===\n")
            if text:
                f.write(text.rstrip() + "\n")
            else:
                f.write("(No notes)\n")
            f.write("\n")  # スライド間の空行


def main() -> None:
    pptx_path = Path(PPTX_PATH).expanduser().resolve()

    if not pptx_path.exists():
        print(f"ERROR: File not found: {pptx_path}", file=sys.stderr)
        sys.exit(1)

    if pptx_path.suffix.lower() not in [".pptx", ".pptm", ".ppt"]:
        print("ERROR: Please specify a PowerPoint file (.pptx/.pptm/.ppt).", file=sys.stderr)
        sys.exit(1)

    if OUTPUT_TXT_PATH:
        out_path = Path(OUTPUT_TXT_PATH).expanduser().resolve()
    else:
        out_path = pptx_path.with_name(pptx_path.stem + "_notes.txt")

    notes = extract_notes_from_pptx(pptx_path)
    write_txt(out_path, notes)

    print(f"Done. Wrote: {out_path}")
    print(f"Slides processed: {len(notes)}")


if __name__ == "__main__":
    main()
