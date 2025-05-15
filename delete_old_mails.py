# -*- coding: utf-8 -*-
"""
title: Outlook 指定フォルダで『受信日時が 3 年以上前』のメールを一括削除

⚠️ 警告 ⚠️
このスクリプトは、指定した Outlook フォルダ内の『3年以上前のメール』を削除します。
実行する前に、必ずメールボックスのバックアップを取ってください。

設定項目 HARD_DELETE を False にすると、「削除済みアイテム」フォルダに移動されます。
True にすると完全に削除され、元に戻すことはできません（取り消し不可）。

使用には十分注意してください。

--------------------------------------------------------------------
pip install pywin32 python-dateutil
--------------------------------------------------------------------
"""

import sys, time
from datetime import datetime, timezone
from dateutil.relativedelta import relativedelta
import pythoncom, win32com.client
from win32com.client import constants as c

# ========= ユーザー設定 =========
SUBFOLDER_PATH = r"（ここにフォルダ名を記入）"     # 例 r"xxx"（'' で受信トレイ直下）
HARD_DELETE    = True    # True=完全削除 / False=削除済みアイテムへ移動
VERBOSE        = True
# ===============================

# Outlook 定数（makepy 未実行環境でも動くようフォールバック）
olInbox = getattr(c, "olFolderInbox", 6)
olDel   = getattr(c, "olFolderDeletedItems", 3)

# ---------- 共通ユーティリティ ----------

def resolve_subfolder(base, rel):
    """'受信トレイ\\サブ1' など相対パスから Folder を取得"""
    if not rel:
        return base
    cur = base
    for part in rel.split("\\"):
        cur = next((f for f in cur.Folders if f.Name.lower() == part.lower()), None)
        if cur is None:
            raise ValueError(f"サブフォルダ '{part}' が見つかりません")
    return cur

def to_utc(dt):
    """aware / naive を問わず UTC aware に変換"""
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=datetime.now().astimezone().tzinfo)
    return dt.astimezone(timezone.utc)

def iter_items_forward(items):
    """GetFirst / GetNext ですべてのアイテムを列挙（通信エラー時に自動リトライ）"""
    while True:
        try:
            itm = items.GetFirst()
            break
        except pythoncom.com_error:
            time.sleep(1)
    while itm:
        yield itm
        retry = 0
        while True:
            try:
                itm = items.GetNext()
                break
            except pythoncom.com_error:
                retry += 1
                if retry > 3:
                    raise                       # 3 回失敗したら停止
                time.sleep(1)

def delete_old(folder, cutoff_utc, hard_delete):
    session = folder.Application.Session
    items   = folder.Items
    items.Sort("[ReceivedTime]")               # 昇順＝最古→最新

    if not hard_delete:
        try:
            del_folder = folder.Store.GetDefaultFolder(olDel)
        except AttributeError:
            del_folder = session.GetDefaultFolder(olDel)

    removed = 0
    for itm in iter_items_forward(items):
        try:
            # メールアイテムのみ対象（Class=43）
            if getattr(itm, "Class", 0) != 43:
                continue
            rt = itm.ReceivedTime
            if to_utc(rt) < cutoff_utc:
                if hard_delete:
                    itm.Delete()
                else:
                    itm.Move(del_folder)
                removed += 1
                if VERBOSE and removed % 100 == 0:
                    print(f"{removed} 件削除…")
        except pythoncom.com_error:
            # S/MIME などアクセス不可でも無視して続行
            continue
    return removed

# ---------- エントリポイント ----------

def main():
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox   = outlook.GetDefaultFolder(olInbox)

    try:
        target = resolve_subfolder(inbox, SUBFOLDER_PATH)
    except ValueError as e:
        print(e, file=sys.stderr)
        return

    cutoff = datetime.now(timezone.utc) - relativedelta(years=3)

    if VERBOSE:
        print(f"基準日 (UTC): {cutoff:%Y-%m-%d %H:%M}")
        print(f"対象フォルダ : {target.FolderPath}")
        print("削除方式     :", "完全削除" if HARD_DELETE else "削除済みアイテムへ移動")
        print("-"*40)

    removed = delete_old(target, cutoff, HARD_DELETE)
    print(f"完了: {removed} 件を処理しました。")
    pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()
