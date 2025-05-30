"""
【スクリプト概要】
    大規模システム開発プロジェクトのPMO向けに、要件定義から保守運用まで網羅した
    フォルダ構成（最大4階層）を自動生成するスクリプトです。

【動作環境】
    - Python 3.1 以降
    - Windows / macOS / Linux に対応
    - 実行後、指定したパス配下にプロジェクトフォルダ一式が作成されます

【使い方】
    python create_project_folders.py
    → 入力に従って、プロジェクト名と出力先フォルダを指定
"""

import os

# ユーザー入力
project_name = input("プロジェクト名を入力してください（例：MyProject_202503）: ").strip()
output_dir = input("作成先のフルパスを入力してください（例：C:/Users/you/Documents）: ").strip()
project_root = os.path.join(output_dir, project_name)

# 作成するフォルダ一覧
folders = [
    "01_プロジェクト管理/01_体制・連絡網",
    "01_プロジェクト管理/02_会議体・議事録/定例会議/アジェンダ",
    "01_プロジェクト管理/02_会議体・議事録/定例会議/議事録",
    "01_プロジェクト管理/02_会議体・議事録/ステアリングコミッティ",
    "01_プロジェクト管理/02_会議体・議事録/その他臨時会議",
    "01_プロジェクト管理/03_進捗管理/ガントチャート",
    "01_プロジェクト管理/03_進捗管理/進捗報告資料",
    "01_プロジェクト管理/04_課題・リスク管理",
    "01_プロジェクト管理/05_変更管理",
    "01_プロジェクト管理/06_成果物一覧・受領記録",

    "02_要件定義/01_業務要件",
    "02_要件定義/02_システム要件",
    "02_要件定義/03_非機能要件",
    "02_要件定義/04_要件定義書・レビュー記録",

    "03_基本設計/01_業務フロー",
    "03_基本設計/02_画面設計",
    "03_基本設計/03_帳票設計",
    "03_基本設計/04_IF設計",
    "03_基本設計/05_データモデル設計",
    "03_基本設計/06_基本設計書・レビュー記録",

    "04_詳細設計/01_機能設計",
    "04_詳細設計/02_内部設計",
    "04_詳細設計/03_詳細設計書・レビュー記録",

    "05_開発・構築/01_ソースコード/フロントエンド",
    "05_開発・構築/01_ソースコード/バックエンド",
    "05_開発・構築/01_ソースコード/バッチ処理",
    "05_開発・構築/02_開発環境構築手順",
    "05_開発・構築/03_構成管理",

    "06_テスト/01_単体テスト",
    "06_テスト/02_結合テスト",
    "06_テスト/03_システムテスト",
    "06_テスト/04_ユーザ受入テスト",
    "06_テスト/05_テスト結果・障害管理",

    "07_移行/01_移行計画",
    "07_移行/02_データ移行設計",
    "07_移行/03_移行ツール・スクリプト",
    "07_移行/04_移行実施結果・ログ",

    "08_教育・マニュアル/01_ユーザマニュアル",
    "08_教育・マニュアル/02_運用手順書",
    "08_教育・マニュアル/03_教育資料",

    "09_運用・保守/01_運用設計",
    "09_運用・保守/02_監視・障害対応",
    "09_運用・保守/03_保守エビデンス",
    "09_運用・保守/04_保守依頼・対応履歴",

    "10_納品物/2025_03_31納品",

    "99_参考資料/01_ベンダー資料",
    "99_参考資料/02_契約・見積関連",
    "99_参考資料/03_その他参考資料",
]

# フォルダ作成処理
for folder in folders:
    path = os.path.join(project_root, folder)
    os.makedirs(path, exist_ok=True)

print("\n✅ プロジェクトフォルダを作成しました！")
print(f"→ {project_root}")
