# -*- coding: utf-8 -*-
import csv
import xlsxwriter

# 出典ファイルの略称 -> 実ファイルパス（トレーサビリティ用。出典ファイル一覧に対応）
SOURCE_FILES = {
    "依頼書": "inputfiles/2026_tdc_open_testbase/01_だんだん動物園入場システムテスト依頼書_20260426.pdf",
}

ROWS = [
    # ID, 要求アイテム, カテゴリ, 対象ファイル(略称), 出典ページ, 備考/Assumption
    ("A1", "入場制限者数を倍増できること（30分あたり30人→60人）", "機能要求", "依頼書", "p.6",
     "V02からの変更点。テスト設計スコープの中核"),
    ("A2", "入場制限者数をパラメータで0〜999人の範囲に容易に変更できること", "機能要求", "依頼書", "p.6",
     "「容易に変更可能」とあるため、設定変更機能自体もテスト対象になり得る"),
    ("A3", "入場ゲートを2台増設し、入場ゲートハブで集約稼働できること", "機能要求", "依頼書", "p.6, p.18",
     "V01→V02では変更がなかった「入場ゲートのハード＆ソフト」がV03で初めて変更される箇所（p.16）"),

    ("B1", "システムが長時間稼働することを保証したい", "品質重点要求", "依頼書", "p.11",
     "信頼性に関する要求。A3（ゲート増設）との関連が強い"),
    ("B2", "時間枠の重複販売が発生しないことを保証したい（制限人数超過の回避）", "品質重点要求", "依頼書", "p.11",
     "整合性・正確性に関する要求。A1/A2のパラメータ変更と直接関係"),
    ("B3", "「密回避」が達成できたことの効果を示したい", "品質重点要求", "依頼書", "p.11",
     "Assumption: 依頼書内に「これは入場システムV01の仕組みにより既に達成している」との注記あり。"
     "新規テスト要求というより既存機能による効果の証跡整理・説明を求めている可能性が高い"),

    ("C1", "決済システムはテスト対象外とする（テストエビデンスあり）", "スコープ要求", "依頼書", "p.22",
     "機材モデル上「決済システム」は影響分析の対象から除外してよい"),
    ("C2", "サブシステム単独で確認可能な要求（会員登録・変更・退会／問い合わせ／画面デザイン／QRコード作成）は開発会社保証済みのため対象外",
     "スコープ要求", "依頼書", "p.11",
     "機能モデル上、これらの機能は「保証済み」のフラグを付けて分類する想定"),
    ("C3", "旧システムから差分のない箇所は運用実績により相応の信頼性が確保されているとみなす", "スコープ要求", "依頼書", "p.11",
     "影響分析フェーズでの「機材依存度」評価時、差分有無が重要な判定軸になる"),
    ("C4", "園内およびWebでの予約購入、発券、入場という一連の流れを特に確認してほしい", "スコープ要求", "依頼書", "p.11",
     "「強くテストしてほしい点」の中核。ユースケース（p.26-28）と直接対応"),
    ("C5", "入場システムV02とV03の差分を強くテストしてほしい", "スコープ要求", "依頼書", "p.22",
     "テスト設計スコープの主軸。差分箇所の特定が前提となる（要：V02/V03差分の構造化）"),
    ("C6", "リスク分析を行い、テスト方針をご提案いただきたい", "スコープ要求", "依頼書", "p.22",
     "顧客は具体的なテスト観点の提示を委ねている＝スコアリングに基づく提案が特に評価される領域"),
    ("C7", "入場ゲートの変更点を考慮したリグレッションテストを設計してほしい", "スコープ要求", "依頼書", "p.22",
     "A3に直接対応。影響分析の優先対象"),
    ("C8", "今後の変更時に大きな問題とならないようなリグレッションテスト体系を考えてほしい", "スコープ要求", "依頼書", "p.22",
     "単発のテスト設計ではなく「再利用可能な回帰テスト基盤」の構築を求めている。本プロジェクト（plan.md）の目的そのものと一致する"),

    ("D1", "感染症状況の緩和に伴う入場者数制限緩和への対応", "背景リスク要求", "依頼書", "p.10",
     "A1/A2のビジネス上の動機"),
    ("D2", "人気動物の影響で入場予約が早期に埋まる状況への対応", "背景リスク要求", "依頼書", "p.10",
     "機能要求としては明記されていないが、将来「予約枠の追加・拡張」要求に発展する可能性がある（Assumption）"),
    ("D3", "近隣の大型マンション建設による同時入場者数増加・入場ゲート付近の滞留リスク増への備え", "背景リスク要求", "依頼書", "p.10",
     "A3（ゲート増設）の直接的な動機。将来さらなるゲート増設や導線変更が起きた際の回帰テスト設計（C8）の前提となるリスクシナリオ"),

    ("E1", "マイルストーン（テスト依頼→設計実施→第1回レビュー(予選)→設計改善→第2回レビュー(決勝)→実施→だんだん市へ提出）に沿って進めること",
     "プロセス制約", "依頼書", "p.7",
     "レビュー2段階構成。本プロジェクトのHuman-in-the-Loopと整合させる必要がある"),
    ("E2", "テスト設計成果物を、だんだん市への補助金申請における開発完了報告（「追加・変更点に品質上の問題がないこと」の証明）に添付する",
     "プロセス制約", "依頼書", "p.9（参考情報）",
     "Assumption: この記述はV02開発時の背景説明（「参考！」マーク付きスライド）であり、"
     "V03のテスト依頼概要には明記されていない。V03でも同じ制約が継続するかは未確認のため issue 化が必要"),
]

HEADER = ["ID", "要求アイテム", "カテゴリ", "対象ファイル", "出典ページ", "備考/Assumption", "レビューコメント"]

# ---- CSV ----
csv_path = "顧客要求アイテム一覧.csv"
with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
    writer = csv.writer(f)
    writer.writerow(HEADER)
    for row in ROWS:
        writer.writerow(list(row) + [""])
    writer.writerow([])
    writer.writerow(["# 出典ファイル一覧（略称 -> ファイルパス）"])
    for alias, path in SOURCE_FILES.items():
        writer.writerow([alias, path])

# ---- Excel ----
xlsx_path = "顧客要求アイテム一覧.xlsx"
workbook = xlsxwriter.Workbook(xlsx_path)
worksheet = workbook.add_worksheet("顧客要求アイテム")
ref_sheet = workbook.add_worksheet("出典ファイル一覧")

header_fmt = workbook.add_format({
    "bold": True, "bg_color": "#4472C4", "font_color": "white",
    "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True,
})
cell_fmt = workbook.add_format({"border": 1, "valign": "top", "text_wrap": True})
review_fmt = workbook.add_format({"border": 1, "valign": "top", "bg_color": "#FFF2CC", "text_wrap": True})
category_colors = {
    "機能要求": "#E2EFDA",
    "品質重点要求": "#FCE4D6",
    "スコープ要求": "#DDEBF7",
    "背景リスク要求": "#FFF2CC",
    "プロセス制約": "#EDEDED",
}
category_fmts = {
    cat: workbook.add_format({"border": 1, "valign": "top", "bg_color": color, "text_wrap": True})
    for cat, color in category_colors.items()
}

for col, title in enumerate(HEADER):
    worksheet.write(0, col, title, header_fmt)

worksheet.set_column(0, 0, 6)    # ID
worksheet.set_column(1, 1, 50)   # 要求アイテム
worksheet.set_column(2, 2, 14)   # カテゴリ
worksheet.set_column(3, 3, 12)   # 対象ファイル
worksheet.set_column(4, 4, 12)   # 出典ページ
worksheet.set_column(5, 5, 55)   # 備考/Assumption
worksheet.set_column(6, 6, 30)   # レビューコメント

for r, row in enumerate(ROWS, start=1):
    rid, item, category, source_file, source_page, note = row
    cat_fmt = category_fmts.get(category, cell_fmt)
    worksheet.write(r, 0, rid, cat_fmt)
    worksheet.write(r, 1, item, cat_fmt)
    worksheet.write(r, 2, category, cat_fmt)
    worksheet.write(r, 3, source_file, cat_fmt)
    worksheet.write(r, 4, source_page, cat_fmt)
    worksheet.write(r, 5, note, cat_fmt)
    worksheet.write(r, 6, "", review_fmt)

worksheet.freeze_panes(1, 0)
worksheet.autofilter(0, 0, len(ROWS), len(HEADER) - 1)

# 出典ファイル一覧シート（略称 -> パスの対応表。トレーサビリティ用）
ref_sheet.write(0, 0, "略称", header_fmt)
ref_sheet.write(0, 1, "ファイルパス", header_fmt)
ref_sheet.set_column(0, 0, 12)
ref_sheet.set_column(1, 1, 90)
for r, (alias, path) in enumerate(SOURCE_FILES.items(), start=1):
    ref_sheet.write(r, 0, alias, cell_fmt)
    ref_sheet.write(r, 1, path, cell_fmt)

workbook.close()

print(f"wrote {csv_path} and {xlsx_path}")
