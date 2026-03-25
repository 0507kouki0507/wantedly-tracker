"""Google スプレッドシートへの書き込みモジュール（ビジュアル重視版）"""

from datetime import datetime
from pathlib import Path
from typing import Optional

import gspread
from gspread import Spreadsheet, Worksheet
from gspread.utils import rowcol_to_a1

CREDENTIALS_FILE = Path(__file__).parent / "credentials.json"
SHEET_ID_FILE = Path(__file__).parent / "data" / "sheet_id.txt"

# ── カラーパレット ────────────────────────
COLOR_HEADER   = {"red": 0.18, "green": 0.34, "blue": 0.55}   # 濃い青
COLOR_BOSHU    = {"red": 0.84, "green": 0.92, "blue": 1.0}    # 薄い青
COLOR_STORY    = {"red": 0.85, "green": 1.0,  "blue": 0.88}   # 薄い緑
COLOR_STOPPED  = {"red": 0.93, "green": 0.93, "blue": 0.93}   # グレー
COLOR_WHITE    = {"red": 1.0,  "green": 1.0,  "blue": 1.0}


def get_client() -> gspread.Client:
    return gspread.service_account(filename=str(CREDENTIALS_FILE))


def get_spreadsheet() -> Spreadsheet:
    client = get_client()
    sheet_id = SHEET_ID_FILE.read_text().strip()
    ss = client.open_by_key(sheet_id)
    print(f"スプレッドシート: {ss.url}")
    return ss


def _get_or_add_ws(ss: Spreadsheet, title: str, rows: int = 500, cols: int = 50) -> Worksheet:
    try:
        return ss.worksheet(title)
    except gspread.WorksheetNotFound:
        return ss.add_worksheet(title=title, rows=rows, cols=cols)


# ══════════════════════════════════════════
#  サマリーシート（毎回全書き換え）
# ══════════════════════════════════════════

SUMMARY_HEADERS = [
    "種別", "タイトル", "状態",
    "初回取得日", "最終取得日",
    "累計PV", "ブックマーク", "応募数", "応援数", "いいね",
]


def _status_order(r: dict) -> int:
    """ソート優先度: 募集中 > ストーリー > 募集停止中"""
    if r.get("status") == "募集中":
        return 0
    elif r.get("article_type") == "ストーリー":
        return 1
    return 2


def update_summary(ss: Spreadsheet, records: list[dict]) -> None:
    """サマリーシートを最新データで全書き換えする"""
    ws = _get_or_add_ws(ss, "サマリー", rows=500, cols=15)
    ws.clear()

    sorted_records = sorted(records, key=lambda x: (_status_order(x), -x["pv"]))

    rows = [SUMMARY_HEADERS]
    for r in sorted_records:
        rows.append([
            r["article_type"],
            r["title"],
            r.get("status", ""),
            r["first_seen"][:10] if r.get("first_seen") else "",
            r["last_seen"][:10]  if r.get("last_seen")  else "",
            r["pv"],
            r.get("bookmark", 0),
            r.get("oubo", 0),
            r.get("ouen", 0),
            r.get("likes", 0),
        ])

    ws.update("A1", rows, value_input_option="USER_ENTERED")

    # ── 書式設定 ──────────────────────────
    reqs = []
    sheet_id = ws.id

    # ヘッダー行
    reqs.append(_fmt_req(sheet_id, 0, 1, 0, len(SUMMARY_HEADERS),
                         bg=COLOR_HEADER, bold=True, fg=COLOR_WHITE, halign="CENTER"))
    # 行ごとに色付け
    for i, r in enumerate(sorted_records, start=1):
        if r.get("status") == "募集停止中":
            bg = COLOR_STOPPED
            fg = {"red": 0.6, "green": 0.6, "blue": 0.6}
        elif r["article_type"] == "ストーリー":
            bg = COLOR_STORY
            fg = None
        else:
            bg = COLOR_BOSHU
            fg = None
        reqs.append(_fmt_req(sheet_id, i, i + 1, 0, len(SUMMARY_HEADERS), bg=bg, fg=fg))

    # 列幅調整
    reqs += [
        _col_width(sheet_id, 0, 1, 80),   # 種別
        _col_width(sheet_id, 1, 2, 350),  # タイトル
        _col_width(sheet_id, 2, 3, 90),   # 状態
        _col_width(sheet_id, 3, 4, 100),  # 初回取得日
        _col_width(sheet_id, 4, 5, 100),  # 最終取得日
    ]
    # 先頭行固定
    reqs.append({"updateSheetProperties": {
        "properties": {"sheetId": sheet_id,
                        "gridProperties": {"frozenRowCount": 1}},
        "fields": "gridProperties.frozenRowCount"
    }})

    if reqs:
        ss.batch_update({"requests": reqs})

    print(f"  サマリーシート更新: {len(records)} 件")


# ══════════════════════════════════════════
#  日別データシート（縦持ち・全履歴）
# ══════════════════════════════════════════

TALL_HEADERS = [
    "日付", "種別", "タイトル", "状態",
    "日別PV", "日別応募", "日別いいね",
    "累計PV", "累計応募", "累計いいね",
]


def update_tall_data(ss: Spreadsheet, records: list[dict], date_str: str) -> None:
    """縦持ち日別データシートに当日分を追記する"""
    ws = _get_or_add_ws(ss, "📋 日別データ", rows=10000, cols=12)

    existing = ws.get_all_values()

    # 初回：ヘッダー作成 + 書式
    if not existing or not existing[0]:
        ws.update("A1", [TALL_HEADERS], value_input_option="USER_ENTERED")
        reqs = [
            _fmt_req(ws.id, 0, 1, 0, len(TALL_HEADERS),
                     bg=COLOR_HEADER, bold=True, fg=COLOR_WHITE, halign="CENTER"),
            _col_width(ws.id, 0, 1, 100),   # 日付
            _col_width(ws.id, 1, 2, 90),    # 種別
            _col_width(ws.id, 2, 3, 320),   # タイトル
            _col_width(ws.id, 3, 4, 90),    # 状態
            {"updateSheetProperties": {
                "properties": {"sheetId": ws.id,
                                "gridProperties": {"frozenRowCount": 1}},
                "fields": "gridProperties.frozenRowCount",
            }},
        ]
        ss.batch_update({"requests": reqs})
        existing = [TALL_HEADERS]

    # 既存日付チェック（同日重複防止）
    existing_dates = {r[0] for r in existing[1:] if r}
    if date_str in existing_dates:
        print(f"  日別データ: {date_str} は既に記録済みです")
        return

    # 追記する行を作成
    rows = []
    for r in sorted(records, key=lambda x: -x["daily_pv"]):
        rows.append([
            date_str,
            r["article_type"],
            r["title"],
            r.get("status", ""),
            r["daily_pv"],
            r["daily_oubo"],
            r["daily_likes"],
            r["pv"],
            r["oubo"],
            r["likes"],
        ])

    ws.append_rows(rows, value_input_option="USER_ENTERED")
    print(f"  日別データ追記: {date_str} / {len(rows)} 件")


# ══════════════════════════════════════════
#  日別サマリーシート（見やすいダッシュボード）
# ══════════════════════════════════════════

DAILY_HEADERS = [
    "種別", "タイトル", "状態",
    "日別PV", "日別応募/いいね",
    "累計PV", "累計応募/いいね",
]


def update_daily_summary(ss: Spreadsheet, records: list[dict], date_str: str) -> None:
    """日別サマリーシートを最新データで全書き換え"""
    ws = _get_or_add_ws(ss, "📊 今日の動向", rows=200, cols=10)
    ws.clear()

    rows: list[list] = []
    rows.append([f"📅 {date_str}  |  Wantedly 日別動向レポート", "", "", "", "", "", ""])
    rows.append(DAILY_HEADERS)

    sorted_records = sorted(records, key=lambda x: (_status_order(x), -x["daily_pv"]))
    for r in sorted_records:
        metric = r["daily_oubo"] if r["article_type"] == "募集" else r["daily_likes"]
        cumulative_metric = r["oubo"] if r["article_type"] == "募集" else r["likes"]
        label = "応募" if r["article_type"] == "募集" else "いいね"
        rows.append([
            r["article_type"],
            r["title"],
            r.get("status", ""),
            r["daily_pv"],
            metric,
            r["pv"],
            cumulative_metric,
        ])

    ws.update("A1", rows, value_input_option="USER_ENTERED")

    # ── 書式設定 ──────────────────────────
    reqs = []
    sheet_id = ws.id
    n = len(rows)

    # タイトル行（濃い青・白文字・太字）
    reqs.append(_fmt_req(sheet_id, 0, 1, 0, 7,
                         bg=COLOR_HEADER, bold=True, fg=COLOR_WHITE))
    # ヘッダー行
    reqs.append(_fmt_req(sheet_id, 1, 2, 0, 7,
                         bg={**COLOR_HEADER, "red": 0.25}, bold=True, fg=COLOR_WHITE, halign="CENTER"))

    # データ行の色分け
    for i, r in enumerate(sorted_records, start=2):
        if r.get("status") == "募集停止中":
            bg = COLOR_STOPPED
        elif r["article_type"] == "ストーリー":
            bg = COLOR_STORY
        else:
            bg = COLOR_BOSHU
        reqs.append(_fmt_req(sheet_id, i, i + 1, 0, 7, bg=bg))

        # 日別PV列（数値が大きいほど目立つ色）
        if r["daily_pv"] > 0:
            reqs.append(_fmt_req(sheet_id, i, i + 1, 3, 5,
                                 bg={"red": 0.7, "green": 0.9, "blue": 0.7}))

    # 列幅調整
    reqs += [
        _col_width(sheet_id, 0, 1, 90),   # 種別
        _col_width(sheet_id, 1, 2, 320),  # タイトル
        _col_width(sheet_id, 2, 3, 100),  # 状態
        _col_width(sheet_id, 3, 4, 90),   # 日別PV
        _col_width(sheet_id, 4, 5, 120),  # 日別応募/いいね
        _col_width(sheet_id, 5, 6, 90),   # 累計PV
        _col_width(sheet_id, 6, 7, 120),  # 累計応募/いいね
    ]

    # 先頭2行固定
    reqs.append({"updateSheetProperties": {
        "properties": {"sheetId": sheet_id,
                        "gridProperties": {"frozenRowCount": 2}},
        "fields": "gridProperties.frozenRowCount"
    }})

    if reqs:
        ss.batch_update({"requests": reqs})

    print(f"  日別サマリー更新: {len(records)} 件")


# ══════════════════════════════════════════
#  PV推移シート（日付×記事のピボット）
# ══════════════════════════════════════════

def update_pivot(ss: Spreadsheet, sheet_name: str, date_str: str,
                 data: dict[str, int], top_n: int = 30) -> None:
    """
    date_str: "2026-03-24"
    data: {title: value}  記事タイトル→数値
    """
    ws = _get_or_add_ws(ss, sheet_name, rows=500, cols=top_n + 2)

    existing = ws.get_all_values()

    if not existing or not existing[0]:
        # 初回: ヘッダー行を作成（上位N記事）
        top_titles = sorted(data, key=lambda t: -data[t])[:top_n]
        header = ["日付"] + [t[:40] for t in top_titles]
        ws.update("A1", [header], value_input_option="USER_ENTERED")
        existing = [header]

        # ヘッダー書式
        _apply_header_format(ss, ws, len(header))

    header_row = existing[0]      # ["日付", タイトル1, タイトル2, ...]
    titles = header_row[1:]       # タイトルリスト

    # 既存の日付行を確認（重複防止）
    existing_dates = [r[0] for r in existing[1:] if r]
    if date_str in existing_dates:
        print(f"  {sheet_name}: {date_str} は既に記録済みです")
        return

    # 新しい行を末尾に追加
    new_row = [date_str] + [data.get(t[:40], data.get(t, 0)) for t in titles]
    ws.append_row(new_row, value_input_option="USER_ENTERED")
    print(f"  {sheet_name}: {date_str} を追加")


def _apply_header_format(ss: Spreadsheet, ws: Worksheet, ncols: int) -> None:
    reqs = [
        _fmt_req(ws.id, 0, 1, 0, ncols, bg=COLOR_HEADER, bold=True,
                 fg=COLOR_WHITE, halign="CENTER"),
        {"updateSheetProperties": {
            "properties": {"sheetId": ws.id,
                            "gridProperties": {"frozenRowCount": 1,
                                               "frozenColumnCount": 1}},
            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"
        }},
        _col_width(ws.id, 0, 1, 100),  # 日付列
    ]
    ss.batch_update({"requests": reqs})


# ══════════════════════════════════════════
#  チャート作成
# ══════════════════════════════════════════

def create_chart_if_needed(ss: Spreadsheet, pivot_sheet_name: str,
                            chart_sheet_name: str, title: str) -> None:
    """ピボットシートからチャートシートを作成する（初回のみ）"""
    try:
        ss.worksheet(chart_sheet_name)
        return  # 既に存在する場合はスキップ
    except gspread.WorksheetNotFound:
        pass

    pivot_ws = ss.worksheet(pivot_sheet_name)
    pivot_id = pivot_ws.id
    n_cols = len(pivot_ws.row_values(1))  # 列数
    n_rows = pivot_ws.row_count

    if n_cols < 2:
        return

    # チャート用の新シートを作成
    chart_ws = ss.add_worksheet(title=chart_sheet_name, rows=30, cols=10)

    series = []
    for col in range(1, min(n_cols, 11)):  # 最大10系列
        series.append({
            "series": {
                "sourceRange": {"sources": [{
                    "sheetId": pivot_id,
                    "startRowIndex": 0, "endRowIndex": n_rows,
                    "startColumnIndex": col, "endColumnIndex": col + 1,
                }]}
            },
            "targetAxis": "LEFT_AXIS",
        })

    req = {"addChart": {"chart": {
        "spec": {
            "title": title,
            "basicChart": {
                "chartType": "LINE",
                "legendPosition": "BOTTOM_LEGEND",
                "axis": [
                    {"position": "BOTTOM_AXIS", "title": "日付"},
                    {"position": "LEFT_AXIS",   "title": "数値"},
                ],
                "domains": [{"domain": {"sourceRange": {"sources": [{
                    "sheetId": pivot_id,
                    "startRowIndex": 0, "endRowIndex": n_rows,
                    "startColumnIndex": 0, "endColumnIndex": 1,
                }]}}}],
                "series": series,
                "headerCount": 1,
            },
        },
        "position": {"overlayPosition": {
            "anchorCell": {"sheetId": chart_ws.id,
                           "rowIndex": 0, "columnIndex": 0},
            "widthPixels": 900, "heightPixels": 500,
        }},
    }}}

    ss.batch_update({"requests": [req]})
    print(f"  チャート作成: {chart_sheet_name}")


# ══════════════════════════════════════════
#  ユーティリティ
# ══════════════════════════════════════════

def update_trend_pivot(ss: Spreadsheet, history_df, latest_records: list[dict]) -> None:
    """記事別×日別の推移ピボットシートを作成・更新"""
    import pandas as pd

    ws = _get_or_add_ws(ss, "📈 記事別推移", rows=500, cols=100)
    ws.clear()

    if history_df.empty:
        return

    # 日付リスト（最新30日）
    all_dates = sorted(history_df["date"].unique())[-30:]
    short_dates = [d[5:].replace("-", "/") for d in all_dates]  # "03/25" 形式

    # 記事ごとに日別PVを計算
    pivot: dict[str, dict[str, int]] = {}
    meta: dict[str, dict] = {}

    for title, group in history_df.groupby("title"):
        group = group.sort_values("date")
        pvs = dict(zip(group["date"], group["pv"]))
        oubo = dict(zip(group["date"], group["oubo"]))
        likes = dict(zip(group["date"], group["likes"]))
        row_meta = group.iloc[-1]
        meta[title] = {
            "article_type": row_meta["article_type"],
            "status": row_meta["status"] or "",
        }
        daily: dict[str, int] = {}
        prev_pv = 0
        for d in all_dates:
            cur = pvs.get(d, prev_pv)
            daily[d] = max(0, cur - prev_pv)
            prev_pv = cur if d in pvs else prev_pv
        pivot[title] = daily

    # ソート順：募集中 → ストーリー → 募集停止中、各グループ内は累計PV降順
    latest_pv = {r["title"]: r["pv"] for r in latest_records}

    def sort_key(title):
        m = meta[title]
        order = _status_order({"status": m["status"], "article_type": m["article_type"]})
        return (order, -latest_pv.get(title, 0))

    sorted_titles = sorted(pivot.keys(), key=sort_key)

    # ヘッダー行
    fixed_cols = ["種別", "タイトル", "状態"]
    header = fixed_cols + short_dates
    rows: list[list] = [header]

    # セクション区切りとデータ行
    sections = [
        (0, "▼ 募集中"),
        (1, "▼ ストーリー"),
        (2, "▼ 募集停止中"),
    ]
    section_row_indices: dict[int, int] = {}  # section_order → row index

    current_section = -1
    for title in sorted_titles:
        m = meta[title]
        sec = _status_order({"status": m["status"], "article_type": m["article_type"]})
        if sec != current_section:
            current_section = sec
            section_row_indices[sec] = len(rows)
            label = next(lbl for s, lbl in sections if s == sec)
            rows.append([label] + [""] * (len(header) - 1))
        daily = pivot[title]
        rows.append([
            m["article_type"],
            title,
            m["status"],
        ] + [daily.get(d, 0) for d in all_dates])

    ws.update("A1", rows, value_input_option="USER_ENTERED")

    # ── 書式設定 ──────────────────────────
    reqs = []
    sheet_id = ws.id
    ncols = len(header)

    # ヘッダー行
    reqs.append(_fmt_req(sheet_id, 0, 1, 0, ncols,
                         bg=COLOR_HEADER, bold=True, fg=COLOR_WHITE, halign="CENTER"))

    # セクション見出し行
    sec_bg = {"red": 0.3, "green": 0.3, "blue": 0.3}
    for sec_order, label in sections:
        if sec_order in section_row_indices:
            ri = section_row_indices[sec_order]
            reqs.append(_fmt_req(sheet_id, ri, ri + 1, 0, ncols,
                                 bg=sec_bg, bold=True, fg=COLOR_WHITE))

    # データ行の色分け
    for i, row in enumerate(rows):
        if i == 0:
            continue
        if not row[0] or row[0].startswith("▼"):
            continue
        article_type = row[0]
        status = row[2] if len(row) > 2 else ""
        if status == "募集停止中":
            bg = COLOR_STOPPED
        elif article_type == "ストーリー":
            bg = COLOR_STORY
        else:
            bg = COLOR_BOSHU
        reqs.append(_fmt_req(sheet_id, i, i + 1, 0, ncols, bg=bg))

    # 列幅調整
    reqs += [
        _col_width(sheet_id, 0, 1, 80),   # 種別
        _col_width(sheet_id, 1, 2, 300),  # タイトル
        _col_width(sheet_id, 2, 3, 90),   # 状態
    ]
    for ci in range(3, ncols):
        reqs.append(_col_width(sheet_id, ci, ci + 1, 55))  # 日付列

    # 先頭1行・3列固定
    reqs.append({"updateSheetProperties": {
        "properties": {"sheetId": sheet_id,
                        "gridProperties": {"frozenRowCount": 1,
                                           "frozenColumnCount": 3}},
        "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount",
    }})

    if reqs:
        ss.batch_update({"requests": reqs})

    print(f"  記事別推移更新: {len(sorted_titles)} 記事 × {len(all_dates)} 日")


def update_guide_sheet(ss: Spreadsheet) -> None:
    """使い方ガイドシートを作成・更新する"""
    ws = _get_or_add_ws(ss, "📖 使い方ガイド", rows=60, cols=6)
    ws.clear()

    C_TITLE  = {"red": 0.18, "green": 0.34, "blue": 0.55}   # 濃い青
    C_SEC    = {"red": 0.23, "green": 0.47, "blue": 0.71}   # 中青
    C_YELLOW = {"red": 1.0,  "green": 0.95, "blue": 0.8}    # 薄い黄
    C_BLUE   = {"red": 0.84, "green": 0.92, "blue": 1.0}    # 薄い青
    C_GREEN  = {"red": 0.85, "green": 1.0,  "blue": 0.88}   # 薄い緑
    C_GRAY   = {"red": 0.93, "green": 0.93, "blue": 0.93}   # グレー
    C_WHITE  = {"red": 1.0,  "green": 1.0,  "blue": 1.0}

    rows = [
        # タイトル
        ["📊 Wantedly Analytics  |  シートの見方", "", "", "", "", ""],
        ["", "", "", "", "", ""],

        # 概要
        ["▌ このスプレッドシートについて", "", "", "", "", ""],
        ["Wantedly の募集・ストーリー記事の閲覧数・応募数を毎朝9時に自動で取得し記録しています。", "", "", "", "", ""],
        ["データは前日比の日別数値と累計数値の両方を確認できます。", "", "", "", "", ""],
        ["", "", "", "", "", ""],

        # タブ説明ヘッダー
        ["▌ タブの使い方", "", "", "", "", ""],
        ["タブ名", "こんなときに使う", "確認頻度", "ポイント", "", ""],

        # 各タブ
        ["📊 今日の動向",
         "昨日と比べて、どの記事が伸びたか確認したいとき",
         "毎日",
         "日別PVが緑ハイライト。上から「募集中 → ストーリー → 募集停止中」の順に並んでいます。", "", ""],

        ["📈 記事別推移",
         "記事ごとに日々の数値がどう変化したか流れを見たいとき",
         "週1〜2回",
         "左3列（種別・タイトル・状態）は固定。日付は右にスクロールすると過去分も見られます。", "", ""],

        ["サマリー",
         "全記事の累計PV・累計応募数のランキングを見たいとき",
         "月1回程度",
         "累計値なので日々の変化は見えません。全体の規模感を把握するのに使います。", "", ""],

        ["📋 日別データ",
         "データをエクスポートしたり、独自に集計・分析したいとき",
         "必要なとき",
         "全データが縦持ち形式で蓄積されています。ピボットテーブルの元データとしても使えます。", "", ""],

        ["", "", "", "", "", ""],

        # 色の凡例
        ["▌ 色の意味", "", "", "", "", ""],
        ["　", "募集中の記事", "", "", "", ""],
        ["　", "ストーリー記事", "", "", "", ""],
        ["　", "募集停止中の記事", "", "", "", ""],
        ["　", "日別PVに動きがあった行（今日の動向のみ）", "", "", "", ""],
        ["", "", "", "", "", ""],

        # 更新・運用
        ["▌ 更新について", "", "", "", "", ""],
        ["更新タイミング", "毎朝 9:00 に自動実行（Mac が起動していること）", "", "", "", ""],
        ["手動実行",       "ターミナルで   uv run python scraper.py   を実行", "", "", "", ""],
        ["セッション切れ", "ターミナルで   uv run python login.py   を実行してログインし直す", "", "", "", ""],
    ]

    ws.update("A1", rows, value_input_option="USER_ENTERED")

    reqs = []
    sid = ws.id

    # タイトル行
    reqs.append(_fmt_req(sid, 0, 1, 0, 6, bg=C_TITLE, bold=True, fg=C_WHITE))
    # セクション見出し行
    for ri in [2, 6, 13, 18]:
        reqs.append(_fmt_req(sid, ri, ri + 1, 0, 6, bg=C_SEC, bold=True, fg=C_WHITE))
    # テーブルヘッダー
    reqs.append(_fmt_req(sid, 7, 8, 0, 4, bg=C_TITLE, bold=True, fg=C_WHITE, halign="CENTER"))
    # 各タブ行の色
    tab_colors = [C_BLUE, C_BLUE, C_GRAY, C_GRAY]
    for i, color in enumerate(tab_colors):
        reqs.append(_fmt_req(sid, 8 + i, 9 + i, 0, 4, bg=color))
    # 色凡例
    reqs.append(_fmt_req(sid, 14, 15, 0, 2, bg=C_BLUE))
    reqs.append(_fmt_req(sid, 15, 16, 0, 2, bg=C_GREEN))
    reqs.append(_fmt_req(sid, 16, 17, 0, 2, bg=C_GRAY))
    reqs.append(_fmt_req(sid, 17, 18, 0, 2, bg={"red": 0.7, "green": 0.9, "blue": 0.7}))
    # 更新セクション行
    for ri in [19, 20, 21]:
        reqs.append(_fmt_req(sid, ri, ri + 1, 0, 1, bold=True))

    # 列幅
    reqs += [
        _col_width(sid, 0, 1, 160),
        _col_width(sid, 1, 2, 380),
        _col_width(sid, 2, 3, 100),
        _col_width(sid, 3, 4, 400),
    ]

    # 行の高さをタイトル行だけ大きく
    reqs.append({"updateDimensionProperties": {
        "range": {"sheetId": sid, "dimension": "ROWS", "startIndex": 0, "endIndex": 1},
        "properties": {"pixelSize": 40},
        "fields": "pixelSize",
    }})

    if reqs:
        ss.batch_update({"requests": reqs})

    print("  使い方ガイド更新")


def delete_unused_sheets(ss: Spreadsheet) -> None:
    """不要なタブを削除する"""
    to_delete = ["PV推移", "応募推移", "PVグラフ", "応募グラフ"]
    for name in to_delete:
        try:
            ws = ss.worksheet(name)
            ss.del_worksheet(ws)
            print(f"  シート削除: {name}")
        except gspread.WorksheetNotFound:
            pass


def _fmt_req(sheet_id: int, r1: int, r2: int, c1: int, c2: int,
             bg=None, bold: bool = False, fg=None,
             halign: Optional[str] = None) -> dict:
    fmt: dict = {}
    fields = []
    if bg:
        fmt["backgroundColor"] = bg
        fields.append("backgroundColor")
    if bold or fg:
        fmt["textFormat"] = {}
        if bold:
            fmt["textFormat"]["bold"] = True
            fields.append("textFormat.bold")
        if fg:
            fmt["textFormat"]["foregroundColor"] = fg
            fields.append("textFormat.foregroundColor")
    if halign:
        fmt["horizontalAlignment"] = halign
        fields.append("horizontalAlignment")
    return {
        "repeatCell": {
            "range": {"sheetId": sheet_id,
                      "startRowIndex": r1, "endRowIndex": r2,
                      "startColumnIndex": c1, "endColumnIndex": c2},
            "cell": {"userEnteredFormat": fmt},
            "fields": "userEnteredFormat(" + ",".join(fields) + ")",
        }
    }


def _col_width(sheet_id: int, c1: int, c2: int, px: int) -> dict:
    return {
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": c1, "endIndex": c2},
            "properties": {"pixelSize": px},
            "fields": "pixelSize",
        }
    }
