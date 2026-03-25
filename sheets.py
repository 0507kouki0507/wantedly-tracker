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


def update_summary(ss: Spreadsheet, records: list[dict]) -> None:
    """サマリーシートを最新データで全書き換えする"""
    ws = _get_or_add_ws(ss, "サマリー", rows=500, cols=15)
    ws.clear()

    rows = [SUMMARY_HEADERS]
    for r in sorted(records, key=lambda x: -x["pv"]):
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
    n = len(rows)

    # ヘッダー行
    reqs.append(_fmt_req(sheet_id, 0, 1, 0, len(SUMMARY_HEADERS),
                         bg=COLOR_HEADER, bold=True, fg=COLOR_WHITE, halign="CENTER"))
    # 行ごとに色付け
    for i, r in enumerate(records, start=1):
        if r.get("status") == "募集停止中":
            bg = COLOR_STOPPED
        elif r["article_type"] == "ストーリー":
            bg = COLOR_STORY
        else:
            bg = COLOR_BOSHU
        reqs.append(_fmt_req(sheet_id, i, i + 1, 0, len(SUMMARY_HEADERS), bg=bg))

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

    sorted_records = sorted(records, key=lambda x: -x["daily_pv"])
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
