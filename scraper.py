"""Wantedly enterprise analytics を全記事スクレイプして Google スプレッドシートに書き込む"""

import re
import sys
from datetime import datetime
from pathlib import Path

from playwright.sync_api import Page, sync_playwright

import db
import sheets

SESSION_FILE = Path(__file__).parent / "data" / "session.json"

BOSHU_URL = "https://www.wantedly.com/enterprise/analytics/projects"
STORY_URL = "https://www.wantedly.com/enterprise/analytics/posts"


# ──────────────────────────────────────────
#  ページ取得
# ──────────────────────────────────────────

def load_full_page(page: Page, url: str) -> str:
    """ページを開き「もっと見る」を全部クリックしてテキストを返す"""
    page.goto(url, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(3000)

    if "login" in page.url or "signin" in page.url:
        print("セッション切れ。login.py を再実行してください。")
        sys.exit(1)

    # 「もっと見る」ボタンを全部クリック
    while True:
        btn = page.locator("text=もっと見る").first
        if btn.count() == 0:
            break
        try:
            btn.click(timeout=3000)
            page.wait_for_timeout(2000)
        except Exception:
            break

    return page.inner_text("body")


# ──────────────────────────────────────────
#  パーサー
# ──────────────────────────────────────────

def parse_boshu(text: str) -> list[dict]:
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    results = []
    i = 0
    while i < len(lines):
        # "・YYYY/MM/DDに編集" でエントリーを特定
        if not re.match(r"^・\d{4}/\d{2}/\d{2}に編集$", lines[i]):
            i += 1
            continue

        # タイトルと状態を後方から探す
        title, status = None, "不明"
        for j in range(i - 1, max(i - 10, -1), -1):
            if lines[j] in ("募集中", "募集停止中"):
                status = lines[j]
                title = lines[j - 1] if j > 0 else None
                break

        # 数値を最大4つ取得（PV, ブックマーク, 応募, 応援の順）
        nums = []
        j = i + 1
        while j < len(lines) and len(nums) < 4:
            if re.match(r"^\d+$", lines[j]):
                nums.append(int(lines[j]))
            else:
                break
            j += 1

        if title and nums:
            results.append({
                "title": title,
                "status": status,
                "pv":       nums[0] if len(nums) > 0 else 0,
                "bookmark": nums[1] if len(nums) > 1 else 0,
                "oubo":     nums[2] if len(nums) > 2 else 0,
                "ouen":     nums[3] if len(nums) > 3 else 0,
            })
        i += 1
    return results


def parse_stories(text: str) -> list[dict]:
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    results = []
    for i, line in enumerate(lines):
        # "YYYY/MM/DD" 単独行 + 前行が "・" → ストーリーエントリー
        if not re.match(r"^\d{4}/\d{2}/\d{2}$", line):
            continue
        if i < 3 or lines[i - 1] != "・":
            continue

        title = lines[i - 3]  # author → ・ → date なので3行前がタイトル

        pv = int(lines[i + 1]) if i + 1 < len(lines) and re.match(r"^\d+$", lines[i + 1]) else 0
        likes = int(lines[i + 2]) if i + 2 < len(lines) and re.match(r"^\d+$", lines[i + 2]) else 0

        results.append({"title": title, "pv": pv, "likes": likes})
    return results


# ──────────────────────────────────────────
#  差分計算
# ──────────────────────────────────────────

def calc_diff(last, key: str) -> int:
    return max(0, key - (last[key] if last else 0)) if isinstance(key, int) else 0


def _dedup(items: list[dict], key: str) -> list[dict]:
    """同じキーの重複を除去（最初の出現を残す）"""
    seen: set = set()
    result = []
    for item in items:
        k = item[key]
        if k not in seen:
            seen.add(k)
            result.append(item)
    return result


def diff(current: int, last_row, field: str) -> int:
    if last_row is None:
        return 0
    return max(0, current - (last_row[field] or 0))


# ──────────────────────────────────────────
#  メイン
# ──────────────────────────────────────────

def main() -> None:
    if not SESSION_FILE.exists():
        print("セッションがありません。先に login.py を実行してください。")
        sys.exit(1)

    db.init_db()
    conn = db.get_connection()

    print("スプレッドシートを準備中...")
    ss = sheets.get_spreadsheet()
    today = datetime.now().strftime("%Y-%m-%d")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            storage_state=str(SESSION_FILE),
            user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36",
        )
        try:
            page = context.new_page()

            # ── 募集 ──────────────────────────
            print("\n【募集】全記事を取得中...")
            boshu_text = load_full_page(page, BOSHU_URL)
            boshu_list = _dedup(parse_boshu(boshu_text), "title")
            print(f"  {len(boshu_list)} 件取得（重複除去後）")

            daily_records = []

            for a in boshu_list:
                key = f"boshu::{a['title']}"
                prev = db.get_last(conn, key)
                daily_pv   = max(0, a["pv"]   - (prev["pv"]   if prev else 0))
                daily_oubo = max(0, a["oubo"] - (prev["oubo"] if prev else 0))
                daily_records.append({
                    "article_type": "募集",
                    "title":        a["title"],
                    "status":       a["status"],
                    "pv":           a["pv"],
                    "oubo":         a["oubo"],
                    "likes":        0,
                    "daily_pv":     daily_pv,
                    "daily_oubo":   daily_oubo,
                    "daily_likes":  0,
                })
                db.save_history(conn, "募集", a["title"], a["status"],
                                a["pv"], a["bookmark"], a["oubo"], a["ouen"], 0)
                db.save_last(conn, key, a["pv"], a["bookmark"], a["oubo"], a["ouen"], 0)

            # ── ストーリー ────────────────────
            print("\n【ストーリー】全記事を取得中...")
            story_text = load_full_page(page, STORY_URL)
            story_list = _dedup(parse_stories(story_text), "title")
            print(f"  {len(story_list)} 件取得（重複除去後）")

            for a in story_list:
                key = f"story::{a['title']}"
                prev = db.get_last(conn, key)
                daily_pv    = max(0, a["pv"]    - (prev["pv"]    if prev else 0))
                daily_likes = max(0, a["likes"] - (prev["likes"] if prev else 0))
                daily_records.append({
                    "article_type": "ストーリー",
                    "title":        a["title"],
                    "status":       "",
                    "pv":           a["pv"],
                    "oubo":         0,
                    "likes":        a["likes"],
                    "daily_pv":     daily_pv,
                    "daily_oubo":   0,
                    "daily_likes":  daily_likes,
                })
                db.save_history(conn, "ストーリー", a["title"], "",
                                a["pv"], 0, 0, 0, a["likes"])
                db.save_last(conn, key, a["pv"], 0, 0, 0, a["likes"])

        finally:
            try:
                context.storage_state(path=str(SESSION_FILE))
            except Exception:
                pass
            browser.close()

    # ── Google Sheets に書き込み ──────────────
    print("\nスプレッドシートを更新中...")

    # 全記事の最新データを取得
    latest = db.get_all_latest(conn)
    records = latest.to_dict("records")

    # サマリーシート（全書き換え）
    sheets.update_summary(ss, records)

    # PVピボット（日付×記事）
    pv_data = {r["title"][:40]: r["pv"] for r in records}
    sheets.update_pivot(ss, "PV推移", today, pv_data, top_n=30)

    # 応募数ピボット（募集のみ）
    oubo_data = {r["title"][:40]: r["oubo"]
                 for r in records if r["article_type"] == "募集"}
    sheets.update_pivot(ss, "応募推移", today, oubo_data, top_n=30)

    # 日別サマリー（見やすいダッシュボードタブ）
    sheets.update_daily_summary(ss, daily_records, today)

    # チャート（初回のみ作成）
    sheets.create_chart_if_needed(ss, "PV推移",   "PVグラフ",   "PV推移グラフ（上位30記事）")
    sheets.create_chart_if_needed(ss, "応募推移", "応募グラフ", "応募数推移グラフ（上位30記事）")

    conn.close()
    print(f"\n完了: {ss.url}")


if __name__ == "__main__":
    main()
