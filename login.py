"""Wantedly に手動ログインしてセッションを保存する"""

from pathlib import Path
from playwright.sync_api import sync_playwright

SESSION_FILE = Path(__file__).parent / "data" / "session.json"


def main() -> None:
    SESSION_FILE.parent.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            )
        )
        page = context.new_page()

        print("ブラウザで以下のページを開いています...")
        print("→ https://www.wantedly.com/enterprise/analytics")
        print()
        print("ログインを求められたらログインしてください。")
        print("アナリティクスページが表示されたら Enter を押してください。")
        page.goto("https://www.wantedly.com/enterprise/analytics", wait_until="domcontentloaded")

        input("\nアナリティクスページが表示されたら Enter を押してください...")

        context.storage_state(path=str(SESSION_FILE))
        print(f"セッション保存完了: {SESSION_FILE}")
        browser.close()


if __name__ == "__main__":
    main()
