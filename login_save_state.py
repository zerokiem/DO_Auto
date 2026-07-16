from pathlib import Path
from playwright.sync_api import sync_playwright

AUTH_DIR = Path("playwright/.auth")
AUTH_DIR.mkdir(parents=True, exist_ok=True)

STATE_FILE = AUTH_DIR / "state.json"

URL = "https://doffice.npt.com.vn/sign-in"  # change to your target URL

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()

    page.goto(URL)

    print("Đăng nhập thủ công trên browser vừa mở.")
    input("Sau khi đăng nhập xong, nhấn Enter ở cửa sổ PowerShell...")

    context.storage_state(path=str(STATE_FILE))
    print(f"Đã lưu session vào: {STATE_FILE}")

    browser.close()