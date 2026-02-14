import argparse
from pathlib import Path
from playwright.sync_api import sync_playwright

HPE_HOME = "https://support.hpe.com/connect/s/"

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="hpe_state.json", help="Output storage state json (cookies/session)")
    ap.add_argument("--url", default=HPE_HOME, help="Start URL (default: HPE Support Center home)")
    args = ap.parse_args()

    out_path = Path(args.out).resolve()

    print(f"➡️ Opening: {args.url}")
    print("ℹ️ Log in in the browser window (MFA etc.).")
    print("ℹ️ When you see your HPE portal loaded, come back here and press ENTER.")
    print(f"ℹ️ State will be saved to: {out_path}")

    with sync_playwright() as p:
        try:
            browser = p.chromium.launch(headless=False)
        except Exception as e:
            msg = str(e)
            if "Executable doesn't exist" in msg or "playwright install" in msg:
                print("\n❌ Playwright browser (Chromium) ontbreekt voor deze Python/venv.")
                print("✅ Fix (run EXACT):")
                print(r"   .\.venv\Scripts\python.exe -m playwright install chromium")
                print("Of doe volledige setup:")
                print(r"   powershell -ExecutionPolicy Bypass -File .\00_Setup.ps1")
                raise SystemExit(2)
            raise

        context = browser.new_context()
        page = context.new_page()
        page.goto(args.url, wait_until="domcontentloaded")

        input("\n✅ Druk ENTER om session state op te slaan... ")

        context.storage_state(path=str(out_path))
        browser.close()

    print(f"✅ Saved state: {out_path}")

if __name__ == "__main__":
    main()
