import os
import sys
import time
import threading
import subprocess
import requests
from playwright.sync_api import sync_playwright, Page

#####################################################################
VERSION = "1.0.0"  # current version of your app
VERSION_URL = "https://yourserver.com/latest_version.txt"  # text file with latest version (e.g. 1.0.1)
EXE_URL = "https://yourserver.com/MyApp-latest.exe"        # URL to your latest compiled exe
APP_NAME = "MyApp.exe"                                     # name of this executable
#####################################################################

def check_for_update():
    try:
        print("🔍 Checking for updates...")
        latest = requests.get(VERSION_URL, timeout=5).text.strip()
        if latest != VERSION:
            print(f"🚀 New version {latest} available. Downloading update...")
            r = requests.get(EXE_URL, timeout=30)
            tmp_file = "update.exe"
            with open(tmp_file, "wb") as f:
                f.write(r.content)

            os.replace(tmp_file, APP_NAME)
            print("✅ Update installed. Restarting application...")
            subprocess.Popen([APP_NAME])
            sys.exit(0)
        else:
            print("✅ You are running the latest version.")
    except Exception as e:
        print("⚠️ Update check failed:", e)

#####################################################################

USER_PROFILES = {
    str(i): fr"C:\Users\Abraham\AppData\Local\Microsoft\Edge\User Data\Profile {i}"
    for i in range(24)
}
EXTENSION_PATH = (
    r"C:\Users\Abraham\AppData\Local\Microsoft\Edge\User Data\Default\Extensions"
    r"\ejbalbakoplchlghecdalmeeeajnimhm\12.17.2_0"
)
BROWSER_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
#####################################################################


def safe_goto(page: Page, url: str, retries=5, delay=2):
    for attempt in range(1, retries + 1):
        try:
            print(f"🌐 Navigating to {url} (attempt {attempt})")
            page.goto(url)
            page.wait_for_load_state("domcontentloaded")
            print("✅ Navigation successful.")
            return page
        except Exception as e:
            print(f"ℹ️ Navigation error: {e}")
            if attempt < retries:
                print(f"🔄 Retrying in {delay} seconds...")
                time.sleep(delay)
            else:
                raise
    return page


def run1(profile_number: str, user_data_dir: str):
    print(f"🚀 Launching Edge with Profile {profile_number}")

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=user_data_dir,
            executable_path=BROWSER_PATH,
            headless=False,
            args=[
                f"--disable-extensions-except={EXTENSION_PATH}",
                f"--load-extension={EXTENSION_PATH}",
            ],
        )

        # open extension page
        extension_page = context.pages[0]
        safe_goto(
            extension_page,
            "chrome-extension://ejbalbakoplchlghecdalmeeeajnimhm/home.html",
        )
        extension_page.set_viewport_size({"width": 1000, "height": 800})

        # open site page
        site_page = context.pages[1]
        safe_goto(site_page, "https://tokenstaking.io/")
        site_page.set_viewport_size({"width": 1000, "height": 800})

        input(f"📌 Press Enter to close Profile {profile_number}")
        context.close()


#####################################################################
# --- Main Entry Point ---
#####################################################################
if __name__ == "__main__":
    check_for_update()  # 🔄 Run the auto-update check first

    choice = input("📂 Enter profile number = ").strip()
    if choice not in USER_PROFILES:
        print(f"❌ Invalid profile number '{choice}'. Please check USER_PROFILES.")
        sys.exit(1)

    print(
        "\nTASK MENU\n"
        "1 = MetaMask Staking\n"
        "2 = MetaMask Farming\n"
        "3 = Level 1 Data\n"
        "4 = BNB Check\n"
        "5 = PVC & BNB Data\n"
    )

    task_choice = input("⚙️ Which task to run = ").strip()
    if task_choice == "1":
        task = run1
    else:
        print("❌ Invalid task choice. Please run again.")
        sys.exit(1)

    thread = threading.Thread(target=task, args=(choice, USER_PROFILES[choice]))
    thread.start()
    thread.join()
    print("🆑 Browser closed.")
