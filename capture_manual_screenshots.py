import os
import time
import json
from playwright.sync_api import sync_playwright

SCREENSHOT_DIR = os.path.join(os.getcwd(), "screenshots_manual")
os.makedirs(SCREENSHOT_DIR, exist_ok=True)

CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

with open("auth_tokens.json", "r") as f:
    tokens = json.load(f)

admin_token = tokens["adminToken"]
user_token = tokens["userToken"]

def run():
    with sync_playwright() as p:
        browser = p.chromium.launch(
            executable_path=CHROME_PATH,
            headless=True
        )
        
        # 1. Public Pages (Login & Register)
        context_pub = browser.new_context(viewport={"width": 1440, "height": 900}, device_scale_factor=2)
        page_pub = context_pub.new_page()

        print("Capturing 01_login.png...")
        page_pub.goto("http://localhost:3000/login", wait_until="networkidle")
        page_pub.wait_for_timeout(1000)
        page_pub.screenshot(path=os.path.join(SCREENSHOT_DIR, "01_login.png"))

        print("Capturing 02_register.png...")
        page_pub.goto("http://localhost:3000/register", wait_until="networkidle")
        page_pub.wait_for_timeout(1000)
        page_pub.screenshot(path=os.path.join(SCREENSHOT_DIR, "02_register.png"))
        context_pub.close()

        # 2. User Pages (with User Token)
        context_user = browser.new_context(viewport={"width": 1440, "height": 900}, device_scale_factor=2)
        context_user.add_cookies([{
            "name": "token",
            "value": user_token,
            "domain": "localhost",
            "path": "/"
        }])
        page_user = context_user.new_page()

        print("Capturing 03_user_dashboard.png...")
        page_user.goto("http://localhost:3000/user/dashboard", wait_until="networkidle")
        page_user.wait_for_timeout(1200)
        page_user.screenshot(path=os.path.join(SCREENSHOT_DIR, "03_user_dashboard.png"))

        print("Capturing 04_user_setor_sampah.png...")
        page_user.goto("http://localhost:3000/user/laporan", wait_until="networkidle")
        page_user.wait_for_timeout(1200)
        page_user.screenshot(path=os.path.join(SCREENSHOT_DIR, "04_user_setor_sampah.png"))

        print("Capturing 05_user_riwayat.png...")
        page_user.goto("http://localhost:3000/user/riwayat", wait_until="networkidle")
        page_user.wait_for_timeout(1200)
        page_user.screenshot(path=os.path.join(SCREENSHOT_DIR, "05_user_riwayat.png"))
        context_user.close()

        # 3. Admin Pages (with Admin Token)
        context_admin = browser.new_context(viewport={"width": 1440, "height": 900}, device_scale_factor=2)
        context_admin.add_cookies([{
            "name": "token",
            "value": admin_token,
            "domain": "localhost",
            "path": "/"
        }])
        page_admin = context_admin.new_page()

        print("Capturing 06_admin_dashboard.png...")
        page_admin.goto("http://localhost:3000/admin/dashboard", wait_until="networkidle")
        page_admin.wait_for_timeout(1200)
        page_admin.screenshot(path=os.path.join(SCREENSHOT_DIR, "06_admin_dashboard.png"))

        print("Capturing 07_admin_laporan.png...")
        page_admin.goto("http://localhost:3000/admin/laporan", wait_until="networkidle")
        page_admin.wait_for_timeout(1200)
        page_admin.screenshot(path=os.path.join(SCREENSHOT_DIR, "07_admin_laporan.png"))

        print("Capturing 08_admin_jenis_sampah.png...")
        page_admin.goto("http://localhost:3000/admin/jenis-sampah", wait_until="networkidle")
        page_admin.wait_for_timeout(1200)
        page_admin.screenshot(path=os.path.join(SCREENSHOT_DIR, "08_admin_jenis_sampah.png"))

        print("Capturing 09_admin_wilayah.png...")
        page_admin.goto("http://localhost:3000/admin/wilayah", wait_until="networkidle")
        page_admin.wait_for_timeout(1200)
        page_admin.screenshot(path=os.path.join(SCREENSHOT_DIR, "09_admin_wilayah.png"))

        print("Capturing 10_admin_users.png...")
        page_admin.goto("http://localhost:3000/admin/users", wait_until="networkidle")
        page_admin.wait_for_timeout(1200)
        page_admin.screenshot(path=os.path.join(SCREENSHOT_DIR, "10_admin_users.png"))

        print("Capturing 11_admin_transaksi.png...")
        page_admin.goto("http://localhost:3000/admin/transaksi", wait_until="networkidle")
        page_admin.wait_for_timeout(1200)
        page_admin.screenshot(path=os.path.join(SCREENSHOT_DIR, "11_admin_transaksi.png"))
        context_admin.close()

        browser.close()
        print("ALL AUTHENTICATED SCREENSHOTS CAPTURED PERFECTLY!")

if __name__ == "__main__":
    run()
