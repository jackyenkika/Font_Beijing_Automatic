# login.py
import os
import sys
from playwright.sync_api import sync_playwright
import app_context
from util import log, get_system_chrome_path

LOGIN_URL =  "https://yzt.beijing.gov.cn/am/UI/Login?realm=%2F&service=bjzwService&goto=https%3A%2F%2Fyzt.beijing.gov.cn%2Fam%2Foauth2%2Fauthorize"
WELCOME_URL = "https://banshi.beijing.gov.cn/pubservice/user/yztLogin?backUrl=/proAccept/urlCheck?baseUrl=aHR0cDovL2JqdC5iZWlqaW5nLmdvdi5jbi9yZW56aGVuZy9vcGVuL2xvZ2luL2dvVXNlckxvZ2luP2NsaWVudF9pZD0yMjc3JnJlZGlyZWN0X3VyaT1odHRwOi8vY29weXJpZ2h0LmJqeHdjYmouZ292LmNuL3ovaW5kZXgmcmVzcG9uc2VfdHlwZT1jb2RlJnNjb3BlPXVzZXJfaW5mbyZzdGF0ZT07aHR0cDovL3l6dC5iZWlqaW5nLmdvdi5jbi9hbS9vYXV0aDIvYXV0aG9yaXplP3NlcnZpY2U9Ymp6d1NlcnZpY2UmcmVzcG9uc2VfdHlwZT1jb2RlJmNsaWVudF9pZD0wMDAwMjExMDBfMDkmc2NvcGU9Y24rdWlkK2lkQ2FyZE51bWJlcitleHRQcm9wZXJ0aWVzK3Jlc2VydmUzJnJlZGlyZWN0X3VyaT1odHRwJTNBLy9jb3B5cmlnaHQuYmp4d2Niai5nb3YuY24vWVpUQXV0aCUzRm1hdHRlckNvZGUlM0QwNzAwMTQyMDAwMTEwMDAwMDAwMDAwMDAwMDIwMDQ0MTAw&serverType=1^2^3^4^5^6^9&isNewEvent=1"

def do_login_and_save_state(storage_path: str)-> bool:
    chrome_path = get_system_chrome_path()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, executable_path=chrome_path)
        context = browser.new_context()
        page = context.new_page()

        try:
            page.goto(LOGIN_URL, timeout=5000)
            app_context.ui.wait_for_choice("請完成掃碼登入後，確認已進入辦事頁，點擊下方【繼續流程】",primary_text="【繼續流程】")

            context.storage_state(path=storage_path)
            log("INFO", "...登入狀態已儲存")

            page.close()
            browser.close()
            return True
        except Exception as e:
            log("ERROR", f"載入登入頁面發生錯誤: {e}")
            page.close()
            browser.close()
            return False


def is_login_state_valid(storage_path: str) -> bool:
    if not os.path.exists(storage_path):
        return False

    chrome_path = get_system_chrome_path()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, executable_path=chrome_path)  # 不顯示瀏覽器
        context = browser.new_context(storage_state=storage_path)

        try:
            page = context.new_page()
            page.goto(WELCOME_URL, timeout=5000)

            # 判斷是否被導向登入頁
            if "Login" in page.url:
                browser.close()
                return False

            browser.close()
            return True
        except Exception as e:
            browser.close()
            return False
