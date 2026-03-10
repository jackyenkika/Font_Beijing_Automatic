#util.py
import app_context
import os
import sys

def log(stage, msg):
    """將訊息輸出到 UI"""
    text = f"[{stage}] {msg}"
    if hasattr(app_context, "ui") and app_context.ui is not None:
        app_context.ui.log(text)
    else:
        # 如果 UI 尚未初始化，就輸出到終端
        print(text)

# 系統 Chrome 路徑
def get_system_chrome_path():
    if sys.platform == "win32":
        # Windows 預設路徑
        paths = [
            os.path.join(os.environ.get("PROGRAMFILES", ""), "Google/Chrome/Application/chrome.exe"),
            os.path.join(os.environ.get("PROGRAMFILES(X86)", ""), "Google/Chrome/Application/chrome.exe"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google/Chrome/Application/chrome.exe"),
        ]
    elif sys.platform == "darwin":
        # macOS
        paths = ["/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"]
    else:
        # Linux
        paths = ["/usr/bin/google-chrome", "/usr/bin/chromium-browser", "/usr/bin/chrome"]

    for path in paths:
        if os.path.exists(path):
            return path
    raise RuntimeError("找不到系統 Chrome，請先安裝 Chrome 或 Edge")

def get_app_data_dir(app_name="KIKA北京著作權自動申報"):
    if sys.platform == "win32":
        base = os.environ.get("LOCALAPPDATA")
    elif sys.platform == "darwin":
        base = os.path.expanduser("~/Library/Application Support")
    else:
        base = os.path.expanduser("~/.local/share")

    path = os.path.join(base, app_name)
    os.makedirs(path, exist_ok=True)
    return path