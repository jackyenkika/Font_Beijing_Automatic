# app.py
import login
import threading
import app_context
from auto_workflow import run_workflow
from ui import AppUI
from util import get_app_data_dir
import os

APP_DATA_DIR = get_app_data_dir()
LOGIN_STATE = os.path.join(APP_DATA_DIR, "login_state.json")
# EXCEL_PATH = "bj_copyright_auto_works.xlsx"

def workflow_loop():
    ui = app_context.ui

    # 👉 請使用者選 Excel
    # ui.log("請選擇要載入的資料試算表！！！")
    excel_path = ui.ask_excel_path()

    if not excel_path:
        ui.log("❌ 未選擇 Excel，流程中止")
        ui.close()
        return

    while True:
        ui.log("==========================================\n🔄 開始新一輪流程")

        try:
            need_login = not login.is_login_state_valid(LOGIN_STATE)
            if need_login:
                ui.log("⚠️ Cookie 無效，需要重新登入")
                is_login = login.do_login_and_save_state(LOGIN_STATE)
            else:
                ui.log("✅ 使用既有登入狀態")
                is_login = True

            if is_login:
                run_workflow(LOGIN_STATE,excel_path)

            choice = ui.wait_for_choice(
                "自動化流程已完成\n\n請確認所有操作結果後選擇下一步",
                primary_text="【結束流程】",
                secondary_text="🔁【重新開始】"
            )

            if choice == "exit":
                ui.log("👋 使用者選擇結束流程")
                ui.close()
                break
            elif choice == "restart":
                ui.log("🔁 使用者選擇重新開始\n")
        except Exception as e:
            ui.log(f"❌ 失敗: {e}")
            break


def main():
    app_context.ui = AppUI()

    # 將 workflow 放到另一個 thread 避免阻塞 mainloop
    threading.Thread(target=workflow_loop, daemon=True).start()

    # 啟動 Tkinter 主迴圈
    app_context.ui.run()


if __name__ == "__main__":
    main()