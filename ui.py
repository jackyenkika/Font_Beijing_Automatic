#ui.py
import tkinter as tk
from tkinter import scrolledtext,filedialog
import queue
import threading

class AppUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("著作權自動申報")
        self.root.geometry("600x600")

        # Log 區
        self.text = scrolledtext.ScrolledText(
            self.root,
            wrap=tk.WORD,
            font=("Helvetica", 12)
        )
        self.text.pack(expand=True, fill="both")
        self.text.insert(tk.END, "🚀 程式啟動中...\n請選擇要載入的資料試算表！！！\n")
        self.text.configure(state="disabled")

        # 按鈕區
        self.button_frame = tk.Frame(self.root)
        self.button_frame.pack(pady=10)
        self.button_frame.pack_forget()

        self.primary_btn = tk.Button(
            self.button_frame,
            font=("Helvetica", 12),
            command=lambda: self._on_choice("exit")
        )
        self.secondary_btn = tk.Button(
            self.button_frame,
            font=("Helvetica", 12),
            command=lambda: self._on_choice("restart")
        )

        # ===== Thread-safe 控制 =====
        self._ui_queue = queue.Queue()
        self._choice_event = threading.Event()
        self._choice_result = None

        # 主 thread 輪詢 queue
        self.root.after(50, self._process_ui_queue)

    # --------------------
    # Thread-safe log
    # --------------------
    def log(self, msg: str):
        def _log():
            self.text.configure(state="normal")
            self.text.insert(tk.END, msg + "\n")
            self.text.see(tk.END)
            self.text.configure(state="disabled")
        self.root.after(0, _log)

    # --------------------
    # Public API（給 background thread 用）
    # --------------------
    def wait_for_choice(self, msg, primary_text="確認", secondary_text=None) -> str:
        self._choice_event.clear()
        self._choice_result = None

        # 丟給主 thread 顯示 UI
        self._ui_queue.put((
            "choice",
            msg,
            primary_text,
            secondary_text
        ))

        # ⏸️ 背景 thread 等結果
        self._choice_event.wait()
        return self._choice_result

    def ask_excel_path(self) -> str | None:
        return filedialog.askopenfilename(
            title="請選擇 Excel 檔案",
            filetypes=[("Excel files", "*.xlsx")]
        )
    # --------------------
    # 主 thread UI 處理
    # --------------------
    def _process_ui_queue(self):
        try:
            while True:
                task = self._ui_queue.get_nowait()
                if task[0] == "choice":
                    _, msg, primary_text, secondary_text = task
                    self._show_choice_ui(msg, primary_text, secondary_text)
        except queue.Empty:
            pass

        self.root.after(50, self._process_ui_queue)

    def _show_choice_ui(self, msg, primary_text, secondary_text):
        # 顯示訊息
        self.text.configure(state="normal")
        self.text.insert(tk.END, msg + "\n")
        self.text.see(tk.END)
        self.text.configure(state="disabled")

        # 清空按鈕
        for w in self.button_frame.winfo_children():
            w.grid_forget()
            w.pack_forget()

        self.button_frame.pack(pady=10)

        if secondary_text is None:
            self.primary_btn.config(text=primary_text)
            self.primary_btn.pack()
        else:
            self.primary_btn.config(text=primary_text)
            self.secondary_btn.config(text=secondary_text)

            self.button_frame.columnconfigure(0, weight=1)
            self.button_frame.columnconfigure(1, weight=1)

            self.primary_btn.grid(row=0, column=0, sticky="ew", padx=10)
            self.secondary_btn.grid(row=0, column=1, sticky="ew", padx=10)

    def _on_choice(self, choice):
        self._choice_result = choice
        self.button_frame.pack_forget()
        self._choice_event.set()

    # --------------------
    def run(self):
        self.root.mainloop()

    def close(self):
        self.root.destroy()