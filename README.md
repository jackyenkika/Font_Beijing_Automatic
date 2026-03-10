# Font_Beijing_Automatic
北京產權局自動化點擊工具


Step 1️⃣ 安裝環境（只要一次）

1. 確認你有 Python（下為macOS環境的操作流程）

若是終端機操作，會是bash區域：

bash
```
python3 --version
```

2. 建立資料夾

bash
```
mkdir bj_copyright_auto
cd bj_copyright_auto
```

3. 建立虛擬環境

bash
```
python3 -m venv venv_arm64
source venv_arm64/bin/activate
```

4. 安裝 Playwright

bash
```
pip install playwright
pip install pandas openpyxl
pip install requests
playwright install
```

Step 2️⃣ 第一支程式
將 step_start.py 貼到安裝環境下的bj_copyright_auto 資料夾中
按照內容提示文字，執行對應步驟

獲取網頁cookie
bash
```
python step_start.py
```

Step 3 自動化流程
在流程開始前，請先確認本日cookie是否已經獲取，如果確定這個cookie一年內沒被其他人使用的話，可以不需要操作步驟2，反之需要執行獲取cookie功能。
獲取過程，會需要掃描登入，完成後會存成login_state.json 的cookie檔案，裡面有過期日期可以確認。

請先將此次欲申請的各式資料填妥於bj_copyright_auto_works.xlsx。
各列對應批次檔會被執行幾個申請案件，各列代表每個申請需要填寫的內容，請確實填寫，以利後續作業。

自動化流程
bash
```
python storage_login_state.py
```


Playwright 主要做的是人的動作

打字:fill()
page.fill("#username", "your_account")

點擊:click()
page.click("#loginBtn")

看畫面:wait_for_selector()

抓取
id > name > text > xpath（最後手段）

＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
建立一般的執行檔
pyinstaller \
  --windowed \
  --onedir \
  --name "KIKA北京著作權自動申報" \
  app.py


產生spec
pyinstaller app.py --name "KIKA北京著作權自動申報" --windowed --onedir

藉由spec產生執行檔
pyinstaller KIKA北京著作權自動申報.spec

KIKA北京著作權自動申報.spec

# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='KIKA北京著作權自動申報',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='KIKA北京著作權自動申報',
)
app = BUNDLE(
    coll,
    name='KIKA北京著作權自動申報.app',
    icon='assets/app.icns',
    bundle_identifier='com.kika.copyright',
)











