# -*- mode: python ; coding: utf-8 -*-
r"""PyInstaller 打包配置：**word工厂**（桌面版，单文件 `word工厂.exe`）。

把依赖一起塞进 exe，不再出 `_internal\` 目录 —— 用户要的"一个 exe 拷到 U 盘就能用"。
（早先另有一份 onedir 的 `build_exe.spec`／`dist\word工厂\`，2026-10-17 按
"只保留最新版本"清掉了，所以这里是**唯一**的打包配置。）

* 代价：每次启动要把依赖解到临时目录（首次启动慢几秒）；
* **数据目录仍在 exe 旁边**：`paths.py` 用的是 ``sys.executable`` 的目录（不是解压目录），
  所以 `word工厂数据\` 会出现在**用户把 exe 放的那个文件夹**里，不会随退出消失。

用法（项目根）：双击 `打包 word工厂.bat`，或
    .buildenv\Scripts\python.exe -m PyInstaller build_exe_onefile.spec --noconfirm ^
        --distpath dist_onefile --workpath build_onefile
"""

import os

ROOT = SPECPATH

a = Analysis(
    ["word工厂.py"],
    pathex=[ROOT],
    binaries=[],
    datas=[
        (os.path.join(ROOT, "wordfactory", "gui", "web", "index.html"),
         "wordfactory/gui/web"),
        (os.path.join(ROOT, "wordfactory", "gui", "web", "wordfactory.ico"),
         "wordfactory/gui/web"),
    ],
    hiddenimports=[
        "webview.platforms.edgechromium",
        "clr_loader",
        "pythonnet",
        # Word/WPS 的 COM（读真实页码 / 导 PDF）：pythoncom 与 win32com 都在函数里 import，
        # 静态分析扫不到 —— 漏了的话 exe 里"读真实页码"静默失效（空白页就删不掉）
        "pythoncom",
        "pywintypes",
        "win32com",
        "win32com.client",
        "win32com.client.dynamic",
        # 这两个是**函数里 import** 的（pageprobe/pdf 里 、
        # 方案导入导出里 ）—— 显式点名，别让静态分析漏掉
        "wordfactory.officecom",
        "wordfactory.planbundle",
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=["tkinter", "unittest", "pydoc_data"],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="word工厂",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,
    icon=os.path.join(ROOT, "wordfactory", "gui", "web", "wordfactory.ico"),
)
