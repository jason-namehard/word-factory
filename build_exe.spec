# -*- mode: python ; coding: utf-8 -*-
r"""PyInstaller 打包配置：**word工厂**（桌面版，原生窗口）。

打包方式照用户的参照物 `水库调洪工具箱v4`：PyInstaller、``console=False``、
**onedir**（exe + `_internal\\` 同目录）—— 整个 `dist\\word工厂\\` 文件夹拷到 U 盘就能用。

用法（项目根）：
    .buildenv\\Scripts\\python.exe -m PyInstaller build_exe.spec --noconfirm ^
        --distpath dist --workpath build
或直接双击 `打包 word工厂.bat`。
"""

import os

ROOT = SPECPATH                              # spec 所在目录 = 项目根

a = Analysis(
    ["word工厂.py"],
    pathex=[ROOT],
    binaries=[],
    datas=[
        # 界面（静态页）—— 引擎只需要这一个数据文件，规则都是代码内置生成的
        (os.path.join(ROOT, "wordfactory", "gui", "web", "index.html"),
         "wordfactory/gui/web"),
    ],
    hiddenimports=[
        # pywebview 在 Windows 上走 Edge WebView2，这几层是运行时才 import 的，静态分析看不到
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
    [],
    exclude_binaries=True,
    name="word工厂",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                 # 不用 upx：把 WebView2 的 .NET 依赖压了容易起不来
    console=False,             # GUI 程序，无控制台黑窗
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="word工厂",
)
