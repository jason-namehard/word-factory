# -*- coding: utf-8 -*-
"""打包成单文件 exe（给面试官演示用）。

    D:\\Hermes\\hermes-agent\\venv\\Scripts\\python.exe build_exe.py

产物：``dist/word工厂.exe``（自带 Python 运行时，别人机器上双击即可；
数据目录 ``word工厂数据\\`` 在 exe 旁边，第一次运行自动写出默认规则）。

要点：
* **单文件**（--onefile）—— 发给别人只要一个文件；
* 带上 ``wordfactory/gui/web/index.html``（界面是数据文件，不打进去 exe 就白屏）；
* pywin32 只在 PDF 导出/读页码时用，装了 Word/WPS 的机器才需要 —— 缺了不影响其它功能；
* 不打包 rules/ 目录：运行时由 ``paths.ensure_defaults()`` 写出来，版本跟着 exe 走。
"""

import os
import shutil
import subprocess
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
DIST = os.path.join(ROOT, "dist")
BUILD = os.path.join(ROOT, "build")
NAME = "word工厂"


def main():
    python = sys.executable
    if os.path.isdir(DIST):
        shutil.rmtree(DIST, ignore_errors=True)
    if os.path.isdir(BUILD):
        shutil.rmtree(BUILD, ignore_errors=True)
    command = [
        python, "-m", "PyInstaller",
        "--noconfirm", "--clean", "--onefile",
        "--name", NAME,
        "--distpath", DIST,
        "--workpath", BUILD,
        "--specpath", BUILD,
        # 界面是数据文件（**必须绝对路径**：--add-data 相对 spec 目录解析）
        "--add-data", os.path.join(ROOT, "wordfactory", "gui", "web", "index.html")
                      + ";wordfactory/gui/web",
        # 入口在项目根（包内脚本当顶层跑会丢相对导入，实测过）
        "--paths", ROOT,
        # 把包本身也打进去（onefile 下由 --paths 提供源码）
        "--hidden-import", "wordfactory",
        "--hidden-import", "wordfactory.gui.server",
        "--hidden-import", "wordfactory.frontend_paths" if False else "wordfactory.paths",
        os.path.join(ROOT, "wordfactory_app.py"),
    ]
    # 控制台窗口：演示时要能看见"服务已启动/地址是什么"，也能靠关窗口停服务
    print(u"打包中（第一次会慢一些，PyInstaller 要收集依赖）…")
    result = subprocess.run(command, cwd=ROOT)
    if result.returncode != 0:
        return result.returncode
    exe = os.path.join(DIST, NAME + ".exe")
    if not os.path.exists(exe):
        print(u"没找到产物：%s" % exe)
        return 1
    size = os.path.getsize(exe)
    print(u"\n完成：%s（%.1f MB）" % (exe, size / 1024.0 / 1024.0))
    print(u"把它拷给别人 → 双击 → 自动起服务并打开浏览器；"
          u"旁边会出现 word工厂数据\\ 目录（规则/方案/临时文件都在那儿）。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
