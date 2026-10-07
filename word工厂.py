# -*- coding: utf-8 -*-
"""word工厂 —— 桌面版入口（**PyInstaller 打的就是这个文件**）。

为什么入口放在项目根、而不是 `wordfactory/` 里面：PyInstaller 会把入口脚本当**顶层脚本**
跑（不是当包的一部分），放在包目录里会丢掉相对导入（2026-09-30 实测踩过）。
这里只做一件事：把项目根塞进 `sys.path`，然后调包里的 `desktop.main()`。

源码直接跑也行：``python word工厂.py``
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from wordfactory.desktop import main   # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
