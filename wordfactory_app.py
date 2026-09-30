# -*- coding: utf-8 -*-
"""打包用的入口脚本（**故意放在项目根，不在包里**）。

原因：PyInstaller 打包 `wordfactory/desktop.py` 时会把它当**顶层脚本**，
包内���相对导入（``from . import paths``）就失效了（实测报
"attempted relative import with no known parent package"）。
本文件用**绝对导入**引用包内模块，PyInstaller 以它为入口就正常。
"""

import sys


def main():
    from wordfactory.desktop import main as desktop_main
    return desktop_main(sys.argv[1:])


if __name__ == "__main__":
    sys.exit(main())
