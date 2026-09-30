# -*- coding: utf-8 -*-
"""桌面版入口：打包后的 exe 双击就跑这个。

做的是"本地服务 + 打开浏览器"这一件事（和 `python -m wordfactory.cli gui` 一样），
差别只在于：**先把数据目录准备好**（首次运行写出内置默认规则），
并且**用窗口标题说清状态**（用户要发给面试官演示，窗口里得能看懂在干什么）。

要真窗口（不弹浏览器）需要 pywebview，本轮不做 —— 先保证"一个 exe 双击就能用"。
"""

import sys
import webbrowser


def main(argv=None):
    from . import paths
    from .gui import server as gui_server

    written = paths.ensure_defaults()
    if written:
        print(u"首次运行：已在 %s 写出默认规则 %s" % (paths.rules_dir(),
                                                     u"、".join(written)))
    print(u"word 工厂（桌面版）")
    print(paths.describe())
    try:
        return gui_server.serve(host=u"127.0.0.1", port=8765, open_browser=True)
    except Exception as exc:                   # noqa: BLE001 - 窗口里要给人话
        print(u"\n启动失败：%s" % exc)
        try:
            webbrowser.open(u"about:blank")
        except Exception:
            pass
        print(u"按任意键关闭窗口…")
        try:
            input()
        except (EOFError, KeyboardInterrupt):
            pass
        return 1


if __name__ == "__main__":
    sys.exit(main())
