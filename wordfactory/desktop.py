# -*- coding: utf-8 -*-
r"""桌面版入口：**一个原生窗口**（不是浏览器标签页），双击 exe 就跑这个。

用户 2026-10-08 拍板：「把这个工具箱从 web 页面撤掉，做成 exe 封装……
我要一个可以直接复制到 U 盘谁都能用的 exe 软件」，打包方式照
`水库调洪工具箱v4.exe`（PyInstaller，`console=False`）。

结构（三层原样保留，只换最外面那层壳）：

    pywebview 原生窗口（Edge WebView2 内核）
        └─ http://127.0.0.1:<随机空闲端口>/   ← 本机服务，只监听 127.0.0.1
             └─ wordfactory/gui/web/index.html（界面**一个字没改**）
                  └─ wordfactory 引擎

* **界面沿用原样**：窗口里的按钮/五页布局就是原来那套（用户："按钮和布局就沿用"）；
* **不再弹浏览器**：pywebview 用系统自带的 Edge WebView2 渲染，没有标签页/地址栏；
* **换机器能用**：规则/方案/临时文件都在 exe 旁边的 ``word工厂数据\``（`paths.py` 管），
  整个文件夹拷到 U 盘再拷出来照样跑；
* **没装 pywebview 也能用**（源码运行）：自动退回"起服务 + 开浏览器"，只是不太像桌面软件。

窗口关掉 = 服务停掉 = 进程退出；「退出」按钮走 `/api/shutdown`，之后窗口自动关闭。
"""

import os
import threading
import time


def _free_port():
    """要一个当前空闲的端口（避免固定 8765 撞上别的东西）。"""
    import socket
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]


def _watch_shutdown(server, window, delay=1.2):
    """页面上的「退出」按钮会调 ``/api/shutdown`` —— 服务停了就把窗口也关掉。

    ``server.shutdown()`` 是标准库里那个"请求停服"的标记，这里轮询它就好：
    不加这段的话，用户点了「退出」只看到"已退出，可以关掉这个窗口"，还得自己点 ×。
    """
    while not getattr(server, "_BaseServer__shutdown_request", False):
        time.sleep(0.3)
    time.sleep(delay)                      # 让页面把"已退出"那句话画出来
    try:
        window.destroy()
    except Exception:                      # noqa: BLE001 - 窗口可能已经被用户关了
        pass


def run_window(server, url, title=u"word工厂"):
    """开原生窗口指着 ``url``，窗口关掉就返回（并把服务停掉）。"""
    import webview
    window = webview.create_window(title, url, width=1180, height=820,
                                   min_size=(900, 620))
    threading.Thread(target=_watch_shutdown, args=(server, window), daemon=True).start()
    webview.start()                       # 阻塞到窗口关闭
    return window


def _browser_fallback(url, thread):
    """原生窗口起不来时的退路：**用浏览器打开同一个界面**（功能一样，只是不像桌面软件）。

    什么时候会走到这儿：目标机器上**没有 Edge WebView2 运行时**（老系统、或被精简过的系统）、
    或者 pywebview 没装。用户要的是"拷到别的电脑双击就能用"，所以**宁可退回浏览器，
    也不能只弹一句"启动失败"就完事**（2026-10-08 补）。
    """
    import webbrowser
    try:
        webbrowser.open(url)
        print(u"（原生窗口起不来，已改用浏览器打开：%s）" % url)
    except Exception as exc:                   # noqa: BLE001
        print(u"打开浏览器也失败了：%s" % exc)
    try:
        while thread.is_alive():
            time.sleep(0.3)
    except KeyboardInterrupt:
        pass
    return 0


def main(argv=None):
    from . import paths
    from .gui import server as gui_server

    written = paths.ensure_defaults()
    if written:
        print(u"首次运行：已在 %s 写出默认规则 %s" % (paths.rules_dir(), u"、".join(written)))
    print(u"word工厂（桌面版）")
    print(paths.describe())

    server, url = gui_server.make_server(host=u"127.0.0.1", port=_free_port())
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        run_window(server, url)
        return 0
    except (ImportError, OSError, RuntimeError, AttributeError) as exc:
        # 没装 pywebview / 没有 WebView2 / 窗口起不来 —— 退回浏览器，保证"还能用"
        print(u"（原生窗口没起来：%s）" % exc)
        return _browser_fallback(url, thread)
    except Exception as exc:               # noqa: BLE001 - 窗口里要给人话
        print(u"\n启动失败：%s" % exc)
        return 1
    finally:
        try:
            server.shutdown()
            server.server_close()
        except Exception:                  # noqa: BLE001
            pass


if __name__ == "__main__":
    raise SystemExit(main())
