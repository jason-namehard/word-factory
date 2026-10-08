# -*- coding: utf-8 -*-
"""跑外部命令的小工具：**永远不弹控制台窗口**。

为什么需要它（2026-10-08 用户报的致命 bug）：打包成 exe 之后进程**没有控制台**
（``console=False``），这时用 ``subprocess`` 起一个控制台程序（比如 ``taskkill``），
Windows 会给它**新开一个黑窗口**闪一下。用户点「导出 PDF」时看到"频繁弹终端窗口"，
而 PDF 其实早就导好了 —— 那几下闪窗就是我们自己收尾时调的 ``taskkill``。
（源码运行时子进程继承父进程的控制台，所以**只有 exe 上看得见**。）

用法：``subproc.run([...])`` / ``subproc.popen([...])``，参数与 ``subprocess`` 完全一样。
"""

import os
import subprocess


def _hide_console(kwargs):
    """Windows 上加 ``CREATE_NO_WINDOW``，别让子进程开新控制台。"""
    if os.name == "nt":
        kwargs.setdefault("creationflags", subprocess.CREATE_NO_WINDOW)
    return kwargs


def run(command, **kwargs):
    return subprocess.run(command, **_hide_console(kwargs))


def popen(command, **kwargs):
    return subprocess.Popen(command, **_hide_console(kwargs))
