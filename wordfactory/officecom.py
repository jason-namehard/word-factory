# -*- coding: utf-8 -*-
"""Word/WPS 的 COM 会话：**每次都自己起一个私有实例，用完保证关干净**。

这个模块存在的唯一理由，是把"跟 Office 打交道"的几条硬规矩收在一处 —— 每条都是踩出来的：

1. **绝不借用用户已经开着的 Office**。旧写法先 ``GetActiveObject`` 找现成实例，看着聪明，
   实际是个坑：机器上只要有一个**自动化残留**的 Word/WPS（没有窗口、卡在某个状态），
   我们就会连上去，然后 ``Documents.Open`` 一直不返回 —— 实测整整 **60 秒超时**。
   后果是"读不到真实页码 → 退回按分页符估算"，**首页/扉页与目录之间那两只空白页就再也删不掉**
   （2026-10-08 用户报的"页码一个都不对"就是这条链）。
   ``DispatchEx`` 实测能拿到**新鲜实例**（新进程、``Version`` 秒回），所以只用它，不碰别人的。

2. **自己起的实例必须死透**。反复起、关不干净会攒下一堆没有窗口的 WINWORD/wps 进程
   （实测机器上攒了 **24 个 WINWORD + 32 个 wps**），越攒越慢，而且下一次 Dispatch
   可能就连到这些僵尸上 —— 正反馈。所以：起之前先拍一张进程快照，起完 diff 出"我们那个 pid"，
   ``Quit`` 之后再等几秒，**还活着就按 pid 结束它自己**（只动自己起的那个）。

3. **超时保护不做"每次调用套一层线程"**（2026-10-08 踩到）：COM 必须**在发起调用的那个线程里**
   ``CoInitialize``，换个线程调就是 ``尚未调用 CoInitialize``（还跨了套间）。
   所以限时放在**整件事**上（``run(timeout=…)`` 起一个工作线程 + ``join``）；
   万一真卡住，就**按 pid 杀掉我们自己起的实例** —— 进程一死，卡住的那个 COM 调用
   立刻以 RPC 错误返回，工作线程自己就结束了（不会留幽灵线程）。

对外接口：``Session(prog_id) / .open(path) / .close()`` 和 ``run(callback, prog_ids, timeout)``。
"""

import os
import threading
import time

#: 整件事（起实例 + 打开文档 + 读页码/导 PDF）的默认预算（秒）
DEFAULT_TIMEOUT = 90.0
#: 自己起的实例 ``Quit`` 之后再等它退出多久，还没走就强制结束
QUIT_GRACE = 8.0

_LOCK = threading.RLock()
#: 我们起过、还没确认退出的进程 id（超时兜底就是靠它精确清理，绝不误伤别人的 Office）
_OUR_PIDS = set()


class OfficeError(Exception):
    """起 Office / 用它干活失败（人话）。"""


def _process_ids():
    """当前所有进程 id（用 psapi，不依赖 psutil）。"""
    import ctypes
    ids = (ctypes.c_ulong * 8192)()
    size = ctypes.c_ulong()
    if ctypes.windll.psapi.EnumProcesses(ids, ctypes.sizeof(ids), ctypes.byref(size)):
        count = size.value // ctypes.sizeof(ctypes.c_ulong)
        return set(ids[:count])
    return set()


#: 只清"Office 应用本体"的进程：``DispatchEx`` 有时会连带拉起共享的辅助服务
#: （云同步之类），那些不是我们的、也不该由我们结束。
_OFFICE_EXES = ("winword.exe", "wps.exe", "et.exe", "wpp.exe", "wpspdf.exe", "excel.exe")


def _image_name(pid):
    """进程的可执行文件名（小写）；拿不到返回空串。**不调外部命令**（免得弹窗）。"""
    import ctypes
    from ctypes import wintypes
    kernel32 = ctypes.windll.kernel32
    handle = kernel32.OpenProcess(0x1000, False, int(pid))   # QUERY_LIMITED_INFORMATION
    if not handle:
        return u""
    try:
        size = wintypes.DWORD(1024)
        buffer = ctypes.create_unicode_buffer(1024)
        if kernel32.QueryFullProcessImageNameW(handle, 0, buffer, ctypes.byref(size)):
            return os.path.basename(buffer.value).lower()
    except Exception:                          # noqa: BLE001
        pass
    finally:
        kernel32.CloseHandle(handle)
    return u""


def _is_office_app(pid):
    name = _image_name(pid)
    return (not name) or (name in _OFFICE_EXES)


def _has_visible_window(pid):
    """这个进程有没有**可见的**顶层窗口？

    用户自己开着的 Word/WPS 一定有窗口；我们那些自动化实例没有 —— 收尾清理时**只清没窗口的**，
    这样绝不会误伤用户正在用的 Office。
    """
    import ctypes
    user32 = ctypes.windll.user32
    result = [False]

    def callback(hwnd, _lparam):
        if user32.IsWindowVisible(hwnd):
            owner = ctypes.c_ulong()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(owner))
            if owner.value == int(pid):
                result[0] = True
                return False
        return True

    try:
        user32.EnumWindows(ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p,
                                              ctypes.c_void_p)(callback), None)
    except Exception:                          # noqa: BLE001 - 查不了就当"有窗口"，宁可不杀
        return True
    return result[0]


def _terminate(pid):
    """按 **pid** 结束我们自己起的那个 Office 进程。

    * 只动**应用本体**（WINWORD/wps/…）：``DispatchEx`` 有时会连带拉起共享辅助服务，
      那不是我们的东西，不能顺手关掉；
    * 走 :mod:`wordfactory.subproc`：打包成 exe 之后，直接 ``subprocess`` 起 ``taskkill``
      会**闪一个黑色终端窗口**（用户 2026-10-08 报的"点导出 PDF 频繁弹窗"就是它）。
    * 绝不用 ``/IM 名字`` 那种批量杀法。
    """
    from . import subproc
    if not _is_office_app(pid):
        return
    try:
        subproc.run(["taskkill", "/PID", str(int(pid)), "/T", "/F"],
                    capture_output=True, timeout=30)
    except Exception:                          # noqa: BLE001 - 收尾失败不该再抛
        pass


def _register_ours(pids):
    with _LOCK:
        _OUR_PIDS.update(pids)


def _forget_ours(pids):
    with _LOCK:
        _OUR_PIDS.difference_update(pids)


def terminate_leftovers():
    """把"我们起过、还没退出"的实例全结束掉（超时兜底；返回结束掉的 pid 列表）。"""
    with _LOCK:
        left = set(_OUR_PIDS)
    for pid in sorted(left):
        _terminate(pid)
    _forget_ours(left)
    return sorted(left)


def _create_app(prog_id):
    """起一个**新的** Office 实例（单独抽出来是为了单测能替换掉真实 COM）。"""
    import win32com.client
    return win32com.client.DispatchEx(prog_id)


class Session:
    """一个私有 Office 实例的生命周期。用 ``with`` 或者记得 ``close()``。"""

    def __init__(self, prog_id):
        self.prog_id = prog_id
        self.pids = set()
        self._killed = set()
        self.documents = []
        before = _process_ids()
        self._before = before
        try:
            self.app = _create_app(prog_id)
            self.pids = _process_ids() - before          # 这次新起来的进程 = 我们的
            _register_ours(self.pids)                    # 先登记，超时才找得到它
            self._configure_private()
            self._health_check()
        except Exception:
            self._kill_our_processes()                   # 起坏了也别留残留
            raise

    # ------------------------------------------------------------------ 配置
    def _configure_private(self):
        """自己起的实例：藏起来、别弹窗、别执行宏。"""
        for name, value in ((u"Visible", False), (u"DisplayAlerts", 0),
                            (u"AutomationSecurity", 3)):
            try:
                setattr(self.app, name, value)
            except Exception:                # noqa: BLE001 - 老版本没这几个属性也照跑
                pass

    def _health_check(self):
        """读一个最便宜的属性：实例是坏的/没起来的话，这里就失败，不用等到打开文档。"""
        version = self.app.Version
        if not version:
            raise OfficeError(u"%s 起来了但读不到版本号" % self.prog_id)

    # ------------------------------------------------------------------ 干活
    def open(self, path):
        """只读打开（``ReadOnly=True``、不进"最近使用的文档"）。"""
        path = os.path.abspath(path)
        documents = self.app.Documents
        try:
            count = int(documents.Count)
        except Exception:                    # noqa: BLE001
            count = 0
        for index in range(1, count + 1):
            try:
                opened = str(documents.Item(index).FullName)
            except Exception:                # noqa: BLE001 - 读不出来的窗口就别管
                continue
            if os.path.normcase(opened) == os.path.normcase(path):
                raise OfficeError(u"这份报告已经在这个 Office 实例里打开着：%s" % path)
        document = documents.Open(path, False, True, False)   # 只读、不进最近文档
        if document is None:
            raise OfficeError(u"Word/WPS 没有返回已打开的文档")
        self.documents.append(document)
        return document

    # ------------------------------------------------------------------ 收尾
    def _kill(self, pid):
        """结束一个我们自己起的 pid（同一个 pid 只动手一次）。"""
        if pid in self._killed:
            return
        self._killed.add(pid)
        _forget_ours({pid})
        _terminate(pid)

    def _kill_our_processes(self):
        for pid in sorted(self.pids & _process_ids()):
            self._kill(pid)
        _forget_ours(self.pids)
        self.pids = set()

    def close(self):
        """关掉我们打开的文档 → ``Quit`` 我们起的实例 → **确认它真没了**。"""
        for document in reversed(self.documents):
            try:
                document.Close(0)            # wdDoNotSaveChanges
            except Exception:                # noqa: BLE001
                pass
        self.documents = []
        try:
            self.app.Quit(0)
        except Exception:                    # noqa: BLE001
            pass
        deadline = time.time() + QUIT_GRACE
        while time.time() < deadline:
            if not (self.pids & _process_ids()):
                break
            time.sleep(0.3)
        self._kill_our_processes()
        self._sweep_late_strays()

    def _sweep_late_strays(self):
        """最后扫一遍：**起会话之后**新冒出来、**没有可见窗口**的 Office 本体也清掉。

        为什么要这一手：WPS 有时"慢半拍"再起一个子孙进程（差集是在起会话那一刻拍的，
        抓不到它）。判据卡得很死 —— 只清**没窗口**的，用户自己开着的 Office 绝不误伤。
        """
        try:
            candidates = _process_ids() - getattr(self, "_before", set())
        except Exception:                    # noqa: BLE001
            return
        for pid in sorted(candidates):
            if not _is_office_app(pid) or _has_visible_window(pid):
                continue
            self._kill(pid)

    def __enter__(self):
        return self

    def __exit__(self, *_exc):
        self.close()
        return False


def run(callback, prog_ids, timeout=DEFAULT_TIMEOUT, per_candidate=None):
    """串行地用 Office 干一件事：依次试每个 ProgID，失败换下一个，错误原样带出来。

    ``callback(session)`` 里想干什么都行（读页码、导 PDF）。预算有两层：

    * ``per_candidate``：**单个候选**的预算。卡住就换下一个候选。
      不分开的话，一个卡住的 Word 会把整轮预算吃光、连带 WPS 那一轮也轮不到
      （2026-10-08 实测：某报告让 Word 干等 SMB 超时 40 秒以上，而 WPS 3 秒就开好了）。
    * ``timeout``：**整件事**的预算（所有候选加起来）。

    超时/失败都**按 pid 清掉我们自己起的实例**（进程一死，卡住的调用立刻以 RPC 错误返回）。
    """
    prog_ids = list(prog_ids)
    if per_candidate is None:
        per_candidate = max(15.0, float(timeout) / max(1, len(prog_ids)))
    result = {}

    def attempt(prog_id, budget):
        """试一个候选；返回 (成功?, 错误文本)。"""
        box = {}

        def worker():
            initialized = False
            session = None
            try:
                import pythoncom
                pythoncom.CoInitialize()     # COM 必须在自己这个线程里初始化
                initialized = True
                with _LOCK:                  # 锁只保护"起会话"那一小段，干活时不占锁
                    session = Session(prog_id)
                box["value"] = callback(session)
            except Exception as error:       # noqa: BLE001 - 换下一个候选
                box["error"] = u"%s: %s: %s" % (prog_id, type(error).__name__, error)
            finally:
                if session is not None:
                    try:
                        session.close()
                    except Exception:        # noqa: BLE001
                        pass
                if initialized:
                    try:
                        pythoncom.CoUninitialize()
                    except Exception:        # noqa: BLE001
                        pass

        thread = threading.Thread(target=worker, daemon=True)
        thread.start()
        thread.join(budget)
        if thread.is_alive():
            killed = terminate_leftovers()
            return False, u"%s: 超过 %.0f 秒没完成（已清掉自己起的实例 %s）" % (
                prog_id, budget, killed or u"（无）")
        if "error" in box:
            return False, box["error"]
        if "value" not in box:
            return False, u"%s: 没有返回结果" % prog_id
        result["value"] = box["value"]
        return True, u""

    errors = []
    deadline = time.time() + float(timeout)
    for prog_id in prog_ids:
        left = deadline - time.time()
        if left <= 0.1:                      # 整轮预算真的用完了，别再起新实例
            errors.append(u"%s: 整轮预算（%.0f 秒）用完，没轮到它" % (prog_id, float(timeout)))
            break
        ok, message = attempt(prog_id, min(float(per_candidate), left))
        if ok:
            return result["value"]
        errors.append(message)
    raise OfficeError(u"；".join(errors))
