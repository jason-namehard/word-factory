# -*- coding: utf-8 -*-
"""Word/WPS 的 COM 会话：**每次都自己起一个私有实例，用完保证关干净**。

这个模块存在的理由，是把"跟 Office 打交道"的几条硬规矩收在一处。**最后两条是 2026-10-09
那次事故之后加的**（详见下面第 4 条），优先级高于前面所有便利性考虑。

1. **绝不借用用户已经开着的 Office**：旧写法先 ``GetActiveObject`` 找现成实例，机器上只要有一个
   卡住的自动化残留实例，``Documents.Open`` 就永远不返回（实测 60 秒超时）→ 读不到真实页码 →
   退回估算口径 → 空白页删不掉。只用 ``DispatchEx`` 起自己的实例。

2. **自己起的实例必须死透**：反复起、关不干净会攒下一堆没有窗口的 WINWORD/wps（实测 24+32 个）。

3. **超时保护不做"每次调用套一层线程"**：COM 必须在发起调用的线程里 ``CoInitialize``。
   限时放在整件事上（``run(timeout=…)`` 起工作线程 + ``join``）；真卡住就结束我们自己起的实例
   —— 进程一死，卡住的 COM 调用立刻以 RPC 错误返回。

4. **⚠️ 结束进程是"三道闸 + 只动自己 diff 出来的 pid"，绝不能扫全机**（2026-10-09 血的事故）：

   那次的事故链：``_is_office_app()`` 写成 ``(not name) or (name in 白名单)`` —— 把"读不到映像名"
   当成了"这是我们起的 Office"。系统/受保护进程（Secure System、Registry、csrss…）
   ``OpenProcess`` 必然被拒（error 5）→ 全部被误判成"我们的"；再叠加
   ``_sweep_late_strays()`` 拿 ``_process_ids() - before`` 当候选集（``_process_ids()`` 失败时
   返回空集 → 候选集变成**全机进程**）→ ``taskkill /T`` 连进程树一起杀 → 杀到会话宿主
   （sihost/svchost）→ **整个 Windows 外壳重建**（explorer、开始菜单、搜索、剪贴板全没了）。
   当晚同型事件 5 次。

   现在（**任何一条不满足就跳过，并记录原因**）：

   * 候选集**只能**是 ``DispatchEx`` 前后 diff 出来的 pid —— **不许**扫"全机新进程差集"，
     也不许从端口/名字反查 pid；
   * 闸① 映像名读得到、且在严格白名单里（``winword.exe`` / ``wps.exe`` …）；
   * 闸② 完整映像**路径**与本次期望一致（本次起 Word 就只认 ``winword.exe`` 所在的那个目录）；
   * 闸③ 进程**创建时间晚于**本次会话开始时间；
   * 附加：进程没有可见窗口（用户自己开着的 Office 一定有窗口 → 绝不误伤）。
   * ``_process_ids()`` 失败一律返回 ``None`` 并**放弃本轮所有清理**（绝不返回空集）；
   * 登记表**按轮**存在、随轮结束清空，不跨会话累积；
   * 真杀之前**再核一遍**三道闸，并打印 pid / 映像名 / 完整路径 / 创建时间供人工核对。

对外接口：``Session(prog_id) / .open(path) / .close()`` 和 ``run(callback, prog_ids, timeout)``。
"""

import os
import threading
import time

#: 整件事（起实例 + 打开文档 + 读页码/导 PDF）的默认预算（秒）
DEFAULT_TIMEOUT = 90.0
#: 自己起的实例 ``Quit`` 之后再等它退出多久，还没走才考虑强制结束
QUIT_GRACE = 8.0

_LOCK = threading.RLock()

#: Office 应用本体的可执行文件名 —— **严格白名单**。读不到名字 ≠ 是我们的人（事故根因）。
_OFFICE_EXES = frozenset((u"winword.exe", u"wps.exe", u"et.exe", u"wpp.exe",
                          u"wpspdf.exe", u"excel.exe"))

#: 本次会话"期望的映像名"：起 Word 就只认 winword.exe，起 WPS 就只认 wps.exe。
_EXPECTED_IMAGE = {
    u"Word.Application": u"winword.exe",
    u"KWPS.Application": u"wps.exe",
    u"WPS.Application": u"wps.exe",
}

#: 关进程前的三道闸 + "没有可见窗口"，逐条写清楚（出问题时要能对着日志说出为什么放过/为什么杀）
_QUERY_LIMITED_INFORMATION = 0x1000


class OfficeError(Exception):
    """起 Office / 用它干活失败（人话）。"""


# --------------------------------------------------------------------------- 进程信息
def _process_ids():
    """当前所有进程 id 的集合；**失败返回 None**。

    ⚠️ 绝不返回空集：调用方拿它做"差集"时，空集会把候选集放大成"全机所有进程"，
    2026-10-09 的事故就是被这个放大的。
    """
    import ctypes
    ids = (ctypes.c_ulong * 8192)()
    size = ctypes.c_ulong()
    try:
        ok = ctypes.windll.psapi.EnumProcesses(ids, ctypes.sizeof(ids), ctypes.byref(size))
    except Exception:                          # noqa: BLE001
        return None
    if not ok:
        return None
    count = size.value // ctypes.sizeof(ctypes.c_ulong)
    if count <= 0:
        return None
    return set(ids[:count])


def _process_info(pid):
    """``(映像名小写, 完整路径, 创建时间戳)`` —— 任一项读不到就给 ``""`` / ``0.0``。

    **读不到就是读不到**：调用方必须把它当成"不是我们的"，跳过。
    """
    import ctypes
    from ctypes import wintypes
    kernel32 = ctypes.windll.kernel32
    try:
        handle = kernel32.OpenProcess(_QUERY_LIMITED_INFORMATION, False, int(pid))
    except Exception:                          # noqa: BLE001
        return u"", u"", 0.0
    if not handle:
        return u"", u"", 0.0
    try:
        size = wintypes.DWORD(1024)
        buffer = ctypes.create_unicode_buffer(1024)
        path = u""
        if kernel32.QueryFullProcessImageNameW(handle, 0, buffer, ctypes.byref(size)):
            path = buffer.value or u""
        created = 0.0
        creation = wintypes.FILETIME()
        exit_time = wintypes.FILETIME()
        kernel = wintypes.FILETIME()
        user = wintypes.FILETIME()
        if kernel32.GetProcessTimes(handle, ctypes.byref(creation), ctypes.byref(exit_time),
                                    ctypes.byref(kernel), ctypes.byref(user)):
            ticks = (creation.dwHighDateTime << 32) | creation.dwLowDateTime
            if ticks:
                # FILETIME 是 1601-01-01 起的 100ns
                created = ticks / 10000000.0 - 11644473600.0
        return os.path.basename(path).lower(), path, created
    except Exception:                          # noqa: BLE001
        return u"", u"", 0.0
    finally:
        try:
            kernel32.CloseHandle(handle)
        except Exception:                      # noqa: BLE001
            pass


def _image_name(pid):
    return _process_info(pid)[0]


def _is_office_app(pid):
    """**严格白名单**：只有"名字读得到"**且**"在白名单里"才算。

    读不到名字（系统/受保护进程必然如此）一律 **False** —— 事故就出在这行原来写的
    ``(not name) or (name in _OFFICE_EXES)``，把"读不到"当成了"是我们的"。
    """
    name = _image_name(pid)
    return bool(name) and (name in _OFFICE_EXES)


def _has_visible_window(pid):
    """这个进程有没有**可见的**顶层窗口？（用户自己开着的 Office 一定有 → 不许动）"""
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


def _gate_reason(pid, prog_id, started_at, expected_dir):
    """三道闸 + 附加检查；返回 ``None`` 表示**全过**，否则返回"为什么跳过"（人话）。"""
    alive = _process_ids()
    if alive is None:
        return u"拿不到进程表（不赌，跳过）"
    if int(pid) not in alive:
        return u"进程已经退出了"
    name, path, created = _process_info(pid)
    if not name:
        return u"读不到映像名（多半是系统/受保护进程）"
    if name not in _OFFICE_EXES:
        return u"映像名 %s 不在白名单" % name
    expected_name = _EXPECTED_IMAGE.get(prog_id)
    if expected_name and name != expected_name:
        return u"映像名 %s 与本次启动的 %s 不符" % (name, expected_name)
    if not path:
        return u"读不到完整路径"
    if not expected_dir:
        return u"拿不到期望目录（无法核对路径）"
    if os.path.dirname(os.path.normcase(path)) != expected_dir:
        return u"路径不在本次 Office 的安装目录里：%s" % path
    if not created:
        return u"读不到创建时间"
    if created <= float(started_at):
        return u"创建时间早于本次会话开始（不是我起的）"
    if _has_visible_window(pid):
        return u"有可见窗口（用户自己开着的 Office，不许动）"
    return None


def _verified_pids(candidates, prog_id, started_at):
    """从候选里挑出**全过闸**的 pid；返回 ``(verified, expected_dir, skipped)``。

    ``candidates`` **只允许**是 DispatchEx 前后 diff 出来的集合。
    """
    expected_name = _EXPECTED_IMAGE.get(prog_id)
    expected_dir = u""
    for pid in sorted(candidates):
        name, path, _created = _process_info(pid)
        if name == expected_name and path:
            expected_dir = os.path.dirname(os.path.normcase(path))
            break
    verified = set()
    skipped = []
    for pid in sorted(candidates):
        reason = _gate_reason(pid, prog_id, started_at, expected_dir)
        if reason:
            skipped.append((pid, reason))
            continue
        verified.add(pid)
    return verified, expected_dir, skipped


def _terminate(pid, prog_id, started_at, expected_dir):
    """结束**已经验证过**的 pid。杀之前**再核一遍**三道闸，并打印详情供人工核对。"""
    from . import subproc
    reason = _gate_reason(pid, prog_id, started_at, expected_dir)
    if reason:
        print(u"[word工厂] 清理跳过 pid=%s：%s" % (pid, reason))
        return False
    name, path, created = _process_info(pid)
    print(u"[word工厂] 结束自己起的 Office 进程：pid=%s 映像=%s 路径=%s 创建时间=%s"
          % (pid, name, path, time.strftime(u"%Y-%m-%d %H:%M:%S", time.localtime(created))))
    try:
        # 走 subproc：exe（无控制台）里裸 subprocess 起 taskkill 会弹一个终端窗口
        subproc.run(["taskkill", "/PID", str(int(pid)), "/F"],
                    capture_output=True, timeout=30)
    except Exception:                          # noqa: BLE001 - 收尾失败不该再抛
        return False
    return True


def kill_registered(registry):
    """结束登记表里**仍然全过闸**的进程（超时兜底的唯一入口）；返回被杀 pid 列表。

    ``registry``：``{pid: (prog_id, started_at, expected_dir)}`` —— **按轮**存在，
    调用完即清空，绝不跨会话累积。
    """
    killed = []
    for pid, context in sorted(registry.items()):
        if _terminate(pid, *context):
            killed.append(pid)
    registry.clear()
    return killed


def _create_app(prog_id):
    """起一个**新的** Office 实例（单独抽出来是为了单测能替换掉真实 COM）。"""
    import win32com.client
    return win32com.client.DispatchEx(prog_id)


class Session:
    """一个私有 Office 实例的生命周期。用 ``with`` 或者记得 ``close()``。"""

    def __init__(self, prog_id, registry=None):
        self.prog_id = prog_id
        self.started_at = time.time()
        self.registry = registry if registry is not None else {}
        self.expected_dir = u""
        self.pids = set()                      # **只装"三道闸全过"的 pid**
        self.cleanup_log = []
        self._killed = set()
        self.documents = []
        before = _process_ids()
        if before is None:
            self.cleanup_log.append(u"起会话时拿不到进程快照 → 本轮不做任何清理")
        try:
            self.app = _create_app(prog_id)
            after = _process_ids()
            if before is not None and after is not None:
                verified, expected_dir, skipped = _verified_pids(
                    after - before, prog_id, self.started_at)
                self.pids = verified
                self.expected_dir = expected_dir
                for pid in verified:
                    self.registry[pid] = (prog_id, self.started_at, expected_dir)
                for pid, reason in skipped:
                    self.cleanup_log.append(u"不处理 pid=%s：%s" % (pid, reason))
            elif before is not None:
                self.cleanup_log.append(u"起完拿不到进程快照 → 本轮不做任何清理")
            self._configure_private()
            self._health_check()
        except Exception:
            self._kill_verified()              # 起坏了也别留残留（同样三道闸）
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
    def _kill_verified(self):
        """只结束**已验证**的 pid（同一 pid 只动手一次），逐个再核三道闸。"""
        for pid in sorted(self.pids):
            if pid in self._killed:
                continue
            self._killed.add(pid)
            if _terminate(pid, self.prog_id, self.started_at, self.expected_dir):
                self.registry.pop(pid, None)
        self.pids = set()

    def _wait_gone(self, grace):
        deadline = time.time() + grace
        while time.time() < deadline:
            alive = _process_ids()
            if alive is None:
                return                         # 查不到就别急着杀，宁可留残留
            if not (self.pids & alive):
                return
            time.sleep(0.3)

    def close(self):
        """关掉我们打开的文档 → ``Quit`` 我们起的实例 → **确认它真没了**。

        清理**只**处理 :attr:`pids` 里那些"diff 出来且三道闸全过"的 pid。
        这里**没有**任何"扫全机新进程"的兜底 —— 那种兜底会误伤系统进程（2026-10-09 事故）。
        """
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
        self._wait_gone(QUIT_GRACE)
        self._kill_verified()

    def __enter__(self):
        return self

    def __exit__(self, *_exc):
        self.close()
        return False


def run(callback, prog_ids, timeout=DEFAULT_TIMEOUT, per_candidate=None):
    """串行地用 Office 干一件事：依次试每个 ProgID，失败换下一个，错误原样带出来。

    ``callback(session)`` 里想干什么都行（读页码、导 PDF）。预算有两层：

    * ``per_candidate``：**单个候选**的预算。卡住就换下一个候选 —— 不然一个卡住的 Word
      会把整轮预算吃光，连 WPS 那一轮都轮不到（实测某报告让 Word 干等 SMB 超时 40 秒+）。
    * ``timeout``：**整件事**的预算（所有候选加起来）。

    超时兜底：结束**本**轮登记表里、仍全过三道闸的进程（进程一死，卡住的调用立刻以
    RPC 错误返回）。登记表只在这一轮里存在，用完即清。
    """
    prog_ids = list(prog_ids)
    if per_candidate is None:
        per_candidate = max(15.0, float(timeout) / max(1, len(prog_ids)))
    result = {}
    registry = {}                          # **按轮**的登记表：{pid: (prog_id, started_at, dir)}

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
                    session = Session(prog_id, registry)
                box["value"] = callback(session)
            except Exception as error:       # noqa: BLE001 - 换下一个候选
                box["error"] = u"%s: %s: %s" % (prog_id, type(error).__name__, error)
                box["cleanup_log"] = getattr(session, "cleanup_log", []) if session else []
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
            killed = kill_registered(registry)     # 只动登记表里全过闸的那些
            return False, u"%s: 超过 %.0f 秒没完成（已结束自己起的实例 %s）" % (
                prog_id, budget, killed or u"（无）")
        if "error" in box:
            return False, box["error"]
        if "value" not in box:
            return False, u"%s: 没有返回结果" % prog_id
        result["value"] = box["value"]
        return True, u""

    try:
        errors = []
        deadline = time.time() + float(timeout)
        for prog_id in prog_ids:
            left = deadline - time.time()
            if left <= 0.1:                  # 整轮预算真的用完了，别再起新实例
                errors.append(u"%s: 整轮预算（%.0f 秒）用完，没轮到它" % (prog_id, float(timeout)))
                break
            ok, message = attempt(prog_id, min(float(per_candidate), left))
            if ok:
                return result["value"]
            errors.append(message)
        raise OfficeError(u"；".join(errors))
    finally:
        registry.clear()                     # 登记表不跨轮存活
