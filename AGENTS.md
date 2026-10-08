# AGENTS.md — word-factory（`E:\Zspace\projects\word-factory`）

> 本文件是**本项目专属**的规则补充；跨工作区的红线与工作准则在用户级
> `C:\Users\18019\.zcode\AGENTS.md`（本文件不重复，只写本项目踩出来的、必须常驻的条目）。

## 一、进程与杀进程（**2026-10-09 事故红线**；强度=MUST/NEVER，**优先于一切功能开发**）

> **事故经过**：2026-10-09 02:01，本项目的运行时（`wordfactory` 的 python 进程）
> 连续起了 10 个 `taskkill /PID <n> /T /F`，**把 Windows 会话宿主（sihost / svchost）
> 连进程树一起杀掉 → 整个外壳重建**（explorer、开始菜单、搜索、剪贴板用户服务全没了；
> 当晚同型事件 5 次）。
>
> **两条直接原因**（都在 `wordfactory/officecom.py`）：
> 1. `_is_office_app()` 写的是 `return (not name) or (name in _OFFICE_EXES)` ——
>    **把"读不到映像名"当成了"这是我们起的 Office"**。系统/受保护进程
>    （Secure System、Registry、csrss…）`OpenProcess(QUERY_LIMITED_INFORMATION)`
>    必然被拒（error 5），于是全被误判成"我们的"；
> 2. `_sweep_late_strays()` 拿 `_process_ids() - before` 当候选集兜底，而
>    `_process_ids()` 在 `EnumProcesses` 失败时返回**空集** → 候选集变成"全机所有进程"。
> 再叠加 `taskkill /T`（连子孙一起杀），就杀到了会话宿主。
>
> **证据**：管理员导出的 4688 审计 `E:\WBspace\2026-10-09-01-20-57\_diag\audit_dump.txt`。

### 三条硬规矩（不可协商）

1. **【NEVER】禁止用文本解析（`netstat` / `tasklist` / `wmic` 的输出）反查 PID。**
   要"哪个进程占了这个端口"，用
   `Get-NetTCPConnection -LocalPort <port> -State Listen | Select-Object -ExpandProperty OwningProcess`。
2. **【MUST】任何 `taskkill` / `Stop-Process` / `TerminateProcess` 只允许作用于
   本次操作自己启动、且已核验的 PID** —— 核验三要素：①映像名在严格白名单里
   （`winword.exe` / `wps.exe`…）；②完整映像路径与本次启动的那个一致；③进程创建时间
   晚于本次会话开始时间。**【NEVER】对"从端口、名称、差集反查出来的 PID"直接强杀**；
   **【NEVER】用 `taskkill /T`**（连子孙整棵树一起杀 —— 正是它端掉了会话宿主）。
   确需结束进程时**先打印** 被杀 PID / 映像名 / 完整路径 / 创建时间，供人工核对。
3. **【NEVER】清理类代码禁止用 `_process_ids() - before` 这类"全进程差集"当候选集。**
   凡"读不到进程信息"，**一律视为"不是我们的"并跳过**；进程快照拿不到就**放弃本轮清理**。
   **绝不允许把"读不到"或"空集"当成"全部都是我们的"。**

### 落地位置（改代码时对照）

* `wordfactory/officecom.py`：`_is_office_app`（严格白名单）、`_gate_reason`（三道闸）、
  `_verified_pids`（候选只来自 diff）、`_terminate`（杀前复核 + 打印）、`kill_registered`
  （按轮登记表，用完即清）；**没有**任何"扫全机"的兜底函数。
* `wordfactory/subproc.py`：所有外部命令（含 `taskkill`）都走它，带 `CREATE_NO_WINDOW`
  （exe 无控制台时裸 `subprocess` 会弹终端窗口，用户 2026-10-08 报过）。
* 单测 `tests/test_officecom.py`：钉住"读不到名字不杀 / 白名单外不杀 / 路径不符不杀 /
  比会话还老不杀 / 有窗口不杀 / 全过闸才杀"，**全程打桩，不真的调 taskkill**。

## 二、其他常驻条目

* **【MUST】改完必须实测**：本项目的"实测"= 真跑一遍 + 用 Word/WPS 打开产物读值
  （XML 绿 ≠ 渲染对，内存改过 ≠ 落盘正确）；exe 相关的改动**必须在打包后的 exe 里再验一遍**
  （很多坑只在无控制台/冻结环境暴露）。
* 交付形态：桌面版 exe 优先；重建双击 `打包 word工厂.bat`；
  数据目录在 exe 旁边的 `word工厂数据\`。
