# -*- coding: utf-8 -*-
"""路径管理：源码跑和**打包成 exe** 都能找到/写对地方。

打包成单文件 exe 之后，程序里那些 ``<项目根>/rules/…`` 的相对路径全都不存在了
（项目根变成 PyInstaller 的临时解压目录，还会每次启动换地方）。所以统一在这里定：

* **源码运行**：规则在项目的 ``rules/``，方案在 ``~/.wordfactory/plans/``，
  临时文件在项目下的 ``临时文件/``（用户 2026-09-27 拍板：临时文件就放 wordfactory 里）；
* **exe 运行**：exe 旁边的 ``word工厂数据\`` 目录（规则/方案/临时文件都在这儿）。
  整个文件夹连同 exe 一起拷走就能换机器 —— 面试演示、换电脑都不用重配。

首次运行（数据目录不存在或缺文件）会把**内置默认规则**写出来，用户可以直接改。
"""

import os
import sys

#: 打包后数据目录的名字（exe 旁边）
DATA_DIR_NAME = u"word工厂数据"
#: 源码运行时项目根（wordfactory/ 的上一级）
SOURCE_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def is_frozen():
    """是不是打包后的 exe。"""
    return bool(getattr(sys, "frozen", False))


def app_dir():
    """程序所在目录：exe 旁边（打包后）／项目根（源码）。"""
    if is_frozen():
        return os.path.dirname(os.path.abspath(sys.executable))
    return SOURCE_ROOT


#: 记着这次进程用的是哪个数据目录（打包版可能"exe 旁边写不进 → 退回用户目录"）
_chosen_data_dir = None


def _writable(path):
    """这个目录能不能写（不存在就顺手建）。写不进去 = U 盘写保护 / 装在了 Program Files。"""
    try:
        if not os.path.isdir(path):
            os.makedirs(path)
        probe = os.path.join(path, u".wf-write-test")
        with open(probe, "wb") as handle:
            handle.write(b"x")
        os.remove(probe)
        return True
    except OSError:
        return False


def data_dir():
    """用户数据根目录（规则、方案、临时文件都放这儿）。

    打包版优先放**exe 旁边**（跟着 U 盘走，换机器不用重配）；**写不进去就退回
    ``%LOCALAPPDATA%\\word工厂``** —— 用户 2026-10-08 要的是"拷到 U 盘谁都能用"，
    真碰上写保护的 U 盘、或装在 Program Files 里，宁可规则落用户目录，
    也不能整个程序起不来（写不进去时 `*_dir()` 全是空目录/异常）。
    """
    global _chosen_data_dir
    if not is_frozen():
        path = SOURCE_ROOT
    else:
        path = _chosen_data_dir or os.path.join(app_dir(), DATA_DIR_NAME)
        if not _writable(path):
            base = os.environ.get("LOCALAPPDATA") or os.path.expanduser(u"~")
            path = os.path.join(base, u"word工厂", DATA_DIR_NAME)
            _writable(path)
        _chosen_data_dir = path
    if not os.path.isdir(path):
        try:
            os.makedirs(path)
        except OSError:
            pass
    return path


def data_dir_is_portable():
    """数据目录是不是就在程序旁边（打包版的正常情形）。"""
    return (not is_frozen()) or (data_dir() == os.path.join(app_dir(), DATA_DIR_NAME))


def rules_dir():
    """外置规则目录（字体/上下标/表格模板/替换规则）。"""
    path = os.path.join(data_dir(), "rules")
    if not os.path.isdir(path):
        try:
            os.makedirs(path)
        except OSError:
            pass
    return path


def rules_path(name):
    return os.path.join(rules_dir(), name)


def replace_rules_dir():
    """用户自建的替换规则（界面「替换规则」页写的）。"""
    path = os.path.join(rules_dir(), "replace-rules")
    if not os.path.isdir(path):
        try:
            os.makedirs(path)
        except OSError:
            pass
    return path


def plans_dir():
    """执行方案目录。打包后放在数据目录里（跟着 exe 走，便于拷贝）。"""
    if is_frozen():
        path = os.path.join(data_dir(), "plans")
    else:
        path = os.path.join(os.path.expanduser(u"~"), u".wordfactory", "plans")
    if not os.path.isdir(path):
        try:
            os.makedirs(path)
        except OSError:
            pass
    return path


def temp_dir():
    """临时文件夹（"运行此方案"的临时版落这儿；用户 2026-09-27：就放 wordfactory 里）。"""
    path = os.path.join(data_dir(), u"临时文件")
    if not os.path.isdir(path):
        try:
            os.makedirs(path)
        except OSError:
            pass
    return path


#: 内置默认规则（缺文件时写出来）：名字 → 生成内容的函数
def _default_tablestyle_json():
    from .tablestyle import DEFAULT_STYLES
    import collections
    import io
    import json
    return json.dumps(DEFAULT_STYLES, ensure_ascii=False, indent=2)


def _default_fonts_json():
    from .fonts import DEFAULT_FONTS
    import collections
    import io
    import json
    return json.dumps(DEFAULT_FONTS, ensure_ascii=False, indent=2)


def _default_subscripts_json():
    from .rules import default_ruleset
    import io
    import json
    return json.dumps(default_ruleset().to_dict()
                      if hasattr(default_ruleset(), "to_dict")
                      else {"schema": 1, "rules": []}, ensure_ascii=False, indent=2)


def _default_replacements_json():
    from .ops import textfix as textfix_op
    import collections
    import io
    import json
    return json.dumps(textfix_op.DEFAULT_REPLACEMENTS, ensure_ascii=False, indent=2)


DEFAULTS = [
    (u"tablestyle.json", _default_tablestyle_json),
    (u"fonts.json", _default_fonts_json),
    (u"subscripts.json", _default_subscripts_json),
    (u"replacements.json", _default_replacements_json),
]


def ensure_defaults(verbose=True):
    """数据目录里缺规则文件就写一份内置默认（只写缺的，不覆盖用户改过的）。

    **先生成再落盘**：以前是先 open(w) 再生成内容，生成一失败就留下 0 字节文件，
    下次启动还当它存在 —— 打包版首次运行就因此写出三个空规则（2026-09-30 实测）。
    失败**照实报出来**，不装作没事。
    """
    import io
    written = []
    problems = []
    for name, maker in DEFAULTS:
        path = rules_path(name)
        if os.path.exists(path) and os.path.getsize(path) > 0:
            continue
        try:
            content = maker()             # 先在内存里生成成功再写
        except Exception as exc:          # noqa: BLE001 - 如实报，不留半个文件
            problems.append(u"%s（%s）" % (name, exc))
            continue
        try:
            with io.open(path, "w", encoding="utf-8") as handle:
                handle.write(content)
            written.append(name)
        except OSError as exc:
            problems.append(u"%s 写不进（%s）" % (name, exc))
    if problems and verbose:
        print(u"⚠ 默认规则没能写全：%s（功能可能缺项，可从源码的 rules/ 手动拷过去）"
              % u"；".join(problems))
    return written


def describe():
    """给用户看的一行（"这是什么意思"页 / 启动时打印）。"""
    return u"程序目录：%s ｜ 数据目录：%s%s" % (
        app_dir(), data_dir(), u"（打包版）" if is_frozen() else u"（源码运行）")
