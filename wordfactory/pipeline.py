# -*- coding: utf-8 -*-
"""配方（pipeline）：把若干宏**按顺序**跑完，最后按模式收尾 —— GUI 和 CLI 共用这一套。

用户 2026-09-21 就定的形态：「新的工具箱也要可以选定某些修改，且排序后一件执行，要不也叫宏算了」。

**两步结构**（2026-09-23 出 GUI 时定下来）：

1. **配方** = 一串内容宏（题注 / 上下标 / 表格清理 / 文本替换 / Markdown / 一键整理），
   **顺序由用户排**（GUI 里拖排序；JSON 里就是数组顺序）；
2. **收尾**按模式（用户 2026-09-21 的两版输出口径）：
   * ``verify`` 预览版（用户 2026-09-27 起，旧叫"验证版"）：把**改动的地方标蓝**，你打开核对；
   * ``formal`` 正式版：**通体黑 + 字体合规**，然后自动体检（`audit`），末行 AUDIT=PASS。

所以「字体/颜色」不是配方里的一步，而是**收尾** —— 否则验证版会被"全部改黑"把标蓝盖掉。

CLI 与 GUI 都调 :func:`run_pipeline`，两条路行为必然一致（有单测钉）。
"""

import collections
import os

from . import fonts as fonts_mod
from . import mark as mark_mod
from . import rules as rules_mod
from .audit import audit as audit_op
from .document import Document
from .ops import captions as captions_op
from .ops import mdclean as mdclean_op
from .ops import table_clean as table_clean_op
from .ops import textfix as textfix_op
from .ops import tidy as tidy_op

#: 模式：verify = 预览版（改动标蓝，旧叫"验证版"）；formal = 正式版（通体黑 + 字体合规）
MODES = ("verify", "formal")

#: 配方里可选的步骤：名字 → (说明, 默认参数)。加新宏只在这里加一行。
STEPS = collections.OrderedDict([
    (u"captions", (u"题注格式统一（表题规则 B + 图题整段居中）", {
        "center_table": True, "center_figure": True, "figure_space": 1,
        "normalize_number": True, "only_before_table": True})),
    (u"sup", (u"上下标规则（外置 JSON）", {"rules": None})),
    (u"tableclean", (u"表格去空格/回车", {"level": 3, "merge_lines": False,
                                       "remove_spaces": False})),
    (u"tablestyle", (u"表格模板（套用款式；支持映射与通篇一致）",
                     {"style": u"通用款·外粗内细", "uniform": False, "tables": u"all",
                      "mapping": None, "wrap_header": None})),
    # textfix（文本替换+对齐）2026-09-27 起从功能清单撤下：有了可自定义的
    # 「替换规则」就不再需要它（用户原话："方案调用的功能就不需要'文本替换对齐'了"）。
    # 命令行 `textfix` 命令仍在；老方案 JSON 里挂着的 textfix 跑的时候会明说。
    (u"replace", (u"替换规则（文本/字体/段落，三段合一）", {"rules": None, "scope": "all"})),
    (u"mdclean", (u"Markdown 标记清理", {})),
    (u"tidy", (u"一键整理（删空白行 + 空白页 + 段尾空格）", {"blank_lines": True, "trailing_spaces": True})),
])

#: 界面上的**中文显示名**（用户 2026-09-27：可选功能要用中文命名 + 中文简介）。
#: 内部代号保持英文（执行方案 JSON、命令行 --steps 都用它），界面显示用这张表。
STEP_LABELS = {
    u"captions": u"题注统一",
    u"sup": u"上下标规则",
    u"tableclean": u"表格清理",
    u"tablestyle": u"表格模板",
    u"replace": u"替换规则",
    u"mdclean": u"Markdown 清理",
    u"tidy": u"一键整理",
}

#: 只有这些步骤的报告里带"可以标蓝的区间"
MARKABLE = (u"captions",)


class PipelineError(Exception):
    """配方或运行出的问题（人话）。"""


def normalize_recipe(steps):
    """把用户给的步骤清单规整成 ``[{op, params}, …]``，顺手校验。

    接受三种写法：``"captions"``（用默认参数）、``{"op": "captions"}``、
    ``{"op": "captions", "params": {...}}``。顺序就是数组顺序 —— 用户要的"排序"。
    """
    normalized = []
    for index, step in enumerate(steps or [], start=1):
        if isinstance(step, str):
            name, params = step, {}
        elif isinstance(step, dict):
            name, params = step.get("op") or step.get("name"), step.get("params") or {}
        else:
            raise PipelineError(u"第 %d 步看不懂：%r" % (index, step))
        if name not in STEPS:
            raise PipelineError(u"第 %d 步的 %r 不是已注册的宏（有的是：%s）"
                                % (index, name, u"、".join(STEPS)))
        merged = dict(STEPS[name][1])
        merged.update(params or {})
        normalized.append({"op": name, "params": merged})
    return normalized


def run_pipeline(source_path, steps, mode="verify", out_path=None, dry_run=False,
                 progress=None, font_rules=None, block_pages=None):
    """按配方跑一遍。返回报告 dict。

    * ``source_path``：输入的 .docx（**不会被改动**，我们在副本上干活）
    * ``steps``：见 :func:`normalize_recipe`（可以是空列表 —— 那就只做收尾）
    * ``mode``：``verify`` / ``formal``
    * ``progress``：可选回调 ``(步骤序号, 步骤名, 该步报告)``，GUI 用它边跑边显示
    * ``block_pages``：可选**真实页码**（``pageprobe`` 用 Word/WPS 排版后读的"第几块在第几页"，
      与 ``list(document.body())`` 一一对应）。给了一键整理的空白页判定才认得出自然
      溢出的页边界；不给就退回分页记号估算。
    """
    recipe = normalize_recipe(steps)
    if mode not in MODES:
        raise PipelineError(u"mode 只能是 %s（现在是 %r）" % (u"/".join(MODES), mode))
    if not os.path.exists(source_path):
        raise PipelineError(u"文件不存在：%s" % source_path)
    font_rules = font_rules or fonts_mod.FontRuleSet(fonts_mod.DEFAULT_FONTS)

    report = {"op": "pipeline", "file": source_path, "mode": mode,
              "steps": [], "finish": None, "dry_run": bool(dry_run),
              "audit": None, "out": None}
    marked = 0
    with Document(source_path, writable_parts=fonts_mod.FONT_PARTS) as doc:
        report["accurate_pages"] = bool(block_pages)
        for index, step in enumerate(recipe, start=1):
            name = step["op"]
            step_report = _run_step(doc, name, step["params"], dry_run, block_pages)
            if not dry_run and mode == "verify" and name in MARKABLE:
                marked += mark_mod.verify(doc, step_report.get("details") or [])
            report["steps"].append({"step": index, "op": name,
                                    "note": STEPS[name][0], "report": step_report})
            if progress:
                progress(index, name, step_report)
        if dry_run:
            report["finish"] = u"（dry-run 不收尾）"
            return report

        changed = bool(marked) or any(
            step["report"].get("changed") or step["report"].get("total")
            or step["report"].get("touched") for step in report["steps"])
        if mode == "formal":
            finish_report = fonts_mod.normalize(doc, font_rules,
                                                parts=list(fonts_mod.FONT_PARTS))
            report["finish"] = u"通体黑 + 字体合规 + 体检"
        else:
            finish_report = {"marked": marked}
            report["finish"] = u"改动标蓝（%d 处）" % marked
        # **收尾本身也算改动**：不然"配方是空、只收尾"时会因为 changed=False 而不写文件
        finished = bool(finish_report.get("fonts") or finish_report.get("colors")
                        or finish_report.get("highlights") or marked)
        if not changed and not finished:
            report["finish"] = u"没有需要改的地方"
            return report
        if not out_path:
            out_path = _default_out(source_path, mode)
        doc.save(out_path)
        report["out"] = out_path
        report["finish_report"] = finish_report
        if mode == "formal":
            report["audit"] = audit_op(out_path, font_rules)
    return report


def _probe_block_pages(document):
    """用 Word/WPS 排一遍版，读"第几块在第几页"（一键整理识别空白页要用）。

    **读不到不是错误**（没装 Word/WPS、弹窗卡住、超时）—— 返回 None，
    调用方退回"按分页记号估算"口径。工具的核心承诺是"不依赖 Word 也能跑"，
    真实页码只是**锦上添花**的那一档精度。
    """
    try:
        from . import pageprobe
    except Exception:                              # noqa: BLE001 - 缺依赖也不该炸
        return None
    path = getattr(document.package, "path", None)
    if not path or not os.path.exists(path):
        return None
    blocks = list(document.body())
    try:
        info = pageprobe.probe(path, blocks=len(blocks) + 5,
                               block_starts=pageprobe.block_start_paragraphs(
                                   document, limit=len(blocks) + 5))
    except pageprobe.PageProbeError:
        return None
    return info.get("block_pages")


def _run_step(doc, name, params, dry_run, block_pages=None):
    """跑一步。每个分支返回一个小报告（含"改了多少"这类数字）。"""
    if name == u"captions":
        return captions_op.apply(doc, params, dry_run=dry_run)

    if name == u"sup":
        root = doc.part()
        if dry_run:
            import copy
            root = copy.deepcopy(root)     # dry-run 不许动真树（apply_to_part 只有"直接改"一条路）
        result = rules_mod.apply_to_part(root, _rules(params.get("rules")), limit=10)
        result["changed"] = result["total"]
        return result

    if name == u"tableclean":
        options = {key: params[key] for key in ("level", "merge_lines", "remove_spaces",
                                                "caption_skip", "full_width_space")
                   if key in params}
        return table_clean_op.clean(doc, options, dry_run=dry_run)

    if name == u"tablestyle":
        from . import tablestyle as tablestyle_mod
        spec = {key: params[key] for key in ("style", "uniform", "tables",
                                             "mapping", "wrap_header")
                if params.get(key) not in (None, False)}
        if "uniform" in params:
            spec["uniform"] = bool(params.get("uniform"))
        return tablestyle_mod.apply_plan(doc, spec, dry_run=dry_run)

    if name == u"replace":
        from . import replace_rules as replace_rules_mod
        if not params.get("rules"):
            return {"op": "replace", "name": u"", "text": 0, "font": 0, "para": 0,
                    "total": 0, "dry_run": bool(dry_run),
                    "note": u"没指定替换规则 —— 在「替换规则」页做好、在执行方案里选中"}
        rules = replace_rules_mod.load(_replace_rules_path(params["rules"]))
        return replace_rules_mod.apply(doc, rules, dry_run=dry_run,
                                       scope=params.get("scope") or "all")

    if name == u"textfix":
        options = {}
        if params.get("replace") is False:
            options["rule_set"] = textfix_op.ReplacementRuleSet(
                {"replacements": [], "align": textfix_op.DEFAULT_REPLACEMENTS["align"]})
        if params.get("align") is False:
            options["fix_align"] = False
        return textfix_op.apply(doc, dry_run=dry_run, **options)

    if name == u"mdclean":
        return mdclean_op.apply(doc, dry_run=dry_run)

    if name == u"tidy":
        options = {key: params[key] for key in ("blank_lines", "trailing_spaces",
                                                "trim_leading", "collapse_space_runs",
                                                "caption_skip", "merge_lines",
                                                "remove_spaces", "scope",
                                                "protect_frontmatter",
                                                "frontmatter_pages",
                                                "blank_pages")
                   if key in params}
        if block_pages:
            options["block_pages"] = block_pages
        return tidy_op.tidy(doc, options, dry_run=dry_run)

    raise PipelineError(u"步骤 %s 没有实现" % name)


def _replace_rules_path(name_or_path):
    """替换规则的参数：给名字就在 rules/replace-rules/ 下找，给路径就直接用。"""
    import os as _os
    if os.path.isabs(name_or_path) or _os.path.exists(name_or_path):
        return name_or_path
    from . import paths as _paths
    base = _paths.rules_dir()
    candidate = _os.path.join(_paths.replace_rules_dir(), u"%s.json" % name_or_path)
    if _os.path.exists(candidate):
        return candidate
    candidate2 = _os.path.join(base, u"%s.json" % name_or_path)
    if _os.path.exists(candidate2):
        return candidate2
    return name_or_path


def _rules(path):
    if not path or not os.path.exists(path):
        return rules_mod.default_ruleset()
    return rules_mod.RuleSet.load(path)


def default_out(source_path, mode, suffix=None):
    """算出这次该写到哪个文件（不检查存在性）。

    ``suffix=None`` 用默认后缀（``（预览版）/（正式版）``）；给了就用它
    （用户界面上"改后缀另存"时传的）。
    """
    stem, ext = os.path.splitext(os.path.basename(source_path))
    if suffix is None:
        suffix = u"（预览版）" if mode == "verify" else u"（正式版）"
    parent = os.path.dirname(os.path.abspath(source_path))
    return os.path.join(parent, u"%s%s%s" % (stem, suffix, ext))


def _default_out(source_path, mode):
    """输出文件名：默认后缀；**已存在就加序号**（命令行/内部兜底，绝不覆盖用户的东西）。

    界面上的主路径是"先问一句"（server 侧 conflict → 用户选覆盖或改后缀），
    这里只是没有界面时的安全默认。
    """
    stem, ext = os.path.splitext(os.path.basename(source_path))
    suffix = u"（预览版）" if mode == "verify" else u"（正式版）"
    parent = os.path.dirname(os.path.abspath(source_path))
    candidate = default_out(source_path, mode)
    index = 2
    while os.path.exists(candidate):
        candidate = os.path.join(parent, u"%s%s%d%s" % (stem, suffix, index, ext))
        index += 1
    return candidate


def format_report(report):
    """人看的运行报告（GUI 直接显示这段）。"""
    mode_text = (u"预览版（改动的地方标蓝，供你核对）" if report["mode"] == "verify"
                 else u"正式版（通体黑 + 字体合规）")
    lines = [u"文件：%s" % report["file"],
             u"模式：%s" % mode_text,
             u"配方：%d 步 —— %s" % (len(report["steps"]),
                                     u" → ".join(step["op"] for step in report["steps"])
                                     or u"（空，只收尾）"),
             u""]
    for step in report["steps"]:
        lines.append(u"  · 第 %d 步 %-10s %s"
                     % (step["step"], step["op"], _summarize(step["report"])))
    if report["dry_run"]:
        lines.insert(0, u"--dry-run：一个字节都没写")
        return u"\n".join(lines)
    if report.get("out"):
        finish = report.get("finish_report") or {}
        lines.append(u"")
        lines.append(u"收尾：%s" % report.get("finish"))
        if finish.get("fonts"):
            for key in sorted(finish["fonts"]):
                lines.append(u"    %-32s %d 处" % (key, finish["fonts"][key]))
        if finish.get("colors"):
            lines.append(u"    颜色改正 %d ｜ 去高亮 %d" % (finish["colors"],
                                                         finish.get("highlights", 0)))
        lines.append(u"")
        lines.append(u"已写出：%s" % report["out"])
        if report.get("audit"):
            if report["audit"]["verdict"] == "PASS":
                lines.append(u"体检：AUDIT=PASS（通体黑、没有不合格的字体）")
            else:
                lines.append(u"体检：AUDIT=FAIL %s" % u"；".join(report["audit"]["reasons"]))
    else:
        lines.append(u"")
        lines.append(u"没有需要改的地方（也没写文件）")
    return u"\n".join(lines)


def _summarize(step_report):
    bits = []
    for key in ("changed", "total", "touched", "marked", "chars_removed"):
        if step_report.get(key):
            bits.append(u"%s %d" % (key, step_report[key]))
    changes = step_report.get("changes") or {}
    # **空白页的结果优先显示**（用户 2026-10-08 点名要看"空白页检索"的效果；
    # 按名字排序它们会排在后面被 [:3] 截掉，所以单独拎出来）
    for key in (u"空白页数", u"空白页删段"):
        if changes.get(key):
            bits.append(u"%s %d" % (key, changes[key]))
    rest = [(key, value) for key, value in sorted(changes.items())
            if key not in (u"空白页数", u"空白页删段")]
    if rest:
        bits.append(u"、".join(u"%s %d" % (key, value) for key, value in rest[:4]))
    return u"、".join(bits) if bits else u"0 处"
