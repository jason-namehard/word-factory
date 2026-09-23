# -*- coding: utf-8 -*-
"""宏「格式规范化」：一条命令把"通体黑 + 去高亮 + 字体合规 + 上下标规则"跑完。

参考宏 `格式规范化.bas`（§2.7）做四件事，本命令对应的关系：

| 宏做的 | 本工具 |
|---|---|
| 全文变黑、去高亮 | ✅ `fonts.normalize`（还顺手按规则换字体） |
| `m` 数字等上标 | ✅ `rules.apply_to_part`（外置的上下标规则文件） |
| 更新目录页码 | ❌ **做不到**（页码是排版结果，见 `PLAN.md` §6） |
| 询问"保存并关闭文档" | ❌ **不做**（工具只管改；存哪、关不关是调用方的事，见 `PLAN.md` §8.3） |

所以它就是**一串已有步骤**：这也是 M4「配方 = 勾选 + 排序 + 一键执行」的雏形
（先做成一条固定顺序的命令，之后再做可编排的）。
"""

import collections

from ..fonts import DEFAULT_FONTS, FONT_PARTS, FontRuleSet, normalize as normalize_fonts
from ..rules import RuleSet, apply_to_part


def format_normalize(document, font_rules=None, text_rules=None, dry_run=False,
                     superscripts=True):
    """按固定顺序跑：**先上下标（会切 run）→ 再字体/颜色**，这样颜色能盖到新切出来的 run 上。

    顺序是有讲究的：上下标规则会把一个 run 切成好几段（`split_run_at`），
    后切的 run 带着原 rPr 的拷贝；先做颜色再做上下标，切出来的两半仍是黑的（也没错），
    但先把结构改完再统一刷颜色/字体，报告里的数字更好解释（"改了多少"不会被切 run 干扰）。
    """
    font_rules = font_rules or FontRuleSet(DEFAULT_FONTS)
    report = {"op": "format-normalize", "steps": [], "dry_run": bool(dry_run)}

    rule_report = None
    if superscripts and text_rules is not None:
        if dry_run:
            # **dry-run 不许动真树**（踩过：`apply_to_part` 只有"直接改"这一条路，
            # 于是 `--dry-run` 也会把上标切出来）。这里在**副本**上跑，数字照样是真的。
            import copy
            root = copy.deepcopy(document.part())
        else:
            root = document.part()
        rule_report = apply_to_part(root, text_rules, limit=10)
        report["steps"].append({"step": u"上下标规则", "total": rule_report["total"],
                                "counts": rule_report["counts"]})

    font_report = normalize_fonts(document, font_rules, dry_run=dry_run,
                                  parts=[name for name in FONT_PARTS
                                         if document.package.has(name)])
    report["steps"].append({"step": u"字体与颜色", "runs": font_report["runs"],
                            "colors": font_report["colors"],
                            "highlights": font_report["highlights"],
                            "fonts": font_report["fonts"]})

    report["fonts"] = font_report["fonts"]
    report["colors"] = font_report["colors"]
    report["highlights"] = font_report["highlights"]
    report["parts"] = font_report["parts"]
    report["superscripts"] = (rule_report or {}).get("counts", {})
    report["total"] = (sum((rule_report or {}).get("counts", {}).values())
                       + font_report["colors"] + font_report["highlights"]
                       + sum(font_report["fonts"].values()))
    if not dry_run and report["total"]:
        document.mark_dirty()
    return report
