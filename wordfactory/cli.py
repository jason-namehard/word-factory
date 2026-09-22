# -*- coding: utf-8 -*-
"""命令行入口：``python -m wordfactory.cli <命令>``

约定（沿用上一轮被验收过的口径）：
- 成功退 0，失败退 1，用法错误退 2；
- ``--json`` 时 stdout 只输出一个 JSON 对象，日志走 stderr；
- **默认不覆盖任何原文件**：要写就必须给明确的输出位置。
"""

import argparse
import io
import json
import os
import sys
import traceback

from . import __version__
from .inspect import format_report, format_text_report, inspect, text_report
from .ooxml import DocxPackage, PackageError
from .rules import RuleError, RuleSet, apply_to_part, default_ruleset, write_default

#: 规则文件默认放这里（项目根下的 rules/）。它就是用户要的"外置接口"。
DEFAULT_RULES_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                                  "rules", "subscripts.json")
DEFAULT_FONTS_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                                  "rules", "fonts.json")


def emit_json(payload):
    data = (json.dumps(payload, ensure_ascii=False, default=str) + "\n").encode("utf-8")
    try:
        sys.stdout.flush()
        sys.stdout.buffer.write(data)
        sys.stdout.buffer.flush()
    except AttributeError:
        sys.stdout.write(data.decode("utf-8"))


def emit_text(text):
    try:
        sys.stdout.write(text if text.endswith("\n") else text + "\n")
        sys.stdout.flush()
    except UnicodeEncodeError:
        sys.stdout.flush()
        encoding = getattr(sys.stdout, "encoding", None) or "utf-8"
        sys.stdout.buffer.write((text + "\n").encode(encoding, "replace"))


def log(message):
    try:
        sys.stderr.write(u"%s\n" % message)
    except Exception:
        pass


def build_parser():
    parser = argparse.ArgumentParser(
        prog="wordfactory", description=u"word 工厂 —— 外置的 Word 文档批量规范化工具箱")
    parser.add_argument("--version", action="version", version=u"word 工厂 %s" % __version__)
    parser.add_argument("--json", dest="as_json", action="store_true",
                        help=u"stdout 只输出一个 JSON 对象")
    sub = parser.add_subparsers(dest="command", metavar="<命令>")

    inspector = sub.add_parser("inspect", help=u"看一眼 .docx 里有什么（只读）")
    inspector.add_argument("path", help=u"要看的 .docx/.docm")

    texter = sub.add_parser("text",
                            help=u"逐段看「逻辑文本 ↔ run」的映射（只读）")
    texter.add_argument("path", help=u"要看的 .docx/.docm")
    texter.add_argument("--grep", default=None, help=u"只看包含这个词的段落")
    texter.add_argument("--limit", type=int, default=20, help=u"最多列几段（默认 20）")
    texter.add_argument("--part", default=None, help=u"指定部件（默认 word/document.xml）")

    rules = sub.add_parser("rules", help=u"上下标规则（外置规则文件）")
    rules_sub = rules.add_subparsers(dest="action", metavar="<动作>")
    check = rules_sub.add_parser("check", help=u"校验规则文件")
    check.add_argument("--rules", default=DEFAULT_RULES_PATH)
    init = rules_sub.add_parser("init", help=u"把默认规则写到文件（不会覆盖已有文件）")
    init.add_argument("--rules", default=DEFAULT_RULES_PATH)
    init.add_argument("--force", action="store_true")
    show = rules_sub.add_parser("show", help=u"把规则打印成人看的表")
    show.add_argument("--rules", default=DEFAULT_RULES_PATH)
    apply_cmd = rules_sub.add_parser("apply", help=u"把规则应用到文档（默认不改原文件）")
    apply_cmd.add_argument("path", help=u"要处理的 .docx")
    apply_cmd.add_argument("--rules", default=DEFAULT_RULES_PATH)
    apply_cmd.add_argument("--out", default=None, help=u"输出文件")
    apply_cmd.add_argument("--outdir", default=None, help=u"输出目录（文件名与输入相同）")
    apply_cmd.add_argument("--dry-run", action="store_true", help=u"只报告会改多少处")

    head = sub.add_parser("captions",
                          help=u"宏：题注格式统一（表题=表X-Y+空格居中+表格居中；图题=整段居中+一个空格）")
    head.add_argument("path", help=u"要处理的 .docx")
    head.add_argument("--out", default=None, help=u"输出文件")
    head.add_argument("--outdir", default=None, help=u"输出目录（文件名与输入相同）")
    head.add_argument("--dry-run", action="store_true", help=u"只列出会怎么改")
    head.add_argument("--mode", choices=("verify", "formal"), default="verify",
                      help=u"verify=改过的部分标蓝供你核对（默认）；"
                           u"formal=通体黑 + 字体合规（交付版）")
    head.add_argument("--fonts", default=DEFAULT_FONTS_PATH,
                      help=u"正式版的字体规则文件（默认 %s）" % DEFAULT_FONTS_PATH)
    head.add_argument("--no-center-table", dest="center_table", action="store_false",
                      default=True, help=u"表题的表格不要居中")
    head.add_argument("--no-center-figure", dest="center_figure", action="store_false",
                      default=True, help=u"图题不要整段居中")
    head.add_argument("--figure-space", type=int, default=1,
                      help=u"图题编号与名字之间留几个空格（默认 1）")
    head.add_argument("--no-normalize-number", dest="normalize_number", action="store_false",
                      default=True, help=u"不要把编号统一成 表X-Y / 图X-Y")
    head.add_argument("--all-captions", dest="only_before_table", action="store_false",
                      default=True, help=u"表题：连带处理后面没跟表格的那些")

    fonts = sub.add_parser("fonts", help=u"正式版的字体规则（外置文件：哪些字体算不合格）")
    fonts_sub = fonts.add_subparsers(dest="action", metavar="<动作>")
    finit = fonts_sub.add_parser("init", help=u"把默认字体规则写到文件")
    finit.add_argument("--fonts", default=DEFAULT_FONTS_PATH)
    finit.add_argument("--force", action="store_true")
    fcheck = fonts_sub.add_parser("check", help=u"校验字体规则文件")
    fcheck.add_argument("--fonts", default=DEFAULT_FONTS_PATH)
    fshow = fonts_sub.add_parser("show", help=u"当前字体规则长什么样")
    fshow.add_argument("--fonts", default=DEFAULT_FONTS_PATH)

    checker = sub.add_parser(
        "audit", help=u"体检：重新打开文件核对「通体黑 + 没有不合格字体」（末行 AUDIT=PASS/FAIL）")
    checker.add_argument("path", help=u"要体检的 .docx/.docm")
    checker.add_argument("--fonts", default=DEFAULT_FONTS_PATH, help=u"字体规则文件")

    recipe = sub.add_parser("recipe", help=u"宏「段落配方」：解析配方（show）／按配方重建段落（rebuild）")
    recipe_sub = recipe.add_subparsers(dest="action", metavar="<动作>")
    rshow = recipe_sub.add_parser("show", help=u"解析并打印配方（.docx 末尾那一段，或直接给 .txt）")
    rshow.add_argument("path", help=u"带配方的 .docx，或配方文本 .txt")
    rbuild = recipe_sub.add_parser("rebuild", help=u"读配方 + 数据表，把重建出来的段落追加到文档末尾")
    rbuild.add_argument("path", help=u"要处理的 .docx（配方也在它里面时不用给 --recipe-file）")
    rbuild.add_argument("--recipe-file", default=None, help=u"配方文本文件（不给就从文档末尾找）")
    rbuild.add_argument("--xlsx", default=None, help=u"数据表 .xlsx（不给就按配方里的 EXCEL_FILE: 去找）")
    rbuild.add_argument("--data-dir", default=None, help=u"去哪个目录找数据表（批量时用）")
    rbuild.add_argument("--out", default=None, help=u"输出文件")
    rbuild.add_argument("--dry-run", action="store_true", help=u"只演练，不写文件")
    rgen = recipe_sub.add_parser(
        "gen", help=u"生成配方：识别高亮（或占位符）→ 写 .xlsx + 把配方追加到文档末尾")
    rgen.add_argument("path", help=u"要处理的 .docx")
    rgen.add_argument("--mode", choices=("highlight", "chars"), default="highlight",
                      help=u"highlight=所有高亮片段算变量（默认）；chars=按占位符字符串算")
    rgen.add_argument("--char", default=None, help=u"模式 chars 的占位符（如 xx）")
    rgen.add_argument("--name", default=None, help=u"配方名（写进标题行）")
    rgen.add_argument("--excel-file", dest="excel_file", default=None,
                      help=u"数据表文件名（默认 数据表.xlsx，会写进配方）")
    rgen.add_argument("--sheet", default=None, help=u"工作表名（默认 Sheet1）")
    rgen.add_argument("--xlsx", default=None, help=u"数据表写到哪（默认与输入文档同目录）")
    rgen.add_argument("--out", default=None, help=u"输出文档（默认与输入同目录加 _配方 后缀）")
    rgen.add_argument("--no-append", dest="append", action="store_false", default=True,
                      help=u"只写数据表，不把配方文本追加进文档")
    rgen.add_argument("--trim-last-char", dest="trim_last_char", action="store_true",
                      default=False,
                      help=u"照抄参考宏那颗「隐形修正」（高亮片段去掉最后一个字符）—— 默认不照抄")
    rgen.add_argument("--dry-run", action="store_true", help=u"只报会识别出什么，不写任何文件")

    tclean = sub.add_parser(
        "tableclean", help=u"宏：表格空格回车删除（去无意义空格）")
    tclean.add_argument("path", help=u"要处理的 .docx")
    tclean.add_argument("--level", choices=(u"1", u"2", u"3"), default=u"3",
                        help=u"1=仅空格；2=仅回车；3=两者（默认，与宏一致）")
    tclean.add_argument("--full-width-space", dest="full_width_space", action="store_true",
                        default=False, help=u"连全角空格 U+3000 一起删（默认不删）")
    tclean.add_argument("--flat", action="store_true", default=False,
                        help=u"照抄参考宏：整格写回、压平 run 级格式（默认保留格式）")
    tclean.add_argument("--out", default=None, help=u"输出文件")
    tclean.add_argument("--outdir", default=None, help=u"输出目录（文件名与输入相同）")
    tclean.add_argument("--dry-run", action="store_true", help=u"只报会改多少，不写文件")
    return parser


def cmd_inspect(args):
    info = inspect(args.path)
    return info, format_report(info)


def cmd_text(args):
    report = text_report(args.path, limit=args.limit, grep=args.grep, part=args.part)
    return report, format_text_report(report, limit=args.limit)


def _load_rules(path):
    """规则文件不存在时回落到内置默认，并明说一声（不要静默用别的东西）。"""
    if os.path.exists(path):
        rule_set = RuleSet.load(path)
        source = path
    else:
        rule_set = default_ruleset()
        source = u"内置默认（%s 不存在）" % path
        log(u"⚠ 没找到规则文件 %s，这次用内置默认规则。" % path)
    problems = rule_set.validate()
    if problems:
        raise RuleError(u"规则文件有问题，先修好再跑：\n  - " + u"\n  - ".join(problems))
    return rule_set, source


def cmd_rules(args):
    action = args.action
    if action == "init":
        path = os.path.abspath(args.rules)
        if os.path.exists(path) and not args.force:
            raise RuleError(u"%s 已经存在；要覆盖请加 --force" % path)
        write_default(path)
        return ({"wrote": path, "rules": len(default_ruleset().rules)},
                u"已写出 %s（%d 条规则）" % (path, len(default_ruleset().rules)))

    if action == "check":
        rule_set, source = _load_rules(args.rules)
        return ({"rules": source, "count": len(rule_set.rules),
                 "enabled": len(rule_set.active), "problems": [], "ok": True},
                u"规则文件没问题：%s（共 %d 条，启用 %d 条）"
                % (source, len(rule_set.rules), len(rule_set.active)))

    if action == "show":
        rule_set, source = _load_rules(args.rules)
        lines = [u"规则：%s" % source, u"共 %d 条，启用 %d 条"
                 % (len(rule_set.rules), len(rule_set.active)), u""]
        rows = []
        for rule in rule_set.rules:
            what = (u"%s → %s" % (rule.match, rule.kinds)) if rule.match \
                else (u"/%s/ → %s %s" % (rule.pattern, rule.target, rule.kind or u""))
            rows.append({"id": rule.id, "enabled": rule.enabled, "what": what,
                         "note": rule.note})
            lines.append(u"  [%s] %-14s %-34s %s"
                         % (u"✓" if rule.enabled else u" ", rule.id, what, rule.note))
        return ({"rules": source, "items": rows}, u"\n".join(lines))

    if action == "apply":
        rule_set, source = _load_rules(args.rules)
        if not args.out and not args.outdir and not args.dry_run:
            raise RuleError(u"要写结果就得给 --out 文件或 --outdir 目录"
                            u"（本工具**不会**覆盖原文件）")
        with DocxPackage(args.path) as pkg:
            root = pkg.xml(DocxPackage.MAIN)
            report = apply_to_part(root, rule_set, limit=10)
            if args.dry_run:
                payload = {"dry_run": True, "file": pkg.path, "rules": source,
                           "total": report["total"], "counts": report["counts"],
                           "details": report["details"]}
                lines = [u"--dry-run：一个字节都没写",
                         u"文件：%s" % pkg.path,
                         u"按规则 %s 会改 %d 处：" % (source, report["total"])]
                for key in sorted(report["counts"]):
                    lines.append(u"  %-18s %d 处" % (key, report["counts"][key]))
                for detail in report["details"][:8]:
                    lines.append(u"    %s：「%s」→ %s"
                                 % (detail["rule"], detail["text"], detail["kind"]))
                return (payload, u"\n".join(lines))
            if report["total"] == 0:
                return ({"dry_run": False, "file": pkg.path, "total": 0,
                         "counts": {}, "out": None},
                        u"没有需要改的地方（%d 条规则都没命中）" % len(rule_set.active))
            pkg.mark_dirty(DocxPackage.MAIN)
            if args.out:
                out_path = pkg.save(os.path.abspath(args.out))
            else:
                out_dir = os.path.abspath(args.outdir)
                if not os.path.isdir(out_dir):
                    os.makedirs(out_dir)
                out_path = pkg.save(os.path.join(out_dir, os.path.basename(pkg.path)))
            lines = [u"已写出：%s" % out_path,
                     u"按规则 %s 共改 %d 处：" % (source, report["total"])]
            for key in sorted(report["counts"]):
                lines.append(u"  %-18s %d 处" % (key, report["counts"][key]))
            return ({"dry_run": False, "file": pkg.path, "out": out_path,
                     "total": report["total"], "counts": report["counts"]}, u"\n".join(lines))

    raise RuleError(u"未知的 rules 动作：%r" % action)


def cmd_captions(args):
    """题注格式统一：算计划 → 落地 → 按模式标蓝（验证版）或通体黑+字体合规（正式版）。"""
    from . import audit as audit_mod
    from . import fonts as fonts_mod
    from . import mark as mark_mod
    from .document import Document
    from .ops import captions as captions_op

    options = {"center_table": bool(args.center_table),
               "center_figure": bool(args.center_figure),
               "figure_space": int(args.figure_space),
               "normalize_number": bool(args.normalize_number),
               "only_before_table": bool(args.only_before_table)}
    if not args.out and not args.outdir and not args.dry_run:
        raise RuleError(u"要写结果就得给 --out 文件或 --outdir 目录"
                        u"（本工具**不会**覆盖原文件）")
    font_rules = None
    #: 正式版换字体要连样式表/编号表一起改（那是"继承来源"），所以这两条要放进白名单
    writable = fonts_mod.FONT_PARTS if args.mode == "formal" else ("word/document.xml",)
    if args.mode == "formal":
        font_rules = fonts_mod.FontRuleSet.load(args.fonts)
        problems = font_rules.check()
        if problems:
            raise RuleError(u"字体规则文件有问题：\n  - " + u"\n  - ".join(problems))

    with Document(args.path, writable_parts=writable) as doc:
        blue_before = mark_mod.count_color(doc) if args.mode == "verify" else 0
        report = captions_op.apply(doc, options, dry_run=args.dry_run)
        marked = 0
        font_report = None
        if not args.dry_run:
            if args.mode == "verify":
                marked = mark_mod.verify(doc, report["details"])
            else:
                font_report = fonts_mod.normalize(doc, font_rules)

        lines = [u"文件：%s" % doc.path,
                 u"模式：%s" % (u"验证版（改过的部分标蓝，供你核对）" if args.mode == "verify"
                               else u"正式版（通体黑 + 字体合规）"),
                 u"题注 %d 个（表 %d / 图 %d），其中 %d 个要改（表 %d / 图 %d）；表格居中 %d 张"
                 % (report["planned"], report["tables"], report["figures"],
                    report["changed"], report["changed_tables"], report["changed_figures"],
                    report["tables_centered"]),
                 u""]
        lines.append(u"%-9s %-4s %-8s %-8s %-6s %s"
                     % (u"编号", u"类型", u"原空格", u"新空格", u"要改", u"名字 / 备注"))
        for detail in report["details"]:
            lines.append(u"%-9s %-4s %-8d %-8d %-6s %s%s"
                         % (detail["number"],
                            u"图" if detail["kind"] == u"\u56fe" else u"表",
                            detail["spaces_before"], detail["spaces_after"],
                            u"是" if detail["changed"] else u"—",
                            detail["name"][:30],
                            (u"  ← " + detail["note"]) if detail["note"] else u""))
        if marked:
            lines.append(u"")
            lines.append(u"验证版：已把 %d 处「改过的那一段」标成蓝色 %s（用 Word 打开就能看出来）"
                         % (marked, mark_mod.VERIFY_BLUE))
            if blue_before:
                lines.append(u"    注意：原件里本来就有 %d 处这个蓝色（不是我标的），"
                             u"看到蓝色的标题之类属于原样保留" % blue_before)
        if font_report:
            lines.append(u"")
            lines.append(u"正式版：处理 %d 个有文字的 run ｜ 颜色改正 %d ｜ 去高亮 %d"
                         % (font_report["runs"], font_report["colors"],
                            font_report["highlights"]))
            for key in sorted(font_report["fonts"]):
                lines.append(u"    字体 %-34s %d 处" % (key, font_report["fonts"][key]))
            if not font_report["fonts"]:
                lines.append(u"    （按当前字体规则，没有需要换的字体）")

        if not args.dry_run:
            if report["changed"] or marked or (font_report and
                                              (font_report["fonts"] or font_report["colors"]
                                               or font_report["highlights"])):
                doc.mark_dirty()
                if args.out:
                    out_path = doc.save(os.path.abspath(args.out))
                else:
                    out_dir = os.path.abspath(args.outdir)
                    if not os.path.isdir(out_dir):
                        os.makedirs(out_dir)
                    out_path = doc.save(os.path.join(out_dir, os.path.basename(doc.path)))
                lines.insert(0, u"已写出：%s" % out_path)
                checked = None
                if args.mode == "formal":
                    # 报告不能自己说自己对：重新打开写出来的文件体检一遍，末行 AUDIT=…
                    checked = audit_mod.audit(out_path, font_rules)
                    lines.append(u"")
                    lines.append(audit_mod.format_audit(checked))
                payload = {"dry_run": False, "out": out_path, "mode": args.mode,
                           "changed": report["changed"], "planned": report["planned"],
                           "tables_centered": report["tables_centered"],
                           "figures": report["figures"], "marked": marked,
                           "blue_before": blue_before,
                           "parts": (font_report or {}).get("parts"),
                           "fonts": (font_report or {}).get("fonts"),
                           "colors": (font_report or {}).get("colors"),
                           "highlights": (font_report or {}).get("highlights"),
                           "audit": (checked or {}).get("verdict"),
                           "audit_reasons": (checked or {}).get("reasons"),
                           "audit_declared": (checked or {}).get("declared"),
                           "details": [dict((k, v) for k, v in d.items() if k != "paragraph")
                                       for d in report["details"]]}
                return (payload, u"\n".join(lines))
            lines.insert(0, u"没有需要改的地方（也没标蓝）")
            return ({"dry_run": False, "out": None, "mode": args.mode, "changed": 0,
                     "planned": report["planned"], "details": []}, u"\n".join(lines))

        lines.insert(0, u"--dry-run：一个字节都没写")
        payload = {"dry_run": True, "mode": args.mode, "changed": report["changed"],
                   "planned": report["planned"],
                   "tables_centered": report["tables_centered"],
                   "figures": report["figures"],
                   "details": [dict((k, v) for k, v in d.items() if k != "paragraph")
                               for d in report["details"]]}
        return (payload, u"\n".join(lines))


def cmd_fonts(args):
    from . import fonts as fonts_mod

    action = args.action
    if action == "init":
        path = os.path.abspath(args.fonts)
        if os.path.exists(path) and not args.force:
            raise RuleError(u"%s 已经存在；要覆盖请加 --force" % path)
        fonts_mod.FontRuleSet(fonts_mod.DEFAULT_FONTS).save(path)
        return ({"wrote": path}, u"已写出 %s" % path)
    rule_set = fonts_mod.FontRuleSet.load(args.fonts)
    problems = rule_set.check()
    if action == "check":
        if problems:
            raise RuleError(u"字体规则有问题：\n  - " + u"\n  - ".join(problems))
        return ({"fonts": args.fonts, "problems": [], "ok": True},
                u"字体规则没问题：%s" % (rule_set.path or u"内置默认"))
    lines = [u"字体规则：%s" % (rule_set.path or u"内置默认（%s 不存在）" % args.fonts),
             u"  keep（原样保留）      ：%s" % u"、".join(sorted(rule_set.keep)),
             u"  replace（明确换掉）   ：%s" % (u"、".join(
                 u"%s→%s" % (k, v) for k, v in sorted(rule_set.replace.items())) or u"（无）"),
             u"  default（其余中文字体）：%s" % (rule_set.default or u"（无 → 不动）"),
             u"  default_latin（西文）：%s" % (rule_set.default_latin or u"（无 → 西文不动）"),
             u"  default_scope        ：%s（%s）"
             % (rule_set.default_scope,
                u"只管中文字体（西文一概不碰）" if rule_set.default_scope != u"all"
                else u"中文字体与西文字体都统一"),
             u"  symbol_fonts（永不碰） ：%s" % u"、".join(sorted(rule_set.symbol_fonts)),
             u"  black_all            ：%s" % rule_set.black_all,
             u"  remove_highlight     ：%s" % rule_set.remove_highlight]
    return ({"fonts": args.fonts, "keep": sorted(rule_set.keep),
             "replace": rule_set.replace, "default": rule_set.default,
             "default_latin": rule_set.default_latin,
             "default_scope": rule_set.default_scope,
             "symbol_fonts": sorted(rule_set.symbol_fonts),
             "black_all": rule_set.black_all, "remove_highlight": rule_set.remove_highlight},
            u"\n".join(lines))


def cmd_recipe(args):
    """宏「段落配方」：`show` 解析配方，`rebuild` 按配方 + 数据表重建段落。"""
    from .ops import recipe as recipe_op
    from .recipe import RecipeError

    if args.action == "show":
        recipe = recipe_op.read_recipe(args.path)
        counts = {}
        for kind, _ in recipe.lines:
            counts[kind] = counts.get(kind, 0) + 1
        lines = [u"配方：%s" % recipe.name,
                 u"  数据表：%s" % recipe.excel_file,
                 u"  工作表：%s" % recipe.sheet_name,
                 u"  变量数：%d" % recipe.variable_count,
                 u"  正文行：%d 行（TEXT %d ／ RAW %d ／ VAR %d）"
                 % (len(recipe.lines), counts.get("TEXT", 0), counts.get("RAW", 0),
                    counts.get("VAR", 0)),
                 u""]
        for warning in recipe.warnings:
            lines.append(u"⚠ %s" % warning)
        number = 0
        for kind, content in recipe.lines:
            if kind == "VAR":
                number += 1
                lines.append(u"    %-4s → 取第 %d 个值" % (kind, number))
            elif kind == "RAW":
                lines.append(u"    %-4s %s" % (kind, content))
            else:
                lines.append(u"    %-4s %s" % (kind, content))
        return ({"recipe": recipe.name, "excel": recipe.excel_file,
                 "sheet": recipe.sheet_name, "variables": recipe.variable_count,
                 "lines": [{"kind": k, "text": t} for k, t in recipe.lines]},
                u"\n".join(lines))

    if args.action == "rebuild":
        from .document import Document
        if not args.out and not args.dry_run:
            raise RuleError(u"要写结果就得给 --out（本工具**不会**覆盖原文件）")
        with Document(args.path) as doc:
            recipe = (recipe_op.read_recipe(args.recipe_file) if args.recipe_file
                      else recipe_op.read_recipe(args.path))
            xlsx = args.xlsx or recipe_op.resolve_xlsx(recipe, args.path, args.data_dir)
            values = recipe_op.values_for(recipe, xlsx)
            report = recipe_op.rebuild(doc, recipe, values, dry_run=args.dry_run)
            out_path = None
            if not args.dry_run:
                out_path = doc.save(os.path.abspath(args.out))
            lines = [u"配方：%s（%d 个变量）" % (recipe.name, recipe.variable_count),
                     u"数据表：%s ｜ 工作表 %s ｜ 取到 %d 个值"
                     % (xlsx, recipe.sheet_name, len(values)),
                     u"重建：%d 个段落%s" % (report["paragraphs"],
                                            u"（其中 %d 处 #数据缺失#）" % report["missing"]
                                            if report["missing"] else u""),
                     u""]
            for warning in recipe.warnings:
                lines.insert(2, u"⚠ %s" % warning)
            if args.dry_run:
                lines.insert(0, u"--dry-run：一个字节都没写")
            else:
                lines.insert(0, u"已写出：%s" % out_path)
            lines.append(u"重建出来的段落正文：")
            for text in report["text"].split(u"\n"):
                lines.append(u"    %s" % text)
            return (dict(report, out=out_path, xlsx=xlsx), u"\n".join(lines))

    if args.action == "gen":
        from .document import Document
        title = args.name or u"默认配方"
        excel_file = args.excel_file or u"数据表.xlsx"
        out_doc = args.out or os.path.join(
            os.path.dirname(os.path.abspath(args.path)),
            os.path.splitext(os.path.basename(args.path))[0] + u"_配方.docx")
        out_xlsx = args.xlsx or os.path.join(
            os.path.dirname(os.path.abspath(out_doc)), excel_file)
        options = {"mode": args.mode, "char": args.char, "name": title,
                   "excel_file": excel_file, "sheet_name": args.sheet or u"Sheet1",
                   "out_xlsx": os.path.abspath(out_xlsx), "append": args.append,
                   "trim_last_char": args.trim_last_char}
        if not args.dry_run and os.path.exists(out_xlsx):
            raise RuleError(u"%s 已经存在；请用 --xlsx 换个位置（不覆盖已有数据表）" % out_xlsx)
        with Document(args.path) as doc:
            report, recipe, xlsx_path = recipe_op.generate(doc, options, dry_run=args.dry_run)
            written = None
            if not args.dry_run:
                written = doc.save(os.path.abspath(out_doc))
        lines = [u"识别模式：%s" % (u"高亮（所有高亮片段）" if args.mode == "highlight"
                                   else u"特定字符 %r" % args.char),
                 u"配方：%s ｜ 变量 %d 个" % (recipe.name, report["variables"]),
                 u"数据表：%s ｜ 工作表 %s" % (excel_file, recipe.sheet_name)]
        for index, (prefix, value) in enumerate(report["rows"], start=1):
            lines.append(u"    %2d. A=%s ｜ B=%s" % (index, prefix[:20] or u"（空）", value[:30]))
        if args.dry_run:
            lines.insert(0, u"--dry-run：一个字节都没写")
        else:
            lines.insert(0, u"已写出：%s" % written)
            lines.append(u"已写出数据表：%s" % xlsx_path)
            lines.append(u"配方已追加到文档末尾（%d 段）" % report["paragraphs"]
                         if args.append else u"（按 --no-append，没往文档里写配方）")
        return (dict(report, out=written, xlsx=xlsx_path), u"\n".join(lines))

    raise RecipeError(u"未知的 recipe 动作：%r" % args.action)


def cmd_tableclean(args):
    """宏「表格空格回车删除」（去无意义空格）：清表格单元格里的空格/换行。"""
    from .document import Document
    from .ops import table_clean as table_op

    if not args.out and not args.outdir and not args.dry_run:
        raise RuleError(u"要写结果就得给 --out 文件或 --outdir 目录"
                        u"（本工具**不会**覆盖原文件）")
    options = {"level": int(args.level), "full_width_space": bool(args.full_width_space),
               "flat": bool(args.flat)}
    with Document(args.path) as doc:
        report = table_op.clean(doc, options, dry_run=args.dry_run)
        lines = [u"文件：%s" % doc.path,
                 u"档位 %d：%s%s" % (report["level"], report["level_note"],
                                    u"（含全角空格）" if report["full_width_space"] else u""),
                 u"表格 %d 张 ｜ 单元格 %d 个 ｜ 其中 %d 个被清理"
                 % (report["tables"], report["cells"], report["changed_cells"]),
                 u"删掉 %d 个字符%s" % (report["chars_removed"],
                                       u"，合并 %d 个段落" % report["paragraphs_merged"]
                                       if report["paragraphs_merged"] else u""),
                 u""]
        if args.dry_run:
            lines.insert(0, u"--dry-run：一个字节都没写")
        else:
            if not report["changed_cells"]:
                lines.insert(0, u"没有需要清理的单元格")
                return (dict(report, out=None), u"\n".join(lines))
            if args.out:
                out_path = doc.save(os.path.abspath(args.out))
            else:
                out_dir = os.path.abspath(args.outdir)
                if not os.path.isdir(out_dir):
                    os.makedirs(out_dir)
                out_path = doc.save(os.path.join(out_dir, os.path.basename(doc.path)))
            lines.insert(0, u"已写出：%s" % out_path)
            report["out"] = out_path
        lines.append(u"与参考宏的两处有意不同：")
        lines.append(u"  ① 只删字符、**保留** run 级格式（宏是整格写回、格式会被压平）；"
                     u"要照抄宏加 --flat")
        lines.append(u"  ② 默认**不删全角空格**（宏也删不掉；中文报告里它常是段首缩进）；"
                     u"要删加 --full-width-space")
        return (report, u"\n".join(lines))


def cmd_audit(args):
    """体检：不信工具的报告，重新打开文件按继承链算一遍。"""
    from . import audit as audit_mod
    from . import fonts as fonts_mod

    rule_set = fonts_mod.FontRuleSet.load(args.fonts)
    problems = rule_set.check()
    if problems:
        raise RuleError(u"字体规则有问题：\n  - " + u"\n  - ".join(problems))
    report = audit_mod.audit(args.path, rule_set)
    return report, audit_mod.format_audit(report)


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv if argv is not None else sys.argv[1:])
    if not args.command:
        parser.print_help()
        return 2
    try:
        if args.command == "inspect":
            payload, human = cmd_inspect(args)
        elif args.command == "text":
            payload, human = cmd_text(args)
        elif args.command == "rules":
            payload, human = cmd_rules(args)
        elif args.command == "fonts":
            payload, human = cmd_fonts(args)
        elif args.command == "captions":
            payload, human = cmd_captions(args)
        elif args.command == "audit":
            payload, human = cmd_audit(args)
        elif args.command == "recipe":
            payload, human = cmd_recipe(args)
        elif args.command == "tableclean":
            payload, human = cmd_tableclean(args)
        else:
            log(u"未知命令：%s" % args.command)
            return 2
    except PackageError as exc:
        if args.as_json:
            emit_json({"ok": False, "error": u"%s" % exc})
        else:
            log(u"错误：%s" % exc)
        return 1
    except RuleError as exc:
        if args.as_json:
            emit_json({"ok": False, "error": u"%s" % exc})
        else:
            log(u"规则有问题：%s" % exc)
        return 1
    except Exception as exc:                       # noqa: BLE001 - 兜底要如实报错
        log(traceback.format_exc())
        if args.as_json:
            emit_json({"ok": False, "error": u"%s: %s" % (type(exc).__name__, exc)})
        else:
            log(u"内部错误：%s: %s" % (type(exc).__name__, exc))
        return 1
    if args.as_json:
        emit_json(dict(payload, ok=True))
    else:
        emit_text(human)
    if args.command == "audit" and payload.get("verdict") != "PASS":
        return 1                       # 体检没过就退 1，方便当闸门串进流程
    return 0


if __name__ == "__main__":
    sys.exit(main())
