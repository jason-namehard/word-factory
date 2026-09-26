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
DEFAULT_STYLES_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                                    "rules", "tablestyle.json")
DEFAULT_REPLACEMENTS_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "rules", "replacements.json")


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
    tclean.add_argument("--exact-macro", dest="exact_macro", action="store_true", default=False,
                        help=u"换行口径也照抄宏：所有换行一律删（默认只删无意义的，见下）")
    tclean.add_argument("--out", default=None, help=u"输出文件")
    tclean.add_argument("--outdir", default=None, help=u"输出目录（文件名与输入相同）")
    tclean.add_argument("--dry-run", action="store_true", help=u"只报会改多少，不写文件")
    _add_tablestyle_parser(sub)
    _add_textfix_parser(sub)
    _add_tidy_parser(sub)
    _add_mdclean_parser(sub)
    _add_format_parser(sub)
    _add_gui_parser(sub)
    _add_pdf_parser(sub)
    return parser


def _add_format_parser(sub):
    """格式规范化（宏里那套"变黑 + 去高亮 + 字体 + 上标"跑成一条命令）。"""
    fmt = sub.add_parser("format", help=u"宏：格式规范化（通体黑 + 去高亮 + 字体合规 + 上下标规则）")
    fmt.add_argument("path", help=u"要处理的 .docx")
    fmt.add_argument("--fonts", default=DEFAULT_FONTS_PATH, help=u"字体规则文件")
    fmt.add_argument("--rules", default=DEFAULT_RULES_PATH, help=u"上下标规则文件")
    fmt.add_argument("--no-superscripts", dest="no_superscripts", action="store_true",
                     default=False, help=u"不跑上下标规则（只做颜色/字体）")
    fmt.add_argument("--out", default=None, help=u"输出文件")
    fmt.add_argument("--outdir", default=None, help=u"输出目录（文件名与输入相同）")
    fmt.add_argument("--dry-run", action="store_true", help=u"只报会改多少，不写文件")


def _add_pdf_parser(sub):
    """PDF 导出（编排外部渲染器）。"""
    pdf = sub.add_parser("pdf", help=u"导出 PDF（编排 Word / WPS / LibreOffice）")
    pdf.add_argument("path", nargs="?", help=u"要导出的 .docx")
    pdf.add_argument("--out", default=None, help=u"输出 .pdf（默认与文档同名的 .pdf）")
    pdf.add_argument("--renderer", default=None, choices=("word", "wps", "libreoffice"),
                     help=u"指定渲染器（默认用第一个可用的）")
    pdf.add_argument("--timeout", type=int, default=300, help=u"超时秒数（默认 300）")
    pdf.add_argument("--hidden", action="store_true", default=False,
                     help=u"不显示渲染器窗口（Word/WPS 可见更容易发现卡住）")
    pdf.add_argument("--dry-run", action="store_true", help=u"只报打算怎么导，不真跑")
    pdf.add_argument("--list", action="store_true", help=u"看本机有哪些渲染器")


def _add_gui_parser(sub):
    """本地网页 GUI + 配方一键跑。"""
    gui = sub.add_parser("gui", help=u"起本地网页 GUI（默认只监听 127.0.0.1）")
    gui.add_argument("--host", default="127.0.0.1", help=u"监听地址（默认只监听本机）")
    gui.add_argument("--port", type=int, default=8765, help=u"端口（默认 8765）")
    gui.add_argument("--root", default=None, help=u"文件浏览的根目录（默认用户主目录）")
    gui.add_argument("--no-browser", dest="no_browser", action="store_true",
                     default=False, help=u"不自动打开浏览器")

    runner = sub.add_parser("run", help=u"按配方一键跑（与 GUI 同一条路）")
    runner.add_argument("path", help=u"要处理的 .docx")
    runner.add_argument("--steps", default=None,
                        help=u"配方步骤，逗号分隔（如 captions,sup,tidy）")
    runner.add_argument("--recipe", default=None, help=u"配方 JSON 文件（含 steps 数组）")
    runner.add_argument("--mode", choices=("verify", "formal"), default="verify",
                        help=u"verify=改过的地方标蓝；formal=通体黑 + 字体合规")
    runner.add_argument("--out", default=None, help=u"输出文件（默认「某报告（验证版）.docx」）")
    runner.add_argument("--dry-run", action="store_true", help=u"只报会改多少，不写文件")


def _add_mdclean_parser(sub):
    """Markdown 标记清理（宏 MarkDown语言清除）。"""
    md = sub.add_parser("mdclean", help=u"宏：把 Markdown 标记清成普通中文文本")
    md.add_argument("path", help=u"要处理的 .docx")
    md.add_argument("--scope", choices=("body", "all"), default="body",
                    help=u"body=只动正文段落（默认）；all=连表格里的也动")
    md.add_argument("--out", default=None, help=u"输出文件")
    md.add_argument("--outdir", default=None, help=u"输出目录（文件名与输入相同）")
    md.add_argument("--dry-run", action="store_true", help=u"只报会改多少，不写文件")


def _add_tidy_parser(sub):
    """一键清理（去空格 / 去无意义空白行）。"""
    tidy_p = sub.add_parser("tidy", help=u"一键清理：无意义空白行 + 段尾空格")
    tidy_p.add_argument("path", help=u"要处理的 .docx")
    tidy_p.add_argument("--no-blank-lines", dest="no_blank_lines", action="store_true",
                        default=False, help=u"不动空白行")
    tidy_p.add_argument("--no-spaces", dest="no_spaces", action="store_true", default=False,
                        help=u"不动段尾空格")
    tidy_p.add_argument("--trim-leading", dest="trim_leading", action="store_true", default=False,
                        help=u"连段首空格也删（危险：可能是用空格做的缩进）")
    tidy_p.add_argument("--collapse-space-runs", dest="collapse_space_runs",
                        action="store_true", default=False,
                        help=u"把段内连续空格压成一个（自动跳过题注段落）")
    tidy_p.add_argument("--include-captions", dest="include_captions", action="store_true",
                        default=False,
                        help=u"去空格/压空格时**连题注段落一起动**（默认跳过 —— 它们的空格是居中用的）")
    tidy_p.add_argument("--merge-lines", dest="merge_lines", action="store_true",
                        default=False,
                        help=u"合并换行：**删掉全部空行**（照 Copy++ 的「合并换行」，见样本）")
    tidy_p.add_argument("--remove-spaces", dest="remove_spaces", action="store_true",
                        default=False,
                        help=u"去除空格：删掉全部半角/不间断空格（照 Copy++ 的「去除空格」）")
    tidy_p.add_argument("--scope", choices=("body", "all"), default="body",
                        help=u"body=只动正文段落（默认）；all=连表格里的也动")
    tidy_p.add_argument("--out", default=None, help=u"输出文件")
    tidy_p.add_argument("--outdir", default=None, help=u"输出目录（文件名与输入相同）")
    tidy_p.add_argument("--dry-run", action="store_true", help=u"只报会改多少，不写文件")


def _add_textfix_parser(sub):
    """文本替换 + 对齐（宏「规划报告一键宏」的文本部分）。"""
    tfix = sub.add_parser("textfix", help=u"宏：批量文本替换 + 两端对齐改左对齐")
    tfix.add_argument("path", help=u"要处理的 .docx")
    tfix.add_argument("--rules", default=DEFAULT_REPLACEMENTS_PATH, help=u"替换规则文件")
    tfix.add_argument("--no-align", dest="no_align", action="store_true", default=False,
                      help=u"只做替换，不动对齐")
    tfix.add_argument("--fix-styles", dest="fix_styles", action="store_true", default=False,
                      help=u"连样式表里的两端对齐也一起改（默认只改段落）")
    tfix.add_argument("--out", default=None, help=u"输出文件")
    tfix.add_argument("--outdir", default=None, help=u"输出目录（文件名与输入相同）")
    tfix.add_argument("--dry-run", action="store_true", help=u"只报会改多少，不写文件")


def _add_tablestyle_parser(sub):
    """表格款式：四个动作一个入口（款式文件是外置的，所以 show/init 也要有）。"""
    tstyle = sub.add_parser("tablestyle",
                            help=u"表格款式（外置规则）：套用 / 采集你调好的表 / 预览")
    tstyle.add_argument("action",
                        choices=("init", "show", "list", "capture", "preview", "apply"),
                        help=u"init 写默认款式；show 看款式；list 遍历表头（挑表用）；"
                             u"capture 采集你调好的表；preview 预览；apply 套用")
    tstyle.add_argument("path", nargs="?", help=u"capture / apply 时要处理的 .docx")
    tstyle.add_argument("--styles", default=DEFAULT_STYLES_PATH, help=u"款式文件")
    tstyle.add_argument("--style", default=None, help=u"款式名（preview 可给多个，逗号分隔）")
    tstyle.add_argument("--table", default=None, help=u"capture 时采集第几张表（默认第 1 张）")
    tstyle.add_argument("--name", default=None, help=u"capture 时存成什么名字")
    tstyle.add_argument("--tables", default="all",
                        help=u"apply 作用范围：all ／ 序号 1,4-6 ／ **表头关键词**（如 \"序号,项目\"）")
    tstyle.add_argument("--wrap-header", dest="wrap_header", action="append", default=None,
                        help=u"把某个表头折成两行展示（可重复；给表头文字或列号，如 --wrap-header 备注）")
    tstyle.add_argument("--out", default=None, help=u"输出文件（capture/preview/apply 用）")
    tstyle.add_argument("--outdir", default=None, help=u"输出目录（apply 用）")
    tstyle.add_argument("--force", action="store_true", help=u"init 时覆盖已有款式文件")
    tstyle.add_argument("--dry-run", action="store_true", help=u"apply 时只报会改多少")


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
            if args.dry_run:
                # dry-run 在**副本**上跑：`apply_to_part` 只有"直接改树"这一条路，在真树上跑虽然
                # 不落盘，但"dry-run 不改树"这条不变量就名存实亡了（`format` 那边踩过同一个坑）。
                import copy
                root = copy.deepcopy(pkg.xml(DocxPackage.MAIN))
            else:
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
               "flat": bool(args.flat), "exact_macro": bool(args.exact_macro),
               "redundant_only": not args.exact_macro,
               "numeric_flatten": not args.exact_macro}
    with Document(args.path) as doc:
        report = table_op.clean(doc, options, dry_run=args.dry_run)
        lines = [u"文件：%s" % doc.path,
                 u"档位 %d：%s%s" % (report["level"], report["level_note"],
                                    u"（含全角空格）" if report["full_width_space"] else u""),
                 u"表格 %d 张 ｜ 单元格 %d 个 ｜ 其中 %d 个被清理"
                 % (report["tables"], report["cells"], report["changed_cells"]),
                 u"删掉 %d 个字符%s" % (report["chars_removed"],
                                       u"，合并 %d 个段落" % report["paragraphs_merged"]
                                       if report["paragraphs_merged"] else u"")]
        if report.get("redundant_only"):
            lines.append(u"换行口径：只删「首尾的 / 连续重复的」换行，**保留单个内部换行**"
                         u"（那是为了让窄列好看故意折的行）");
            if report.get("breaks_kept"):
                lines.append(u"          保留了 %d 个内部换行" % report["breaks_kept"])
            if report.get("numeric_flatten"):
                lines.append(u"          整格是纯数字的单元格例外：换行一律删（数字里断行是脏数据）")
        else:
            lines.append(u"换行口径：照抄参考宏 —— **所有**换行一律删（--exact-macro）")
        lines.append(u"")
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


def cmd_tablestyle(args):
    """表格款式：外置款式文件 + 套用 + 采集（把你调好的表存成款式）+ 预览 + 列表。"""
    from . import tablestyle as ts
    from .document import Document

    if args.action == "init":
        # **先写再校验**：init 是"生成一份新款式文件"，不该因为**旧文件**有问题就跑不动
        # （踩过：旧款里"内框和外框一样粗"把 init 拦住了 → 永远换不成新默认款）。
        path = os.path.abspath(args.styles)
        if os.path.exists(path) and not args.force:
            raise RuleError(u"%s 已经存在；要覆盖请加 --force" % path)
        ts.StyleSet(ts.DEFAULT_STYLES).save(path)
        return ({"wrote": path, "styles": len(ts.DEFAULT_STYLES["styles"])},
                u"已写出 %s（%d 个内置款式：%s）"
                % (path, len(ts.DEFAULT_STYLES["styles"]),
                   u"、".join(ts.DEFAULT_STYLES["styles"])))

    # init 已提前返回；剩下四个动作都要用**现有**款式文件 —— 到这里才校验
    if args.action == "list":
        # **遍历表头**：把每张表的序号 + 表头文字列出来 —— 用户要的"按表头叫这些表格"，
        # 也是将来 GUI 那排复选框的数据源（`--json` 下是结构化的 summaries）。
        with Document(args.path) as doc:
            summaries = ts.table_summaries(doc)
        lines = [u"文件：%s" % doc.path, u"共 %d 张表：" % len(summaries), u""]
        lines.append(u"  %-4s %-7s %-7s %-8s %s" % (u"序号", u"行数", u"列数", u"列宽", u"表头"))
        for item in summaries:
            lines.append(u"  %-4d %-7d %-7s %-8s %s"
                         % (item["index"], item["rows"],
                            item["columns"] if item["columns"] else u"?",
                            u"等宽" if item["equal_widths"] else u"不匀",
                            item["header_text"][:60] or u"（表头是空的）"))
        lines.append(u"")
        lines.append(u"选表方式（`apply --tables`）：`all` ／ 序号 `1,4-6` ／ "
                     u"**表头关键词**（如 `--tables \"序号,项目\"` —— 表头里出现这些词的都会被选中）")
        return ({"file": doc.path, "tables": summaries}, u"\n".join(lines))

    style_set = ts.StyleSet.load(args.styles)
    problems = style_set.check()
    if problems:
        raise RuleError(u"款式文件有问题：\n  - " + u"\n  - ".join(problems))

    if args.action == "show":
        lines = [u"款式文件：%s" % (style_set.path or u"内置默认（%s 不存在）" % args.styles),
                 u"共 %d 个款式：" % len(style_set.styles)]
        for style in style_set.styles.values():
            lines.append(u"  【%s】%s" % (style.name, style.note or u""))
            lines.append(u"      框线：%s"
                         % (u"、".join(u"%s=%s" % (edge, spec.get("val"))
                                      for edge, spec in style.borders.items()) or u"（不动）"))
            lines.append(u"      内边距：%s ｜ 垂直居中：%s ｜ 列宽：%s ｜ 表格对齐：%s"
                         % (u",".join(u"%s:%s" % kv for kv in sorted(style.cell_margins.items()))
                            or u"（不动）", style.v_align or u"（不动）",
                            style.column_widths or u"（不动）",
                            style.table_align or u"（不动）"))
            lines.append(u"      表头：%s"
                         % (u",".join(u"%s=%s" % kv for kv in style.header.items()) or u"（不动）"))
        return ({"styles": list(style_set.styles), "path": style_set.path}, u"\n".join(lines))

    if args.action == "capture":
        if not args.out:
            raise RuleError(u"要保存款式就得给 --out（写款式文件；不会动输入文档）")
        with Document(args.path) as doc:
            data = ts.capture(doc, int(args.table or 1))
        name = args.name or u"款式%s" % (args.table or 1)
        target = ts.StyleSet.load(args.out) if os.path.exists(args.out) else ts.StyleSet()
        target.put(name, data)
        target.save(os.path.abspath(args.out))
        lines = [u"已把 %s 的第 %s 张表**现在的样子**存成款式【%s】"
                 % (doc.path, args.table or 1, name),
                 u"写出：%s（共 %d 个款式）" % (os.path.abspath(args.out), len(target.styles)),
                 u""]
        for key, value in data.items():
            lines.append(u"  %-22s %s" % (key, value if not isinstance(value, dict)
                                          else u", ".join(u"%s=%s" % kv for kv in value.items())))
        lines.append(u"")
        lines.append(u"以后 `tablestyle apply --style %s` 就能把这张表的样子复用到别的文档。" % name)
        return ({"captured": data, "style": name, "out": os.path.abspath(args.out)},
                u"\n".join(lines))

    if args.action == "preview":
        names = [name.strip() for name in (args.style or u"").split(u",") if name.strip()]
        styles = [style_set.get(name) for name in names] if names else list(style_set.styles.values())
        if not args.out:
            raise RuleError(u"预览要写文件：请给 --out（例如 preview.docx），然后用 Word 打开挑")
        path = ts.build_preview(os.path.abspath(args.out), styles)
        lines = [u"预览文档：%s" % path,
                 u"里面有 %d 种款式 × %d 种代表性表格（每张表都标了款式名）："
                 % (len(styles), len(ts.PREVIEW_SHEETS))]
        for style in styles:
            lines.append(u"  【%s】%s" % (style.name, style.note or u""))
        lines.append(u"")
        lines.append(u"用 Word 打开看一眼，选中哪个就把名字给 `tablestyle apply --style 名字`。")
        lines.append(u"注意：预览只是「样子参考」—— 真正的高度/列宽由 Word 排版决定，我们只按内容估。")
        return ({"preview": path, "styles": [style.name for style in styles]}, u"\n".join(lines))

    if args.action == "apply":
        if not args.out and not args.outdir and not args.dry_run:
            raise RuleError(u"要写结果就得给 --out 文件或 --outdir 目录"
                            u"（本工具**不会**覆盖原文件）")
        style = style_set.get(args.style)
        if args.wrap_header:
            # 命令行指定的表头折行**叠加**在款式之上（款式文件里也能写 header_wrap）
            for name in args.wrap_header:
                for part in name.split(u","):
                    part = part.strip()
                    if part and part not in style.header_wrap:
                        style.header_wrap.append(part)
        with Document(args.path) as doc:
            report = ts.apply(doc, style, selector=args.tables, dry_run=args.dry_run)
            out_path = None
            if not args.dry_run and report["total"]:
                if args.out:
                    out_path = doc.save(os.path.abspath(args.out))
                else:
                    out_dir = os.path.abspath(args.outdir)
                    if not os.path.isdir(out_dir):
                        os.makedirs(out_dir)
                    out_path = doc.save(os.path.join(out_dir, os.path.basename(doc.path)))
            lines = [u"文件：%s" % doc.path,
                     u"款式【%s】：%s" % (style.name, style.note or u""),
                     u"表格 %d 张，选中 %d 张，改了 %d 张（共 %d 处）"
                     % (report["tables"], report["selected"], report["changed_tables"],
                        report["total"])]
            for key in sorted(report["changes"]):
                lines.append(u"    %-14s %d 处" % (key, report["changes"][key]))
            if args.dry_run:
                lines.insert(0, u"--dry-run：一个字节都没写")
            elif out_path:
                lines.insert(0, u"已写出：%s" % out_path)
            else:
                lines.insert(0, u"没有需要改的地方（这些表已经是这个样子）")
            return (dict(report, out=out_path), u"\n".join(lines))

    raise RuleError(u"未知的 tablestyle 动作：%r" % args.action)


def cmd_textfix(args):
    """宏「规划报告一键宏」的文本部分：批量替换 + 两端对齐改左对齐。"""
    from .document import Document
    from .ops import textfix as textfix_op

    if not args.out and not args.outdir and not args.dry_run:
        raise RuleError(u"要写结果就得给 --out 文件或 --outdir 目录"
                        u"（本工具**不会**覆盖原文件）")
    rule_set = textfix_op.ReplacementRuleSet.load(args.rules)
    problems = rule_set.check()
    if problems:
        raise RuleError(u"替换规则有问题：\n  - " + u"\n  - ".join(problems))
    with Document(args.path) as doc:
        report = textfix_op.apply(doc, rule_set, dry_run=args.dry_run,
                                  fix_align=not args.no_align, fix_styles=args.fix_styles)
        out_path = None
        if not args.dry_run and report["total"]:
            if args.out:
                out_path = doc.save(os.path.abspath(args.out))
            else:
                out_dir = os.path.abspath(args.outdir)
                if not os.path.isdir(out_dir):
                    os.makedirs(out_dir)
                out_path = doc.save(os.path.join(out_dir, os.path.basename(doc.path)))
    lines = [u"文件：%s" % doc.path,
             u"规则：%s（%s）" % (rule_set.name, rule_set.path or u"内置默认"),
             u"全文 %d 段 ｜ 替换 %d 处 ｜ 对齐改正 %d 段 ｜ 合计 %d 处"
             % (report["paragraphs"], report["replaced"], report["align_fixed"],
                report["total"])]
    for key in sorted(report["changes"]):
        lines.append(u"    %-24s %d 处" % (key, report["changes"][key]))
    if report["align_by_source"]:
        lines.append(u"    对齐改正的来源：%s"
                     % u"、".join(u"%s %d 段" % kv for kv in
                                  sorted(report["align_by_source"].items())))
        lines.append(u"    （**含「通过样式继承」的那批**：只改段落自己的属性会漏掉它们）")
    for item in report["replacements"][:8]:
        lines.append(u"    %r → %r 命中 %d 次（例：%s）"
                     % (item["from"], item["to"], item["count"], item["paragraph"]))
    if args.dry_run:
        lines.insert(0, u"--dry-run：一个字节都没写")
    elif out_path:
        lines.insert(0, u"已写出：%s" % out_path)
    else:
        lines.insert(0, u"没有需要改的地方")
    return (dict(report, out=out_path), u"\n".join(lines))


def cmd_tidy(args):
    """一键清理：无意义空白行 + 段首/段尾空格（从网上粘来的文字最常见的两种脏）。"""
    from .document import Document
    from .ops import tidy as tidy_op

    if not args.out and not args.outdir and not args.dry_run:
        raise RuleError(u"要写结果就得给 --out 文件或 --outdir 目录"
                        u"（本工具**不会**覆盖原文件）")
    options = {"blank_lines": not args.no_blank_lines,
               "trailing_spaces": not args.no_spaces,
               "trim_leading": args.trim_leading,
               "collapse_space_runs": args.collapse_space_runs,
               "merge_lines": args.merge_lines,
               "remove_spaces": args.remove_spaces,
               "caption_skip": not args.include_captions,
               "scope": args.scope}
    with Document(args.path) as doc:
        report = tidy_op.tidy(doc, options, dry_run=args.dry_run)
        out_path = None
        if not args.dry_run and report["total"]:
            if args.out:
                out_path = doc.save(os.path.abspath(args.out))
            else:
                out_dir = os.path.abspath(args.outdir)
                if not os.path.isdir(out_dir):
                    os.makedirs(out_dir)
                out_path = doc.save(os.path.join(out_dir, os.path.basename(doc.path)))
    lines = [u"文件：%s" % doc.path,
             u"范围：%s" % (u"正文段落" if options["scope"] == "body"
                            else u"全文（含表格）"),
             u"合计 %d 处" % report["total"]]
    for key in sorted(report["changes"]):
        lines.append(u"    %-22s %d" % (key, report["changes"][key]))
    if not options["trim_leading"]:
        lines.append(u"    （段首空格默认**没动** —— 怕碰到「用空格当缩进」的文档；要动加 --trim-leading）")
    if options["merge_lines"]:
        lines.append(u"    （合并换行：空白行**全删**，不是压成一个 —— 与你给的 Copy++ 样本一致）")
    if options["remove_spaces"]:
        lines.append(u"    （去除空格：半角与不间断空格全删；**全角空格保留**，它常是段首缩进）")
    if not options["collapse_space_runs"]:
        lines.append(u"    （段内连续空格默认**没动** —— 它会把表题的空格居中压掉；要动加 --collapse-space-runs，"
                     u"会自动跳过题注段落）")
    if args.dry_run:
        lines.insert(0, u"--dry-run：一个字节都没写")
    elif out_path:
        lines.insert(0, u"已写出：%s" % out_path)
    else:
        lines.insert(0, u"没有需要清理的地方")
    return (dict(report, out=out_path), u"\n".join(lines))


def cmd_mdclean(args):
    """宏「MarkDown语言清除」：把 Markdown 标记清成普通中文文本。"""
    from .document import Document
    from .ops import mdclean as md_op

    if not args.out and not args.outdir and not args.dry_run:
        raise RuleError(u"要写结果就得给 --out 文件或 --outdir 目录"
                        u"（本工具**不会**覆盖原文件）")
    with Document(args.path) as doc:
        report = md_op.apply(doc, {"scope": args.scope}, dry_run=args.dry_run)
        out_path = None
        if not args.dry_run and report["touched"]:
            if args.out:
                out_path = doc.save(os.path.abspath(args.out))
            else:
                out_dir = os.path.abspath(args.outdir)
                if not os.path.isdir(out_dir):
                    os.makedirs(out_dir)
                out_path = doc.save(os.path.join(out_dir, os.path.basename(doc.path)))
    lines = [u"文件：%s" % doc.path,
             u"范围：%s" % (u"正文段落" if args.scope == "body" else u"全文"),
             u"扫了 %d 段，其中 %d 段带 Markdown 标记并已清理（共 %d 处）"
             % (report["paragraphs"], report["touched"], report["total"])]
    for key in sorted(report["changes"]):
        lines.append(u"    %-12s %d 处" % (key, report["changes"][key]))
    for item in report["samples"]:
        lines.append(u"    例：%s" % item["before"])
        lines.append(u"      → %s" % item["after"])
    if args.dry_run:
        lines.insert(0, u"--dry-run：一个字节都没写")
    elif out_path:
        lines.insert(0, u"已写出：%s" % out_path)
    else:
        lines.insert(0, u"没有需要清理的 Markdown 标记")
    lines.append(u"")
    lines.append(u"说明：只改文字、**不重建 run** —— 匹配到的那一小段被替换，其余文字的格式原样保留。")
    return (dict(report, out=out_path), u"\n".join(lines))


def cmd_format(args):
    """宏「格式规范化」：通体黑 + 去高亮 + 字体合规 + 上下标规则（固定顺序一条命令跑完）。"""
    from .document import Document
    from .fonts import DEFAULT_FONTS, FontRuleSet
    from .ops import normalize as normalize_op
    from .rules import RuleSet, default_ruleset

    if not args.out and not args.outdir and not args.dry_run:
        raise RuleError(u"要写结果就得给 --out 文件或 --outdir 目录"
                        u"（本工具**不会**覆盖原文件）")
    font_rules = FontRuleSet.load(args.fonts)
    problems = font_rules.check()
    if problems:
        raise RuleError(u"字体规则有问题：\n  - " + u"\n  - ".join(problems))
    if args.rules and os.path.exists(args.rules):
        text_rules = RuleSet.load(args.rules)
        rule_problems = text_rules.validate()
        if rule_problems:
            raise RuleError(u"上下标规则有问题：\n  - " + u"\n  - ".join(rule_problems))
    elif args.no_superscripts:
        text_rules = None
    else:
        text_rules = default_ruleset()
    with Document(args.path, writable_parts=("word/document.xml", "word/styles.xml",
                                            "word/numbering.xml")) as doc:
        report = normalize_op.format_normalize(doc, font_rules, text_rules,
                                               dry_run=args.dry_run,
                                               superscripts=not args.no_superscripts)
        out_path = None
        if not args.dry_run and report["total"]:
            if args.out:
                out_path = doc.save(os.path.abspath(args.out))
            else:
                out_dir = os.path.abspath(args.outdir)
                if not os.path.isdir(out_dir):
                    os.makedirs(out_dir)
                out_path = doc.save(os.path.join(out_dir, os.path.basename(doc.path)))
    lines = [u"文件：%s" % doc.path, u"步骤（固定顺序）："]
    for step in report["steps"]:
        if step["step"] == u"上下标规则":
            lines.append(u"  ① 上下标规则：%d 处" % step["total"])
            for key in sorted(step["counts"]):
                lines.append(u"        %-14s %d" % (key, step["counts"][key]))
        else:
            lines.append(u"  ② 字体与颜色：处理 %d 个有文字的 run ｜ 颜色改正 %d ｜ 去高亮 %d"
                         % (step["runs"], step["colors"], step["highlights"]))
            for key in sorted(step["fonts"]):
                lines.append(u"        %-34s %d 处" % (key, step["fonts"][key]))
    lines.append(u"")
    lines.append(u"合计 %d 项（各步相加，单位不同仅供概览）" % report["total"])
    lines.append(u"改写的部件：%s" % (u"、".join(report["parts"]) or u"（无）"))
    lines.append(u"")
    lines.append(u"这个宏里还有两件事**本工具不做**（说清楚免得你等）：")
    lines.append(u"  · 更新目录页码 —— 页码是排版结果，纯 XML 算不出（`PLAN.md` §6，另算）")
    lines.append(u"  · 询问「保存并关闭文档」—— 工具只管改，存哪/关不关是调用方的事（`PLAN.md` §8.3）")
    if args.dry_run:
        lines.insert(0, u"--dry-run：一个字节都没写")
    elif out_path:
        lines.insert(0, u"已写出：%s" % out_path)
    else:
        lines.insert(0, u"没有需要改的地方")
    return (dict(report, out=out_path), u"\n".join(lines))


def cmd_gui(args):
    """起本地网页 GUI（只监听 127.0.0.1）。"""
    from .gui import server as gui_server

    root = os.path.abspath(args.root) if args.root else os.path.expanduser(u"~")
    return {"ok": True}, gui_server.serve(host=args.host, port=args.port,
                                          open_browser=not args.no_browser, root=root)


def cmd_run(args):
    """按配方跑（CLI 版的一键执行；和 GUI 同一条路）。"""
    from . import pipeline as pipeline_mod

    recipe = []
    if args.recipe:
        import json as json_mod
        with io.open(args.recipe, "r", encoding="utf-8-sig") as handle:
            recipe = json_mod.load(handle).get("steps") or []
    elif args.steps:
        recipe = [step.strip() for step in args.steps.split(u",") if step.strip()]
    report = pipeline_mod.run_pipeline(args.path, recipe, mode=args.mode,
                                       out_path=args.out, dry_run=args.dry_run)
    return report, pipeline_mod.format_report(report)


def cmd_pdf(args):
    """导出 PDF：编排外部渲染器（Word / WPS / LibreOffice）。"""
    from .ops import pdf as pdf_op

    if args.list:
        renderers = pdf_op.detect_renderers()
        lines = [u"本机的 PDF 渲染器："]
        for item in renderers:
            lines.append(u"  %s %-14s %s"
                         % (u"✓" if item["available"] else u"✗", item["name"], item["detail"]))
        lines.append(u"")
        lines.append(u"（本工具只做编排，不自己渲染；一个都没装时会给出人话提示）")
        return ({"renderers": renderers}, chr(10).join(lines))

    if args.dry_run:
        plan = pdf_op.plan(args.path, args.out or u"", prefer=args.renderer)
        return ({"plan": plan},
                u"打算这么导（没真跑）：" + chr(10)
                + u"  " + plan["command"] + chr(10)
                + u"  渲染器：%s（%s）" % (plan["renderer"], plan["detail"]))
    report = pdf_op.export(args.path, args.out, prefer=args.renderer,
                           timeout=args.timeout, visible=not args.hidden)
    return report, pdf_op.format_report(report)


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
        elif args.command == "tablestyle":
            payload, human = cmd_tablestyle(args)
        elif args.command == "textfix":
            payload, human = cmd_textfix(args)
        elif args.command == "tidy":
            payload, human = cmd_tidy(args)
        elif args.command == "mdclean":
            payload, human = cmd_mdclean(args)
        elif args.command == "format":
            payload, human = cmd_format(args)
        elif args.command == "gui":
            payload, human = cmd_gui(args)
        elif args.command == "run":
            payload, human = cmd_run(args)
        elif args.command == "pdf":
            payload, human = cmd_pdf(args)
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
