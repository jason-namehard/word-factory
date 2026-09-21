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

    head = sub.add_parser("header", help=u"宏：表头格式统一（表X-Y + 空格居中 + 表格居中）")
    head.add_argument("path", help=u"要处理的 .docx")
    head.add_argument("--out", default=None, help=u"输出文件")
    head.add_argument("--outdir", default=None, help=u"输出目录（文件名与输入相同）")
    head.add_argument("--dry-run", action="store_true", help=u"只列出会怎么改")
    head.add_argument("--no-center-table", dest="center_table", action="store_false",
                      default=True, help=u"不要把表格居中")
    head.add_argument("--no-normalize-number", dest="normalize_number", action="store_false",
                      default=True, help=u"不要把编号统一成 表X-Y")
    head.add_argument("--all-captions", dest="only_before_table", action="store_false",
                      default=True, help=u"连带处理后面没跟表格的 表X-Y 段")
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


def cmd_header(args):
    """表头格式统一：算计划 → （非 dry-run 时）落地 → 报告。"""
    from .document import Document
    from .ops import header as header_op

    options = {"center_table": bool(args.center_table),
               "normalize_number": bool(args.normalize_number),
               "only_before_table": bool(args.only_before_table)}
    if not args.out and not args.outdir and not args.dry_run:
        raise RuleError(u"要写结果就得给 --out 文件或 --outdir 目录"
                        u"（本工具**不会**覆盖原文件）")
    with Document(args.path) as doc:
        report = header_op.apply(doc, options, dry_run=args.dry_run)
        lines = []
        if args.dry_run:
            lines.append(u"--dry-run：一个字节都没写")
        lines.append(u"文件：%s" % doc.path)
        lines.append(u"表题段 %d 个，其中 %d 个要改；表格居中 %d 张"
                     % (report["planned"], report["changed"], report["tables_centered"]))
        lines.append(u"")
        lines.append(u"%-10s %-9s %-9s %s" % (u"编号", u"原空格", u"新空格", u"表格名"))
        for detail in report["details"]:
            lines.append(u"%-10s %-9d %-9d %s"
                         % (detail["number"], detail["spaces_before"],
                            detail["spaces_after"], detail["name"][:34]))
        if not args.dry_run:
            if report["changed"] == 0 and report["tables_centered"] == 0:
                payload = {"dry_run": False, "changed": 0, "out": None,
                           "planned": report["planned"], "details": report["details"]}
                return (payload, u"\n".join(lines + [u"", u"没有需要改的地方"]))
            doc.mark_dirty()
            if args.out:
                out_path = doc.save(os.path.abspath(args.out))
            else:
                out_dir = os.path.abspath(args.outdir)
                if not os.path.isdir(out_dir):
                    os.makedirs(out_dir)
                out_path = doc.save(os.path.join(out_dir, os.path.basename(doc.path)))
            lines.insert(1, u"已写出：%s" % out_path)
            payload = {"dry_run": False, "changed": report["changed"], "out": out_path,
                       "planned": report["planned"],
                       "tables_centered": report["tables_centered"],
                       "details": report["details"]}
            return (payload, u"\n".join(lines))
        payload = {"dry_run": True, "changed": report["changed"],
                   "planned": report["planned"],
                   "tables_centered": report["tables_centered"],
                   "details": report["details"]}
        return (payload, u"\n".join(lines))


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
        elif args.command == "header":
            payload, human = cmd_header(args)
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
    return 0


if __name__ == "__main__":
    sys.exit(main())
