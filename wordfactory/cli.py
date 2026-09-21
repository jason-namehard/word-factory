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
from .inspect import format_report, inspect
from .ooxml import PackageError


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
    return parser


def cmd_inspect(args):
    info = inspect(args.path)
    return info, format_report(info)


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv if argv is not None else sys.argv[1:])
    if not args.command:
        parser.print_help()
        return 2
    try:
        if args.command == "inspect":
            payload, human = cmd_inspect(args)
        else:
            log(u"未知命令：%s" % args.command)
            return 2
    except PackageError as exc:
        if args.as_json:
            emit_json({"ok": False, "error": u"%s" % exc})
        else:
            log(u"错误：%s" % exc)
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
