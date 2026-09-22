# -*- coding: utf-8 -*-
"""「段落配方」文本格式的读写（契约见 `docs/REFERENCE-MACROS.md` §3.5）。

配方是**嵌在文档末尾**的一段纯文本，长这样（生成端 `段落配方生成器.bas:98-124`）：

    （空行）
    === 段落配方 [土方计算] ===
    TEXT:1、
    VAR:1|数据表.xlsx!Sheet1!B2
    TEXT:本期
    VAR:2|数据表.xlsx!Sheet1!B3
    TEXT:万m
    TEXT:
    EXCEL_FILE:数据表.xlsx
    SHEET_NAME:Sheet1
    VARIABLE_COUNT:2
    === 配方结束 ===
    （空行）

**这几条逐字照抄参考宏，不能"改得更合理"**（否则 Word 宏读不了、或读出来的东西不一样）：

* 换行统一成一个 `?` 再切行（`段落重配.bas:47-50`）→ 任何换行风格都能读；
* 三个头部字段按**前缀长度**识别（`13`/`11`/`15` 个字符），值取 `Mid(line, 12)` 之后并 `Trim`
  （`段落重配.bas:60-69`）；**扫描的是全文**，不限于配方区间；
* `VARIABLE_COUNT` 必须 `IsNumeric` 才算数，且**三项齐备且变量数 > 0** 才算解析成功（`:75`）；
* 正文行只看区间（`=== 段落配方` 与 `=== 配方结束` 之间，`InStr` 判定，`:219-229`）；
* `TEXT:` 的区间是 `Mid(line, 6)`（前缀 5 字符）；**空 `TEXT:` = 一个换行**；
* **`VAR:` 只用来数序号**（`currentVar++`），行里写的那个编号与 `!B n` 地址**读端完全不看**
  （`段落重配.bas:255-262`），所以变量值的取用顺序 = 它在配方里出现的顺序；
* `TEXT:` 之前那段文本进了 Excel 的 A 列，**读端不读 A 列**（A 列只是给人看的上下文）。
"""

import re

#: 三个头部字段的前缀（长度照抄参考宏：13 / 11 / 15）
HEAD_EXCEL = u"EXCEL_FILE:"
HEAD_SHEET = u"SHEET_NAME:"
HEAD_COUNT = u"VARIABLE_COUNT:"
#: 区间标记（生成端写出来的就是这个样子，`段落配方生成器.bas:98/:124`）
SECTION_OPEN = u"=== 段落配方"
SECTION_CLOSE = u"=== 配方结束 ==="
#: 正文行前缀
TEXT_PREFIX = u"TEXT:"
VAR_PREFIX = u"VAR:"

#: 重建时缺变量值用的占位（参考宏 `段落重配.bas:261`）
MISSING = u"\u0023\u6570\u636e\u7f3a\u5931\u0023"          # #数据缺失#


class RecipeError(Exception):
    """配方文本有问题（人话错误）。"""


def split_lines(recipe_text):
    """按参考宏的规矩切行：换行统一成 ``?`` 再切，逐行 ``Trim``，空行丢掉。"""
    unified = (recipe_text or u"").replace(u"\r\n", u"?").replace(u"\r", u"?") \
        .replace(u"\n", u"?")
    return [line.strip() for line in unified.split(u"?") if line.strip()]


class Recipe(object):
    """一份配方：头部三项 + 正文行序列。"""

    def __init__(self, excel_file=u"", sheet_name=u"", variable_count=0,
                 name=u"", lines=None, warnings=None):
        self.excel_file = excel_file
        self.sheet_name = sheet_name
        self.variable_count = variable_count
        self.name = name                       # 只在标题行里有，Excel 里不存
        self.lines = list(lines or [])         # [("TEXT", 内容) | ("VAR", 序号) | ("RAW", 原文), …]
        #: 可疑但不致命的地方（例如变量数与正文里的 `VAR:` 行数对不上）。
        #: **不能直接报错**：参考宏遇到这种配方照样往下跑，多出来的变量补 `#数据缺失#`
        #: （`段落重配.bas:256-262`），报错反而会挡住用户本来能用的配方。
        self.warnings = list(warnings or [])

    # ---------------------------------------------------------------- 解析
    @classmethod
    def parse(cls, recipe_text):
        """解析配方文本；失败了抛 :class:`RecipeError`（把原因说清楚）。"""
        lines = split_lines(recipe_text)
        excel_file = sheet_name = u""
        variable_count = 0
        name = u""
        body = []
        in_section = False
        for line in lines:
            # 头部三项在**全文**扫（参考宏就是这么扫的）
            if line.startswith(HEAD_EXCEL):
                excel_file = line[len(HEAD_EXCEL):].strip()
            elif line.startswith(HEAD_SHEET):
                sheet_name = line[len(HEAD_SHEET):].strip()
            elif line.startswith(HEAD_COUNT):
                value = line[len(HEAD_COUNT):].strip()
                if _is_number(value):
                    variable_count = int(float(value))
            # 区间：**状态机**，与参考宏一致（`段落重配.bas:219-229`）。
            # 期间再出现一个开标记只是重新置位；已经不看的行不受影响。
            if SECTION_OPEN in line:
                in_section = True
                if not name:
                    match = re.search(u"\\[(.*?)\\]", line)
                    name = match.group(1).strip() if match else u""
                continue
            if SECTION_CLOSE in line:
                in_section = False
                continue
            if not in_section:
                continue
            if line.startswith(HEAD_EXCEL) or line.startswith(HEAD_SHEET) \
                    or line.startswith(HEAD_COUNT):
                continue                       # 头部行在区间里也要跳过（`段落重配.bas:232-234`）
            if line.startswith(TEXT_PREFIX):
                body.append(("TEXT", line[len(TEXT_PREFIX):]))
            elif line.startswith(VAR_PREFIX):
                body.append(("VAR", None))
            else:
                # **没有前缀的行也算正文**：`cleanText` 只去掉首尾的 CR/LF，TEXT 内容里的换行会留在
                # 配方里，于是那半截就成了一行"没有前缀"的文本。参考宏把它当"换行 + 原文"处理
                # （`段落重配.bas:264-268`），这里必须同样收下，不能丢。
                body.append(("RAW", line))
        problems = []
        if not excel_file:
            problems.append(u"缺少 EXCEL_FILE:")
        if not sheet_name:
            problems.append(u"缺少 SHEET_NAME:")
        if variable_count <= 0:
            problems.append(u"VARIABLE_COUNT 不是正数（%r）" % variable_count)
        if not body:
            problems.append(u"配方区间里没有正文行（TEXT:/VAR:）")
        if problems:
            raise RecipeError(u"配方解析失败：\n  - " + u"\n  - ".join(problems))
        warnings = []
        body_vars = len([1 for kind, _ in body if kind == "VAR"])
        if body_vars != variable_count:
            warnings.append(u"变量数对不上：头部说 %d 个，正文里有 %d 个 `VAR:` 行 —— "
                            u"多出来的变量会补 `#数据缺失#`（与参考宏一致）"
                            % (variable_count, body_vars))
        return cls(excel_file, sheet_name, variable_count, name, body, warnings)

    # ---------------------------------------------------------------- 还原
    def reconstruct(self, values):
        """按参考宏 `ReconstructParagraph`（`段落重配.bas:186-280`）拼出纯文本。

        ``values`` 是变量值列表（下标 0 = 第 1 个变量，也就是 xlsx 的 `B2`）。
        返回的字符串里用 ``\\n`` 分段 —— 每个 ``\\n`` 对应文档里的一个段落。
        """
        out = []
        previous_was_text = False
        at_line_start = True          # 上一动作已经"换过行"了吗（见下面数字规则的说明）
        index = 0
        for position, (kind, content) in enumerate(self.lines):
            if kind == "RAW":
                # 无前缀行 = 换行 + 原文（`段落重配.bas:264-268`）
                out.append(u"\n" + content)
                previous_was_text = True
                at_line_start = False
                continue
            if kind == "TEXT":
                if content == u"":
                    # 空 `TEXT:` 表示一个换行。**但紧跟一条无前缀行时不再补换行** ——
                    # 那条无前缀行自己就带着换行（生成端写 `TEXT:<空>` + `后半句` 表示的
                    # 其实是**一个**段落分隔，参考宏会把两者都算成换行 → 凭空多出一个空白段）。
                    # 这是用户 2026-09-22 反馈的"突然多出一个回车"里的一类，我们不复刻这个缺陷。
                    if self._next_is_raw(position):
                        previous_was_text = False
                        continue
                    out.append(u"\n")
                    previous_was_text = False
                    at_line_start = True
                    continue
                # 「2、xxx」这种编号段要另起一行。**但如果刚刚才换过行（前面是空 TEXT:），
                # 就不能再来一个** —— 否则会凭空多出一个空白段（参考宏就是这么多的，
                # 也正是用户反馈的"突然多出一个回车"）。
                if not previous_was_text and not at_line_start and _starts_with_number(content):
                    out.append(u"\n")
                out.append(content)
                previous_was_text = True
                at_line_start = False
            else:
                index += 1
                if index <= len(values):
                    value = values[index - 1]
                else:
                    value = MISSING
                out.append(u"" if value is None else u"%s" % value)
                previous_was_text = False
                at_line_start = False
        return u"".join(out)

    def _next_is_raw(self, position):
        """下一条正文行是不是"无前缀行"（它自带换行）。"""
        return (position + 1 < len(self.lines) and self.lines[position + 1][0] == "RAW")

    def footer(self):
        """配方尾部的三个头部字段（生成端写在正文之后）。"""
        return [u"%s%s" % (HEAD_EXCEL, self.excel_file),
                u"%s%s" % (HEAD_SHEET, self.sheet_name),
                u"%s%d" % (HEAD_COUNT, self.variable_count)]

    def to_text(self):
        """还原成配方文本（生成端格式：标题行 + 正文 + 尾部字段 + 结束行）。

        `VAR:` 行照参考宏的完整写法 `VAR:n|文件!表!B(n+1)` —— 虽然读端不看这些字段，
        但**人要看、也比对得上**（`段落配方生成器.bas:189` 等处就是长这样）。
        """
        lines = [u"%s [%s] %s" % (SECTION_OPEN, self.name or u"默认配方", u"===")]
        number = 0
        for kind, content in self.lines:
            if kind == "TEXT":
                lines.append(u"%s%s" % (TEXT_PREFIX, content))
            elif kind == "RAW":
                lines.append(content)              # 没前缀的行照原样写
            else:
                number += 1
                lines.append(u"%s%d|%s!%s!B%d"
                             % (VAR_PREFIX, number, self.excel_file, self.sheet_name, number + 1))
        lines.extend(self.footer())
        lines.append(SECTION_CLOSE)
        return u"\r\n".join(lines)


def _is_number(text):
    try:
        float(text)
        return True
    except (TypeError, ValueError):
        return False


def _starts_with_number(text):
    """首字符是不是数字（对应参考宏的 `IsNumeric(Left(textContent, 1))`）。

    用 `str.isdigit()`：ASCII 数字与全角数字都算 —— VBA 的 `IsNumeric` 在这些情况下也是真。
    """
    return bool(text) and text[0].isdigit()
