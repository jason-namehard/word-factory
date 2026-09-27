# 给 word 工厂加一个新功能（宏）——插件指南

> 这份文档就是界面「新增功能…」按钮弹出的那份提示。工具的成长性靠它：
> **可选功能不定死，后期随时扩**。

## 一句话

一个功能 = 一个 Python 模块（放进 `wordfactory/ops/`）+ 在 `wordfactory/pipeline.py` 的
`STEPS` 表里登记一行。**登记完重启 GUI，它就自动出现在「可跑的功能」列表里**——
左边勾选、拖动排序、存进执行方案，全部自动生效，界面代码一行不用改。

## 文件格式与架构

- **语言**：Python 3（本项目只用 **标准库**，不装第三方包——除非先跟用户说好）。
- **放哪**：`wordfactory/ops/你的功能.py`。文件名 = 功能的英文代号（如 `pagebreaks.py`）。
- **架构**：算子层不碰界面、不碰命令行——它只跟"打开的文档对象"打交道：

```python
# -*- coding: utf-8 -*-
"""一句话说明这个功能干什么。"""

def apply(document, params, dry_run=False):
    """对打开的文档执行本功能；返回一份**改动报告** dict。

    document : wordfactory.document.Document —— 用 document.part() 拿 XML 树，
               document.mark_dirty() 声明改过（保存时只重写改过的部件）；
    params   : dict —— 本功能的参数（STEPS 里登记的默认参数 + 用户覆盖）；
    dry_run  : True 时**一个字节都不改**，只报"会改多少"。
    """
    report = {"op": "你的功能", "changes": {}, "total": 0}
    # ...在这里干活...
    if not dry_run and report["total"]:
        document.mark_dirty()
    return report
```

## 登记（一行）

在 `wordfactory/pipeline.py` 的 `STEPS` 里加：

```python
STEPS = OrderedDict([
    ...
    (u"你的功能", (u"中文说明（会显示在界面上）", {"参数名": 默认值})),
])
```

- **名字用英文代号**（GUI 的勾选列表、执行方案、命令行 `run --steps` 都用它）；
- 中文说明写在第二个元素里，界面上跟在名字后面；
- 默认参数写第三个位置（dict），用户在执行方案里可以覆盖。

## 四条纪律（都是实测踩出来的，照做能少走弯路）

1. **幂等**：同一份文档跑两遍，第二遍必须报 0 处改动（"报要改就必须真改"，
   计数要跟眼睛看到的一致——假数字比不改还糟）。
2. **`dry_run` 不许改树**：要预演就在**深拷贝**上干活，或先算"要改哪些"再动手。
3. **只改该改的**：容器层是部件级保真——你 `mark_dirty` 哪个部件哪个才重写；
   别顺手把没改的部件也标脏。
4. **前置区保护**：封面/扉页/签字页里的空白行、空格是排版，清理类功能必须跳过
   （`wordfactory/frontmatter.py` 的 `protected_elements()` 现成可用）。

## 测试

`tests/` 下加一个测试文件，用 `tests/fixtures.py` 造最小文档，钉住：
幂等、dry-run 不动树、计数与实际一致。跑法：

```bash
"D:\Hermes\hermes-agent\venv\Scripts\python.exe" -m unittest discover -s tests -t .
```

## 两类不需要写 Python 的"功能"

- **数据类规则**（替换表、上下标、字体口径）：放 `rules/*.json`，改规则不用改程序；
- **表格款式**：`rules/tablestyle.json` 里的模板，界面「表格款式」页直接点选。
