# word 工厂（word-factory）

**外置的 Word 文档批量规范化工具箱**：不依赖 Word、不依赖 WPS、不运行 VBA 宏，
直接读写 `.docx` 的 OOXML（zip + XML，**只用 Python 标准库**），把一批文档按你排好的顺序一次改到位。

> 立项：2026-09-21，用户裁决。前身是 `E:\Zspace\projects\MacroToolbox`（Word 宏工具箱，用 VBA 宏实现同样的目的）——
> 那个项目已经交付并验收通过，仍然可用；本项目是**换一条路实现同一批功能**，不修改它。

## 为什么换一条路

| | 依赖 Word 的宏工具箱 | 本项目的 XML 路线 |
|---|---|---|
| 需要 Word/WPS | 需要 | **不需要** |
| 宏与安全提示 | 有（宏被拦就没法用） | 没有 |
| 批量速度 | 一个文档起一次 Word | 纯文件读写，几十上百个文件也快 |
| 结果可复现 | 受 Word 版本/设置影响 | **同一输入同一输出**（可 diff、可入库） |
| 目录页码 | Word 更新（会遇到"重置目录"的老毛病） | **做不到**（见下面的边界） |
| 复杂排版渲染 | Word 说了算 | **做不到**（同上） |

## 三个名词（先说清楚，后面全靠它）

- **宏**：一个**具名的、带参数的修改步骤**。名字沿用你的说法叫"宏"，但**它不是 VBA 宏**——
  例如 `格式规范化`、`表格去空格`、`段落配方生成`、`表头格式统一`、`特殊字符替换`。
- **配方（pipeline）**：把若干宏**勾选 + 排序**后的执行清单，一次跑完。
- **作业（job）**：配方 + 一批文件（可给目录），跑完给每个文件一份**改动报告**。

## 明确不做（这几条是"做不到"，不是"先不做"）

1. **算页码**。目录里的页码是**排版结果**——谁在第几页取决于行距、字体、图、分页符……
   XML 里没有这个信息，**任何纯 XML 工具都算不出页码**。目录页码这件事**单独另算**（见 `PLAN.md` 的"页码"一节）。
2. **自己渲染成 PDF**。`.docx → .pdf` 需要**排版/渲染引擎**，XML 层做不到。
   本工具做的是**编排**：挑文件、命名、输出目录、挑渲染器、超时与失败报告
   （`pdf` 命令 + GUI 的「导出 PDF」按钮）。渲染本身交给本机装着的引擎——
   Word / WPS / LibreOffice，装了哪个用哪个，一个都没装就明说（见 §PDF 导出）。
3. 不处理老的 `.doc`（二进制格式）——先用别的工具转成 `.docx`。
4. 不动 `.docm` 里的 VBA 工程（只改文档正文部分，宏部件原样保留）。

## 参考与来源

- `reference/vba-macros/`：**功能规格的真正来源**——12 个已经在用的 VBA 宏（原件在
  `D:\百度网盘\宏相关（主机）\word宏`，此处为只读参考副本）。
  逐宏规格见 `docs/REFERENCE-MACROS.md`。
- `reference/delivered-macros/`：旧项目从真 `Normal.dotm` 导出的 5 个宏（`*claw`），
  同样只作参考，不再维护。

## 目录

```
wordfactory/        代码（纯标准库；不装任何第三方包）
  gui/               本地 Web 界面（stdlib http.server + 前端页面）
  ops/               算子（宏）：题注 / 上下标 / 表格 / 文本 / 整理 / PDF
  pipeline.py        配方引擎（GUI 与命令行共用）
docs/                规格、架构、验收策略、未决问题
gui-prototype/       GUI 布局设计稿（SVG 原型，定稿用的，不参与运行）
reference/           VBA 宏原件（只读参考，不参与运行）
tests/               单测与差分测试夹具
tmp/                 临时件（不入库）
```

## 怎么跑

```bash
# 1) 看一个文档的结构：部件清单 + 段落/表格/run 统计（只读）
"D:\Hermes\hermes-agent\venv\Scripts\python.exe" -m wordfactory.cli inspect "某文档.docx"

# 2) 逐段看「逻辑文本 ↔ run」的映射 —— 判断一个宏能不能做对，先看这个
"D:\Hermes\hermes-agent\venv\Scripts\python.exe" -m wordfactory.cli text "某文档.docx" --limit 20
#    --grep 表5-1     只看包含这个词的段落

# 3) 上下标规则（外置规则文件，可以用任何编辑器改；工具按正则查找并应用）
"D:\Hermes\hermes-agent\venv\Scripts\python.exe" -m wordfactory.cli rules init     # 写出默认规则
"D:\Hermes\hermes-agent\venv\Scripts\python.exe" -m wordfactory.cli rules show     # 看规则表
"D:\Hermes\hermes-agent\venv\Scripts\python.exe" -m wordfactory.cli rules check    # 校验（改完先跑这个）
"D:\Hermes\hermes-agent\venv\Scripts\python.exe" -m wordfactory.cli rules apply "某文档.docx" --dry-run
"D:\Hermes\hermes-agent\venv\Scripts\python.exe" -m wordfactory.cli rules apply "某文档.docx" --outdir out
```

### 上下标规则文件长什么样（`rules/subscripts.json`）

两种写法混用，顺序即优先级，**先匹配先应用**（前面规则占了的字，后面规则不会再动）：

```jsonc
// 1) 字面量 + 逐字符类型：复刻参考宏 智能上下标.bas 的字典
{"id": "m2", "match": "m2", "kinds": "NS", "note": "m2 —— m 正常、2 上标"}
//   kinds 与 match 等长：N 正常 / S 上标 / B 下标

// 2) 正则 + 作用目标：用户自己要的"用正则查找并应用"
{"id": "unit-m", "pattern": "m([2-9]+)", "target": "group:1", "kind": "superscript",
 "enabled": false, "note": "通用上标：m 后跟 2-9 就上标（早期宏的做法，m1 不动）"}
//   target 也可以是 "all"（整段匹配都设）；还有 not_before / not_after 表达边界约束
```

> 默认规则取自 `智能上下标.bas:74-97` 的字典（Vmax/Qpl/m2..m5/cm2/km2/qm/Qm/KP/CV/H24P/hR/H24/°C/mm）。
> 「m 后加数字无脑上标」那条**通用上标**是早期宏的局限，默认关着 —— 字典规则已经覆盖 m2..m5，
> 要无脑覆盖 m6..m9 就把它打开。

（其余能力见 `PLAN.md` 的阶段表；每个阶段做完都会在这里补用法。）

### 题注统一（`captions`）—— 表题与图题是**两套规矩**

|  | 表题 | 图题 |
|---|---|---|
| 例子 | `表4.2-1` + 空格 + `水库洪水计算成果表` | `图4.2-1` + **1 个空格** + `欧峪水库30年一遇调洪演算图` |
| 段首缩进 | 缩进 2 字符（`w:firstLineChars="200"`，**不是空格**） | **清掉**缩进（不然后面居中会被顶偏） |
| 对齐 | 左对齐，靠**空格**把名字推到版心中间（**规则 B**） | **整段居中**（`w:jc="center"`） |
| 编号 | 统一成 `表X-Y`（编号里的空格去掉） | 统一成 `图X-Y` |
| 顺带 | 把紧跟其后的表格居中（`w:tblPr/w:jc`） | —— |

> 两套规矩的差别是用户 2026-09-21 明确的：
> 「图是"图4.2-1 欧峪水库30年一遇调洪演算图"整个居中，从 图这个字开始就居中，然后中间空一个格子」。
> 规则 B（表题空格数）是从**你现有 12 个真实表题**反推出来的：名字中点落在「编号之后 → 版心右边界」的中点，
> 实测平均差 +0.67 格、10/12 落在 ±3 格内（`tests/test_op_header.py` 里有这条不变量）。

### 两版输出：`--mode verify`（验证版）/ `--mode formal`（正式版）

```bash
python -m wordfactory.cli captions "某文档.docx" --mode verify --outdir out   # 验证版：改过的地方标蓝
python -m wordfactory.cli captions "某文档.docx" --mode formal --outdir out   # 正式版：通体黑 + 字体合规
```

- **验证版**：把**工具动过的那一段**（题注的「编号 + 中间空格」）标成蓝色 `0000FF`，其余一个字都不动 ——
  你在 Word 里打开就能判断改得对不对。**名字部分不涂蓝**（名字没被改），所以蓝色边界 = 改动边界。
  原件里本来就有的蓝色（这份文档标题里有 1 处）工具会明说"不是我标的"。
- **正式版**：所有文字改黑 + 去真高亮 + 按规则换字体，然后**把写出来的文件重新体检一遍**，末行给 `AUDIT=PASS`。
- **报"要改"就必须真改**：已经合规的题注报 0 处、不写文件。有单测钉住这条——
  实测踩过一次"图题每遍都报要改"（原因是备注写死了，`changed` 又看备注），数字跟眼睛看到的不一致就是 bug。

### 体检（`audit`）—— 判据可复算，不看工具自己的报告

```bash
python -m wordfactory.cli audit "某文档.docx"     # 末行 AUDIT=PASS / AUDIT=FAIL；FAIL 时退出码 1（可当闸门）
```

它**重新打开文件**，按 OOXML 的继承链（run 自己的 `rFonts` → 字符样式 → 段落样式 → `docDefaults` → 主题字体）
算出每个有文字的 run **实际生效**的字体，再逐项核对：不合格字体 / 非黑颜色 / 真高亮。分两档报：

- **FAIL**：会显示成文字的问题（生效字体不合格、非黑颜色、真高亮）；
- **提示**：当前不显示成文字、但留着旧字体的地方（段落标记、样式定义、编号表、字体清单）；
  主题字体里如果有不合格字体，本工具**不改主题**，也会把这一点明说。

### 字体规则文件（`rules/fonts.json`）——"不合格的字体"由你定

```jsonc
{
  "keep": ["宋体", "黑体", "楷体", "Times New Roman"],   // 原样保留
  "replace": {"仿宋": "宋体", "仿宋_GB2312": "宋体"},     // 明确映射（四个属性都换）
  "default": "宋体",                                    // 其余**中文**字体统一成它
  "default_latin": "Times New Roman",                   // 其余**西文**字体统一成它
  "default_scope": "all",                               // all=西文也统一；eastAsia=西文一概不碰
  "symbol_fonts": ["Symbol", "Wingdings", "Webdings", "Marlett", "ZapfDingbats", ...],  // 永不触碰
  "black_all": true, "remove_highlight": true
}
```

**现口径是用户 2026-09-21 拍的**（原话：「西文统一为 Times new roman，中文常见的几种基本都是
宋体、仿宋GB2312、黑体、楷体」）：中文留 宋体/黑体/楷体、仿宋（含 `_GB2312`）→ 宋体、其余中文兜成宋体；
西文统一 Times New Roman；符号字体永不碰。

三条实现口径（都是实测逼出来的，改了会影响结果）：

1. **`w:rFonts` 有四个属性**（`ascii`/`hAnsi`/`eastAsia`/`cs`），四个都换 ——
   只换两个，实测还会剩 738 处旧字体。
2. **样式表与编号表也要换**：run 自己不写字体时，字体是**继承**来的
   （实测这份文档的字符样式 `26` 就是仿宋，只改正文会漏）。
3. **中文与西文各有各的兜底，且符号字体永不碰** —— 西文兜成宋体是错的（宋体不是西文字体）；
   符号字体（Symbol/Wingdings）换成 TNR/宋体后 ✔ ➜ ★ 会掉字形。
   > 实测（这份文档）：仿宋 → 宋体 1729 处、Arial → Times New Roman 18 处、Tahoma → Times New Roman 12 处。

### 段落配方（`recipe`）—— 与 Word 宏**互读同一套格式**

参考宏 `段落配方生成器`（把选中的段落变成"配方" + 写一份 `.xlsx`）与 `段落重配`（读配方 + 读 xlsx → 重建段落）。
本工具照 `docs/REFERENCE-MACROS.md` §3 的契约实现，**同一份配方文本、同一份 xlsx，两边都能读**：

```bash
python -m wordfactory.cli recipe show "带配方.docx"        # 解析并打印配方（也支持直接给 .txt）
python -m wordfactory.cli recipe gen "模板.docx" --name 土方计算
#   把正文里的高亮片段识别成变量 → 写 数据表.xlsx + 把配方追加到文档末尾
#   --mode chars --char xx   改用"占位符字符串"识别；--no-append 只写数据表不写文档
python -m wordfactory.cli recipe rebuild "带配方.docx" --out 重建后.docx
#   --recipe-file F  配方不在文档里时单独给；--xlsx X / --data-dir D 指定数据表位置；--dry-run 先演练
```

契约（**逐条照抄参考宏，不"改得更合理"**）：

| 项 | 规则 |
|---|---|
| 配方文本 | `=== 段落配方 [名] ===` … `=== 配方结束 ===`，正文行只有 `TEXT:` 与 `VAR:` 两种 |
| 尾部字段 | `EXCEL_FILE:` / `SHEET_NAME:` / `VARIABLE_COUNT:`（**全文扫**，不限区间） |
| 切行 | 任何换行风格都统一成 `?` 再切、逐行 `Trim`、空行丢掉 |
| 数据表 | **一个表**（名字 = `SHEET_NAME:`）、`A1=项目` `B1=数值`、**第 n 个变量在第 n+1 行 B 列** |
| `VAR:` 行 | 读端**只数序号**，行里写的编号与 `B n` 地址一概不看（取用顺序 = 出现顺序） |
| A 列 | 是"给人看的上下文"，**读端不读** |
| 值不够 | 补 `#数据缺失#`（不是报错） |

已知差异（会写在报告里，不藏着）：① 参考宏把结果追加到**当前文档**，我们追加但**写新文件**（绝不覆盖输入）；
② 参考宏用 Excel 的 `AutoFit` 定列宽，XML 层没有等价写法，我们按内容**估**宽度；
③ 参考宏靠**选区**决定处理范围，我们靠**规则**（用户 2026-09-21 已裁决接受）：模式 1 = 正文里所有高亮片段，
模式 2 = 指定的占位符字符串；已存在的配方区间会自动跳过，免得把上一份配方当正文再生成一遍；
④ 参考宏识别高亮时有个「**隐形修正**」（把高亮片段**去掉最后一个字符**，`段落配方生成器.bas:177-213`）——
那是为了绕开 Word 在边界字符上读 `HighlightColorIndex` 的偏差。XML 里高亮状态是**精确**的，
多砍一个字符反而会把值改错（高亮的 `3.5` 会变成 `3.`），所以**默认不照抄**；
要跟宏逐字一致就加 `--trim-last-char`。
> ④ 这条**还没与金标准比对过**（规格书自己标注"必须用 Word+宏跑一份再对比，不能靠读代码断言"）。
> 手里有一个"宏生成的文档 + 它的数据表"样本就能立刻settle，见 `PLAN.md` §5 第 8 条。

### 表格去空格/回车（`tableclean`）—— 宏：表格空格回车删除

```bash
python -m wordfactory.cli tableclean "某文档.docx" --dry-run        # 先看会改多少
python -m wordfactory.cli tableclean "某文档.docx" --level 1 --outdir out
```

三档（照参考宏 `表格空格回车删除.bas:144-158`）：`1` 仅空格（半角 + 不间断空格）、
`2` 仅回车（段落标记 + 手动换行）、`3` 两者（**默认**）。

**换行口径（用户 2026-09-22 裁决 = 保守 + 数字例外）**：只删「**首尾的**」和「**连续重复的**」换行，
**保留单个内部换行** —— 实测你的报告里那 9 格全是"为了让窄列好看而故意折行"的表头
（`防洪标↵准`、`大坝右侧/开敞↵式`、`实测库容↵（万m³）`），不该被拉直。
例外：**整格是纯数字的单元格**（`37.47` 被断成两行）→ 换行一律删。
要完全照抄宏（所有换行一律删）加 `--exact-macro`。

> 实测对比（你的报告，12 张表 / 554 格）：**新口径 0 格需要动**（9 格被正确认成"故意的"，
> 其中 6 个内部换行被保留、3 个两段式表头原样）；`--exact-macro` 则是 9 格、删 6 字符 + 合并 3 段。

**两处有意与宏不同**（都要知道，不然会以为做错了）：

1. **只删字符、保留 run 级格式**。宏是整格写回（`：162`），单元格里的字体/加粗/上下标/颜色会被压成一种；
   XML 层只删字符，结果比宏**更保真**。要跟宏一样压平就加 `--flat`。
2. **默认不删全角空格 `U+3000`**（宏也删不掉它）。中文报告里全角空格常被用来做段首缩进。

> ⚠ **档位 2/3 会删掉"为了让窄列好看而手动折行"的换行**。实测你自己的报告
> （`欧峪水库…复核报告.docx`）：12 张表 554 个单元格里，档位 1 命中 **0** 格，
> 档位 2/3 命中 **9** 格 —— 而这 9 格**全都是故意折行的表头**：
> `防洪标\n准`、`控制泄\n洪`、`防洪高水\n位`、`大\n坝`、`大坝右侧/开敞\n式`（×2）、
> `实测库容\n（万m³）`（×3）。**动手前先看清楚这些要不要保留。**

### 表格款式（`tablestyle`）—— 外置模板 + 采集 + 预览 + 遍历表头挑表

用户 2026-09-22 定的口径（原话）：

> 「常见的表格格式一般是外框1.5，内框是标准（例如0.5磅），**一定是外框粗内框细**。
>   然后一般数字结果都是居中排版，第一行的标题一定居中…**只有长文本，例如说明、备注左对齐，
>   其余一律居中**。长文本的判别方式是一行表格放不下，它需要在一个单元格内另起一行，这就用左对齐方案」

```bash
python -m wordfactory.cli tablestyle init                        # 写出 rules/tablestyle.json（6 个模板）
python -m wordfactory.cli tablestyle show                        # 看模板参数
python -m wordfactory.cli tablestyle list 报告.docx               # **遍历表头**：序号/行数/列数/列宽/表头
python -m wordfactory.cli tablestyle preview --out 预览.docx       # 每种模板各渲染 3 张表，用 Word 挑
python -m wordfactory.cli tablestyle apply 报告.docx --style 通用款·外粗内细 --tables "水位,库容" --outdir out
python -m wordfactory.cli tablestyle capture 报告.docx --table 3 --name 我调的 --out rules/tablestyle.json
```

**6 个出厂模板**（命名沿用你的习惯）：

| 模板 | 特点 |
|---|---|
| `通用款·外粗内细` | **默认**：外框 1.5 磅 / 内框 0.5 磅；表头居中加粗跨页重复；其余一律居中；**只有一行放不下的长文本左对齐** |
| `通用款·全居中` | 偷懒款：不分长短全居中（"不会犯错，谈不上最好看但不至于丑"） |
| `通用款·均布列宽` | 默认 + 各列等宽（均布单元格） |
| `通用款·自适应列宽` | 默认 + **按内容自适应**：每列至少放得下表头，其余按内容分；**允许总宽超出页边距** |
| `三线表·学术款` | 上下粗线 + 表头下细线、无竖线；其余同默认款 |
| `三线表·学术款·均布列宽` | 三线表 + 各列等宽 |

**长文本是按宽度判的，不是按字数**：算这一格文字的实际宽度（中文全角、其余半角），比它所在列
（含合并单元格的跨列）减掉左右内边距后的可用宽度还宽 → 一行放不下 → 左对齐。宽度算不出来时
（表没有 `w:tblGrid`）**保守判"不是长文本"**，即保持居中。

**按表头叫表格**（你要的"遍历表头、表头左边复选框"）：`tablestyle list` 列出每张表的
序号 + 表头文字；`apply --tables "水位,库容"` 就是**按表头关键词挑表**（表头里出现这些词的都选中），
也照旧支持 `all` / `1,4-6`。将来 GUI 的那排复选框，数据源就是这个清单（`--json` 下是结构化数组）。

**表头折行**：`--wrap-header 备注`（可重复，给表头文字或列号）把那个表头折成两行展示，
折点在**中点附近、避开数字与标点**（不会把 `1985`、`3.3` 劈开）—— 折行后那一列可以窄一半，
自适应就更容易排得开。

> 实测你的报告（12 张表）：`通用款·外粗内细` 命中 **139 处**，其中**只有 9 个格子**因"一行放不下"
> 改成左对齐（其余本来就居中，所以改动很轻）；**只重写 `word/document.xml`**，其余 41 个部件
> 逐字节未变；**791 段文字一字未动**；外框 12（1.5 磅）/ 内框 4（0.5 磅）✓ 外比内粗 ✓。
> 样例：[模板预览 v3](E:\Zspace\output60922_wordfactory_表格模板_预览v3.docx)、
> [默认款样例 v3](E:\Zspace\output60922_wordfactory_表格_默认款样例v3.docx)

### 格式规范化（`format`）—— 一条命令跑完一串步骤

宏 `格式规范化.bas` 的对应实现（§2.7）。它是本项目**第一个"配方"**：
把已有步骤按固定顺序串起来跑 —— 这也是 M4「勾选 + 排序 + 一键执行」的雏形。

```bash
python -m wordfactory.cli format "报告.docx" --dry-run          # 先看会改多少
python -m wordfactory.cli format "报告.docx" --outdir out
#   --fonts F 字体规则；--rules R 上下标规则；--no-superscripts 只做颜色/字体
```

固定顺序与理由：**先上下标规则 → 再字体与颜色**。上下标会把一个 run 切成几段
（`split_run_at`，两半各留 `rPr`），先把结构改完再统一刷颜色/字体，报告里的数字更好解释。

**宏里有两件事本工具明确不做**（命令输出里也会写一遍，免得你等）：

| 宏做的 | 为什么不照做 |
|---|---|
| 更新目录页码 | 页码是**排版结果**，纯 XML 算不出来（`PLAN.md` §6，单独另算） |
| 询问"保存并关闭文档" | 工具只管改；存哪、覆不覆盖、关不关是调用方的事（`PLAN.md` §8.3） |

> 实测你的报告：`format` 报出上下标规则 **0 处**、字体与颜色 1761 项
> （仿宋→宋体 1729、西文→Times New Roman 23、颜色改正 1、去高亮 1）；
> 改写的部件 = `word/document.xml` + `styles.xml` + `numbering.xml`（其余逐字节未变）。

### 一键整理（`tidy`）/ Markdown 清理（`mdclean`）/ 文本替换与对齐（`textfix`）

```bash
python -m wordfactory.cli tidy "报告.docx" --dry-run
python -m wordfactory.cli mdclean "报告.docx" --dry-run
python -m wordfactory.cli textfix "报告.docx" --rules rules/replace.json --dry-run
```

- **`tidy`**：段尾空格、空白行、首尾空白段、连续空格；外加照 Copy++ 的
  `--merge-lines`（合并换行）与 `--remove-spaces`（去空格），后者**默认跳过题注段**
  （题注的空格是排版用的，删了就歪）。每条都能单独关（`--no-blank-lines` 等）。
- **`mdclean`**：清 Markdown 痕迹（`**加粗**`→正文、`` `代码` ``→中文引号、段首 `#`、
  列表符号、多余星号、双空格）。替换是**局部**的——同一段里没碰到的字保持原 run 不动。
- **`textfix`**：外置 JSON 替换表（跨 run 替换，格式跟第一个 run 走）+ 把两端对齐改回左对齐
  （两端对齐在短行上会把字拉开，很丑）。

### 配方编排（`run`）—— 勾选 + 排序 + 一键执行

```bash
python -m wordfactory.cli run "报告.docx" --steps captions,tidy --mode verify --dry-run
python -m wordfactory.cli run "报告.docx" --steps captions,tidy --mode formal --outdir out
python -m wordfactory.cli run "报告.docx" --recipe 配方.json --mode formal   # 配方文件（可排序、可勾选）
```

`--steps` 的顺序**就是执行顺序**（不是固定的）；`--mode verify` 只把改过的地方标蓝，
`--mode formal` 收尾做通体黑 + 字体合规 + 体检。配方 JSON 长这样：

```jsonc
{"name": "报告规范化", "steps": [
  {"id": "captions", "enabled": true},              // 题注统一
  {"id": "sup", "rules": "rules/subscripts.json"},  // 上下标规则
  {"id": "tableclean", "level": 3},
  {"id": "textfix", "rules": "rules/replace.json"},
  {"id": "mdclean"},
  {"id": "tidy", "merge_lines": false}
]}
```

GUI 与命令行**共用同一套引擎**（`wordfactory/pipeline.py`）——GUI 上勾选、拖动排序，
存下来的配方文件命令行原样能跑。

### PDF 导出（`pdf`）—— 编排外部渲染器

```bash
python -m wordfactory.cli pdf --list                    # 本机有哪些渲染器
python -m wordfactory.cli pdf "报告.docx" --dry-run     # 只报打算怎么导
python -m wordfactory.cli pdf "报告.docx" --renderer word --timeout 300
```

- 渲染器按 Word → WPS → LibreOffice 的顺序挑第一个装了的；`--renderer` 可指定。
- **页数从 PDF 自己数**（读 `/Count`），不信 Word 的 `ComputeStatistics`——
  实测同一份文档它报 1 页，PDF 里其实是 24 页。
- 用 COM 驱动时**只读打开、导完就关、绝不去 kill 进程**；超时会明说"窗口可能还开着"，
  并告诉你手动关掉即可。
- **只有 `.docx` 本身没问题才导得出来**——所以导出成功同时也是"文件没被我们写坏"的一条实证。

## 保真与兼容性：两处会悄悄弄坏文档的坑

工具每写一个文件都靠两条不变式兜底（`tests/test_ooxml.py` 钉住）：

1. **只重写改过的部件**：没碰的部件按**原字节**复制（连 zip 的时间戳、压缩方式都不变），
   所以"除了正文别的都没动"可以被逐字节验证。被改的那个部件会整份重序列化，
   想证明"没误伤"要靠段落级 / 元素级核对，不能只看字节数。
2. **命名空间声明一个都不能少**：`mc:Ignorable="w14 w15 wp14"` 这类属性**引用前缀**，
   而 ElementTree 只给"树里真用到的"命名空间发声明。Word 习惯在根上多声明几个备用的
   （`wp14` 常常全文一次都没出现），声明一丢，`Ignorable` 就指向未声明前缀——
   **Word 打开时报"文件可能已经损坏"**。实测踩过：那一刻工具产出的每个 `.docx`
   （连只跑一步 `tidy` 的）都打不开，而原样复制的部件一点问题没有。
   修法是保存时按原件把根上的声明补齐；两条回归测试盯着这条。

## 界面（GUI）

```bash
python -m wordfactory.cli gui                 # 起本地服务，默认 http://127.0.0.1:8792
python -m wordfactory.cli gui --port 9000 --no-browser
```

浏览器里能干的事：**选文件 → 勾宏 → 拖拽排序 → 选表格款式 / 按表头挑表 → 跑验证版或正式版 → 看改动报告 → 体检 → 导出 PDF**。
配方可存成 JSON 文件，命令行 `run --recipe` 原样能跑（两边共用 `pipeline.py`）。
表格款式页的复选框，数据源就是 `tablestyle list` 的同一份清单。

布局不是拍脑袋定的：先用 `gui-prototype/index.html`（SVG 拖拽设计稿）把界面排到用户满意，
再照它写成真界面。**服务只绑 127.0.0.1**，下载走白名单，退出按钮即关服务。

## 状态

**M1 内核 / M2 三个宏 / M3 题注与配方 / M4 配方编排 + GUI / M5 PDF 编排全部落地**；
目录（TOC）外观按用户意见**押后**，页码仍"另算"。305 条单测全绿。

- [x] 仓库与参考件入库
- [x] `docs/REFERENCE-MACROS.md`（12 个参考宏的逐宏规格）
- [x] M1a 容器层：只重写改过的部件（保真）+ OOXML 前缀注册
- [x] M1b 文本层：**跨 run 替换**（一句话被 Word 切成多块也能正确替换，格式跟第一个 run 走）
- [x] M1c 切 run：只给区间内的字设属性（切完两半各自保留 rPr）
- [x] **上下标规则外置接口**（用户 2026-09-21 要求）：JSON 规则文件 + 正则/字面量两种写法
      + `rules check/show/apply`，先匹配先应用、幂等（重跑报 0 处）
- [x] **M3a 题注统一**（表题规则 B + 图题整段居中）—— 在用户真文档副本上实测：
      791 段里只有 17 段（12 表 + 5 图）文字有变，且差异**只在空格上**；其余 39~41 个部件逐字节未变
- [x] **两版输出**：验证版标蓝（蓝 run 18 个 = 我标的 17 + 原件本来有的 1）/ 正式版通体黑
- [x] **`audit` 体检**：独立复算，正式版输出 `AUDIT=PASS`（生效字体里没有不合格的）
- [x] **字体口径已由用户拍定**（2026-09-21）：中文留 宋体/黑体/楷体、仿宋→宋体；西文统一 Times New Roman
- [x] **M3b 读的一半：`recipe show` / `recipe rebuild`** —— 纯标准库读写 `.xlsx`（不用 openpyxl）
      + 配方文本格式按 §3 契约逐条照抄；实测"带配方的文档 → 重建"追加 8 段、原有段落一字不动、
      只重写 `word/document.xml`
- [x] **M3b 写的一半：`recipe gen`** —— 高亮 / 占位符两种识别模式、写 xlsx、配方追加到文档末尾。
      实测端到端：模板（2 处高亮）→ 生成（变量 2 个、数据表 3.5/12.8、配方 11 段）→
      换新数据（88.8/246.0）→ 重配 → 段落骨架与模板一致、数字已替换
- [x] M3b 收尾：**与宏产出的真实文件比对通过**（用户 2026-09-22 给的样本，见
      `docs/GOLD-STANDARD.md`）：我们的读端读得进宏的 xlsx 与配方、重建结果与宏自己写的 39 个
      内容段落**逐字一致**；「隐形修正」经用户明确是**宏的缺陷**，故默认不照抄
- [x] **M2 三个宏齐了**：`tableclean`（表格去空格/回车，三档 + 保守口径）、
      `textfix`（外置替换表 + 两端对齐改左对齐）、`tidy`（段尾空格/空白行 + 照 Copy++ 的
      `--merge-lines` / `--remove-spaces`）、外加 `mdclean` 清 Markdown 痕迹
- [x] **表格款式**：外置模板（6 款）+ 采集 + 预览 + 按表头关键词挑表 + 表头折行
- [x] **M4 配方编排**：`pipeline.py` 引擎（勾选 + 排序 + 一键执行 + 改动报告），
      命令行 `run` 与 GUI 共用
- [x] **M4 GUI**：本地 Web 界面（选文件/勾宏/拖拽排序/表格款式/配方存取/预览/报告/体检/导出 PDF），
      布局经 SVG 设计稿定稿；只绑 127.0.0.1
- [x] **M5 PDF 编排**：`pdf` 命令 + GUI 按钮，Word / WPS / LibreOffice 三选一，
      页数从 PDF 自己数，超时与失败给人话
- [x] **保真与兼容性**：未改部件逐字节不变；`mc:Ignorable` 前缀补齐（修掉"Word 打不开产物"的坑）
- [ ] 目录（TOC）外观：**押后**（用户 2026-09-22 意见；页码仍单独另算）
