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
2. **渲染成 PDF**。`.docx → .pdf` 同样需要**排版/渲染引擎**。
   本工具能做的是**编排**（挑文件、命名、输出目录、批量、失败重试），渲染本身要交给一个引擎
   （Word / WPS / LibreOffice）。所以"导出 PDF"会做成**可选后端**，装了哪个用哪个。
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
docs/               规格、架构、验收策略、未决问题
reference/          VBA 宏原件（只读参考，不参与运行）
tests/              单测与差分测试夹具
tmp/                临时件（不入库）
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

## 状态

**M1 内核 + M3a 题注统一已落地**（含两版输出与体检）；段落配方（M3b）未开始。

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
- [ ] M3b：段落配方生成 / 段落重配（读写 `.xlsx`）
- [ ] M2：其余宏（格式规范化 / 去无意义空格 / 特殊字符替换）
- [ ] M4：配方编排（选定 + 排序 + 一键批量）+ 改动报告 + GUI
- [ ] M5：PDF 导出（外部渲染器编排）
