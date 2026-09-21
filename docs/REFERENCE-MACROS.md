# REFERENCE-MACROS — 12 个 Word VBA 宏的逐宏规格（XML 重写依据）

> **本文件的用途**：`word-factory` 不依赖 Word/WPS、不跑 VBA，直接读写 `.docx` 的 OOXML。
> 因此**这些宏的行为就是功能规格本身**——XML 引擎必须做出同样效果。
> **来源（只读）**：`D:\百度网盘\宏相关（主机）\word宏\` 下的 12 个 `.bas` 原件，
> 项目内只读副本在 `reference/vba-macros/`（同一份字节）。
> **编码实测**：12 个 `.bas` **全部是 cp936（GBK）**；项目内 `reference/delivered-macros/` 的 5 个 `*claw.bas` 是 **UTF-8**。
> **行号约定**：本文所有 `文件名:行号` 一律按**解码后的行**计数（cp936 解码，`\r\n` 分行），
> 与 `reference/vba-macros/` 里的副本一致，可直接用编辑器跳转核对。
> **证据等级**：正文标注为「**规格**」的是从代码里直接读出来的确定事实；
> 标为「**推演**」的是我按代码逻辑推导（未跑 Word 验证）；
> 标为「**待验证**」的是 XML 落地判断或需要实测确认的项。**凡未标注者不得当作已实测。**

---

## 0. 先看这一节：读宏的三个前提

### 0.1 这批宏的"输入"分两种，不可混用

| 输入口径 | 含义 | 本文标记 |
|---|---|---|
| **选中内容** | 只处理 `Selection.Range`；没选中就弹提示退出 | 【选中】 |
| **全文** | 处理 `ActiveDocument.Range` / `.Content` / `.Tables` / `.Paragraphs`，**与当前选区无关** | 【全文】 |
| **全故事** | `ActiveDocument.StoryRanges` 的**全部故事**（正文 + 页眉页脚 + 脚注尾注 + 文本框 + 批注…） | 【全故事】 |

> **对 XML 引擎的含义**：外置工具没有"选区"（`PLAN.md` §5.3 也点到这条）。
> 【选中】类宏必须改造成**规则驱动**（如"处理所有高亮文字"），否则无法复刻。

### 0.2 Word 对象 → OOXML 元素速查（后续每节只写差异）

| Word 对象/属性 | OOXML 落点 | 备注 |
|---|---|---|
| `Font.Color = wdColorBlack` | `w:rPr/w:color w:val="000000"` | 与 `ColorIndex` 不同，见 §2.7 |
| `Font.ColorIndex = wdBlack` | `w:rPr/w:color w:val="auto"` | 交付版用的就是这个 |
| `HighlightColorIndex = wdNoHighlight` | **删除** `w:rPr/w:highlight` | |
| `HighlightColorIndex <> wdNoHighlight` | 存在 `w:rPr/w:highlight` | 用于"识别高亮" |
| `Font.Superscript = True` | `w:rPr/w:vertAlign w:val="superscript"` | 取消 = 删该元素 |
| `Font.Subscript = True` | `w:rPr/w:vertAlign w:val="subscript"` | |
| `Font.Bold = True` | `w:rPr/w:b`（建议同时写 `w:bCs`） | |
| `Font.NameFarEast = "宋体"` | `w:rPr/w:rFonts/@w:eastAsia="宋体"` | |
| `Font.Name = "Times New Roman"` | `w:rPr/w:rFonts/@w:ascii` 与 `@w:hAnsi` | 与上面是**不同属性** |
| `ParagraphFormat.Alignment = wdAlignParagraphJustify` | `w:pPr/w:jc w:val="both"` | **不是 "justify"** |
| `ParagraphFormat.Alignment = wdAlignParagraphLeft` | `w:pPr/w:jc w:val="left"` | |
| `Paragraph.LeftIndent` | `w:pPr/w:ind/@w:left`（单位 twips = pt×20） | |
| `ParagraphFormat.FirstLineIndent`（本文未见用例） | `w:pPr/w:ind/@w:firstLine` | |
| `Table.Borders(wdBorderTop).LineWidth` | `w:tblPr/w:tblBorders/w:top/@w:sz`（1/8 pt） | 见附录 A.2 |
| `Cell.Range.Text` 里的 `Chr(7)` | **不是字符**，是 `w:tc` 边界 | 见 §2.11 坑 |
| `Chr(13)` 段落标记 | **不是字符**，是 `w:p` 边界 | |
| `Chr(11)` 手动换行 | `w:r/w:br` | |
| `TablesOfContents(1).UpdatePageNumbers` | 域 `w:fldChar` + `w:instrText " TOC ..."` | XML 层**做不到**，见 §2.7 |
| `Table.Range.Editors.Add wdEditorEveryone` | `w:permStart/@w:edGrp="everyone"` + `w:permEnd` | 见 §2.2 |

### 0.3 12 个宏里**互相冲突**的地方（必须先裁决，否则必然做错）

1. **单位上标的触发集合不一致**：
   `单位上标.bas` 用 `m[2-9]{1,}`，`格式规范化.bas` 用 `m[0-9]{1,}`。
   → `m1`（及 `m0`）**只被格式规范化处理，不被单位上标处理**。同一台机器上两个宏行为不同。
2. **智能上下标有两份互不相同的实现**（`智能上下标.bas` FSM 版 vs `智能上下标（正则）.bas` 正则版）：
   字典不同（正则版多 4 个键）、识别机制不同、边界规则只存在于 FSM 版。
   → 必须由用户指定"哪一份是规格"。本文两份都写清，差异单列（§2.5 / §2.6）。
3. **表格去空格有"原件版"与"交付版"两种语义**（原件是**全量删除**，交付版是**只裁尾部**）。
   → 见 §2.11 与附录 A.3。

---

## 1. 总表

| # | 文件名 | 入口 Sub | 一句话作用 | 输入 | 弹交互 | 依赖外部文件 |
|---|---|---|---|---|---|---|
| 1 | `MarkDown语言清除.bas` | `MarkDown语言清除` | 把选中文本里的 Markdown 标记清成普通中文文本（行内代码变中文引号） | 【选中】 | MsgBox ×2（错误时 1 个） | 无 |
| 2 | `全选表格.bas` | `全选表格` | 给全文所有表格加"所有人可编辑"权限再全选、再删权限（等价于"选中所有表格"） | 【全文】 | 无 | 无 |
| 3 | `冒号前加粗.bas` | `冒号前加粗` | 选中范围内，逐句把**首个冒号及其之前**的文字加粗 | 【选中】 | MsgBox ×2（+ 空选区提示） | 无 |
| 4 | `单位上标.bas` | `单位上标` | 全文把 `m2`…`m9` 形式的单位数字设为上标 | 【全文】 | MsgBox（计数） | 无 |
| 5 | `智能上下标.bas` | `智能上下标调整` | 选中范围内按**字典 + 状态机**智能判定 `m2`/`Vmax`/`H24` 等的上下标 | 【选中】 | MsgBox（完成） | 无 |
| 6 | `智能上下标（正则）.bas` | `智能上下标调整` | 同上，但用**正则 alternation** 识别；字典更多、无边界规则 | 【选中】 | MsgBox（完成 / 错误） | 无 |
| 7 | `格式规范化.bas` | `格式规范化` | 全文变黑 + 去高亮 + `m` 数字上标 + 更新目录页码，最后**询问保存并关闭文档** | 【全故事】+【全文】 | MsgBox ×3（含"是否保存并关闭"） | 无 |
| 8 | `段落配方生成器.bas` | `段落配方生成器` | 把选中段落变成"配方"（变量/文本序列）写到文档末尾，并可**自动生成 `.xlsx`** | 【选中】 | InputBox ×3~5 + MsgBox ×2 | **写 `.xlsx`**（Excel COM） |
| 9 | `段落重配.bas` | `段落重配` | 读选中范围的配方文本 + 对应 `.xlsx`，把变量填回去重建成段落（追加到文档末尾） | 【选中】 | MsgBox ×3 | **读 `.xlsx`**（Excel COM） |
| 10 | `目录修改.bas` | `删除低于当前选定目录层级的目录` | 按 **LeftIndent 0/28/56** 判断选中目录层级，删掉更低层级的段落 | 【选中】+【全文删除】 | MsgBox ×3 | 无 |
| 11 | `表格空格回车删除.bas` | `表格空格回车删除_全文修改` / `_选定修改` | 删除表格单元格内的空格/不间断空格/回车/手动换行（可选三档） | 【全文】与【选中】两个入口 | InputBox ×1 + MsgBox ×2 | 无 |
| 12 | `规划报告一键宏.bas` | `规划报告一键宏_重构` | 报告一键标准化：文本替换 ×2 + 两端对齐改左对齐 + 中文字体仿宋→宋体 | 【全文】 | MsgBox ×2 | 无 |

**关于第 4 条要单独说明**：用户原话提到"标准表头格式是段首空两格 + `表XX-X 表格名字` + 用空格居中"。
**这条规则在 12 个宏里没有实现**——详见 §4（这是本文最重要的一个"没找到"）。第 12 条 `规划报告一键宏.bas`
是最可能放它的地方，但它只做替换/对齐/字体三件事。

---

## 2. 逐宏规格

### 2.1 `MarkDown语言清除.bas`（92 行）

**入口**：`Sub MarkDown语言清除()` — `MarkDown语言清除.bas:2`

#### 它到底改什么【规格】
- 把**选中范围的整段文字**读成字符串，正则清理后**整段写回**：
  - 读：`MarkDown语言清除.bas:15` `originalText = selectedRange.text`
  - 写：`MarkDown语言清除.bas:20` `selectedRange.text = processedText`
- **不改任何格式属性**，只改文字。但因为是"整段文本回写"，Word 会**重建该范围的内容**。

#### 算法步骤【规格】
1. `Selection.Type = wdSelectionIP`（没选中）→ 提示后退出。（`:4-7`）
2. `On Error GoTo ErrorHandler`（`:9`），关屏更（`:10`）。
3. 取 `Selection.Range` 的文本（`:13-15`）。
4. 调 `RemoveMarkdownSymbolsWithQuotes` 做转换（`:18`）。
5. 写回选中范围（`:20`），弹"转换完成！"（`:22`）。
6. `RemoveMarkdownSymbolsWithQuotes`（`:30-92`）内部顺序：
   a. 建 `VBScript.RegExp`，`Global=True`、`IgnoreCase=False`（`:32-34`）
   b. **行内代码**：`` `X` `` → `“X”`（`:37-38`）
   c. **粗体**：`**X**` → `X`（`:41-42`）
   d. **斜体**：`*X*` → `X`（`:44-45`）
   e. `Split(text, vbCr)` 按段落切（`:48`）
   f. 每段：若 `Trim(para)` 首字符是 `#` → **循环剥掉所有前导 `#`**（`:55-62`）
   g. `LTrim` 后若首字符是 `*` / `-` / `+` → **只剥掉一个**（`:64-67`）
   h. `Replace(para, "*", "")` 删掉段内所有 `*`（`:69`）
   i. 若 `para Like "#.*"` 或 `"##.*"` → 删掉首个 `.` 之前的部分（`:71-81`）
   j. `Do While InStr(para, "  ") > 0: Replace(para, "  ", " ")` 把连续双空格折叠成单空格（`:83-85`）
   k. `Trim(para)`（`:87`）
   l. `Join(paragraphs, vbCr)` 还原（`:91`）

#### 写死的规则/常量【规格】
```vb
' MarkDown语言清除.bas:33-34
regEx.Global = True
regEx.IgnoreCase = False

' MarkDown语言清除.bas:37  →  行内代码：反引号内容换成中文左右双引号
regEx.pattern = "`([^`]+)`"
text = regEx.Replace(text, "“$1”")

' MarkDown语言清除.bas:41  →  粗体
regEx.pattern = "\*\*([^*]+)\*\*"
text = regEx.Replace(text, "$1")

' MarkDown语言清除.bas:44  →  斜体
regEx.pattern = "\*([^*]+)\*"
text = regEx.Replace(text, "$1")

' MarkDown语言清除.bas:48
paragraphs = Split(text, vbCr)

' MarkDown语言清除.bas:55-67  →  标题/列表前缀剥离
If Left(Trim(para), 1) = "#" Then
    tempPara = Trim(para)
    Do While Left(tempPara, 1) = "#"
        tempPara = Trim(Mid(tempPara, 2))
    Loop
    para = tempPara
End If
para = LTrim(para)
If Left(para, 1) = "*" Or Left(para, 1) = "-" Or Left(para, 1) = "+" Then
    para = Trim(Mid(para, 2))
End If

' MarkDown语言清除.bas:69
para = Replace(para, "*", "")

' MarkDown语言清除.bas:71-80  →  "#." / "##." 接缝处理
If para Like "#.*" Or para Like "##.*" Then
    dotPos = InStr(para, ".")
    beforeDot = Left(para, dotPos)
    afterDot = Trim(Mid(para, dotPos + 1))
    para = beforeDot & afterDot
End If

' MarkDown语言清除.bas:83-85  →  连续空格折叠
Do While InStr(para, "  ") > 0
    para = Replace(para, "  ", " ")
Loop
```

#### 边界与坑
- **`:71` 是死代码（推演）**：`:55-62` 已经把段首所有 `#` 剥掉了，等到 `:71` 时 `para` 不可能再以 `#` 开头 → 该分支永不进入。
- **写死两步的顺序错位**：作者本意大概是"`##.` 这样的编号点要保留"（`:79` 是 `beforeDot & afterDot`，**连 `.` 一起丢掉**），但因为 `#` 已被剥掉，实际不会执行。
- **`vgCr` 与 Word 段落标记**：`Selection.Range.Text` 里的段落标记是 `Chr(13)`（= `vbCr`），所以 `Split(text, vbCr)` 能正确分段；但**表格单元格结束标记是 `Chr(7)`**，`Split` 不会拆它，会把它当作普通字符留在段内 → 若选中跨表格单元格，输出会夹带 `Chr(7)`。
- **整段回写会丢 run 级格式（XML 层必须处理）**：`selectedRange.text = ...` 之后 Word 用**一个 run** 覆盖原范围 → 原选区内的加粗/上下标/字体/域会被压平。XML 层若照抄，等于把选区内所有 run 合并成一串同格式文本。**这是一条要用户确认的"是否保真"决策**。
- **`[^`]` 能跨越换行**：`[^`]` 是"非反引号"，**包含 CR**，所以 `` `甲<CR>乙` `` 这种跨段的残缺反引号会被配对替换（正则版无 `MultiLine` 限制，`[^`]` 不受影响）。
- **无 `Trim` 全文**：`:87` 只对每段 Trim，段落数不变；`Join` 用 `vbCr` 拼回，**不会多出/少掉段落**。

#### 在 OOXML 里怎么落地（待验证）
- 本质上是对**逻辑文本**做 3 次正则 + 逐段规则清洗。XML 层必须先建"逻辑文本 ↔ run 序列"映射（`PLAN.md` §2 已定），在逻辑文本上做替换，再按映射回写。
- **难点**：替换长度变化 → 必须能**拆分/合并 run**（`**粗体**` 删掉 4 个字符后，若原 run 边界落在 `**` 中间，需要重建 run）。
- **难点**：`Chr(7)`（单元格标记）在 XML 里不是字符而是 `w:tc` 边界；`Chr(13)` 是 `w:p` 边界。跨单元格/跨段落的选中在逻辑文本层要显式拼出这些"虚拟字符"。

---

### 2.2 `全选表格.bas`（12 行）

**入口**：`Sub 全选表格()` — `全选表格.bas:2`

#### 它到底改什么【规格】
这是一个**借 Word 的"编辑权限"机制来全选表格**的技巧：
- 给每个表格的整个范围加 `wdEditorEveryone` 编辑权限（`:5-7`）
- 全选所有"所有人可编辑"的区域（`:8`）
- 立刻把所有这种权限区删掉（`:9`）

```
全选表格.bas:3-9
    Dim tbl As Table
    Application.ScreenUpdating = False
    For Each tbl In ActiveDocument.Tables
        tbl.Range.Editors.Add wdEditorEveryone
    Next
    ActiveDocument.SelectAllEditableRanges wdEditorEveryone
    ActiveDocument.DeleteAllEditableRanges wdEditorEveryone
    Application.ScreenUpdating = True
```

#### 算法步骤【规格】
1. 关屏更（`:4`）
2. 遍历 `ActiveDocument.Tables`，每表的 `Range.Editors.Add wdEditorEveryone`（`:5-7`）
3. `SelectAllEditableRanges wdEditorEveryone` → 选中所有表格（`:8`）
4. `DeleteAllEditableRanges wdEditorEveryone` → 删掉刚加的权限（`:9`）
5. 开屏更（`:10`）

#### 写死的规则/常量【规格】
- 唯一常量：`wdEditorEveryone`（Word 内置枚举）。

#### 边界与坑
- **净效果存疑（未确定）**：`:9` 删掉了它自己刚加的权限，但**也会删掉文档里原本就存在的**其它 `wdEditorEveryone` 权限区 → **可能破坏用户已有的"可编辑区域"设置**。这是本宏最大的风险。
- 它是唯一一个**没有 `MsgBox`、没有错误处理**的宏：一旦 `Editors.Add` 失败（如文档受保护），宏直接以 Word 的运行时错误中断。
- 它**不返回结果**，只改变 Selection —— 用户实际用途是"手动按一次，然后接着用别的宏/手动操作处理所有表格"。

#### 在 OOXML 里怎么落地（待验证）
- Word 的编辑权限落成 `w:permStart` / `w:permEnd`（`@w:edGrp="everyone"`）。
- **待验证**：`SelectAllEditableRanges` + `DeleteAllEditableRanges` 之后，Word 是否真在 `document.xml` 留下过 `w:permStart` 痕迹（若最终被清理，则 XML 层**无需复刻**，本宏在 XML 路线下可以整体视为"选择动作"，由调用方用"所有表格"这一规则替代）。
- **建议**：本宏在 XML 引擎里**不实现为文档变换**，而实现为"作用域选择器 = 所有表格"。

---

### 2.3 `冒号前加粗.bas`（36 行）

**入口**：`Sub 冒号前加粗()` — `冒号前加粗.bas:2`

#### 它到底改什么【规格】
- 对选中范围内**每一句**，找到**第一个**冒号（中文优先），把**从句子开头到该冒号（含冒号）**的范围设 `Font.Bold = True`。
- 只动 `w:rPr/w:b`，不动文字。

```
冒号前加粗.bas:14-32
    For Each sentence In selRange.Sentences
        Dim txt As String
        txt = sentence.text
        Dim colonPos As Long
        colonPos = InStr(txt, "：")   ' 中文冒号
        If colonPos = 0 Then colonPos = InStr(txt, ":")   ' 英文冒号
        If colonPos > 0 Then
            Dim boldPart As Range
            Set boldPart = sentence.Duplicate
            boldPart.End = boldPart.Start + colonPos   ' 包括冒号
            boldPart.Font.Bold = True
        End If
    Next sentence
```

#### 算法步骤【规格】
1. 记录 `Selection.Range`（`:4`）
2. 空选区（`Start = End`）→ `MsgBox "请先选择文本"` 并退出（`:6-9`）
3. 关屏更（`:11`）
4. 遍历 `selRange.Sentences`（`:14`）：
   - 取句子文本（`:16`）
   - `InStr(txt, "：")` 找中文冒号；找不到再找英文 `":"`（`:20-21`）
   - 有冒号 → 复制句子范围（`:26`），把**结束位置**设为 `Start + colonPos`（`:27`，注释写明"包括冒号"）
   - 该范围 `Font.Bold = True`（`:30`）
5. 开屏更 + `MsgBox "冒号及前面文字已加粗!"`（`:34-35`）

#### 写死的规则/常量【规格】
- 冒号字符：**先** `"："`（U+FF1A 全角），**后** `":"`（U+003A 半角）—— 只取**第一个**命中。
- 加粗范围包含冒号本身（`:27` 的 `+colonPos` 使 `End` 指向冒号后一个字符）。
- 无其它常量。

#### 边界与坑
- **偏移量语义有隐患（推演）**：`colonPos` 是**字符串下标**（`InStr` 返回 1-based 字符序号），而 `boldPart.End = boldPart.Start + colonPos` 是 **Range 位置数**。Word 的 Range 位置与字符串下标**在含段落标记/域/表格标记时不等价**（一个 `Chr(13)` 在字符串里算 1 字符、在 Range 里也占 1 位置，通常一致；但域的代码段在 `.Text` 里不出现，Range 里却占位置）→ 句子里有域（页码/引用）时会**错位加粗**。
- **误伤**：英文冒号在中文报告里很常见（`时间：` `比例 1:500` `12:30`），任何含 `:` 的句子都会被加粗到冒号为止。
- **只处理第一个冒号**：同一句里第二个冒号之后不再加粗。
- **`sentences.Duplicate`**：句末标点被 Word 划入句子，冒号在最后一位时 `End` 会越界 1 个字符 → 把下一句首字符一并加粗（推演）。

#### 在 OOXML 里怎么落地（待验证）
- 目标属性：`w:rPr/w:b`（建议同时 `w:bCs`）。
- **难点 1：分句**。XML 里没有"句子"对象。Word 的分句规则（句号/问号/叹号/省略号，且要避开小数点、编号）需要自己实现；**建议先把这条宏的输入口径改为"每段"或"限定标记"**，否则分句歧义会带来差分测试不通过。
- **难点 2：Range 位置 ↔ 字符下标换算**。XML 层天然按字符工作，反而比 VBA 更准；但**域（`w:fldChar`/`w:instrText`）与 `w:br`/`w:tab` 要不要计入偏移**必须明确定义，否则与 Word 结果对不上。
- **难点 3：run 拆分**。要"加粗前 N 个字符"，必须把跨越第 N 个字符的 run 在第 N 个字符处切开，只给前半设 `w:b`。

---

### 2.4 `单位上标.bas`（36 行）

**入口**：`Sub 单位上标()` — `单位上标.bas:2`

#### 它到底改什么【规格】
- 用 **Word 通配符查找**在**全文**（`ActiveDocument.Range`，`:6`）里找 `m` 后跟 1 个及以上数字的模式，把**数字部分**（跳过 `m`）设 `Font.Superscript = True`。
- 只动 `w:rPr/w:vertAlign`。

```
单位上标.bas:9-31
    With rng.Find
        .ClearFormatting
        .text = "m[2-9]{1,}"  ' 匹配m后1个及以上数字
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop
        Do While .Execute
            ' 定位到数字部分（不移动结束位置！）
            rng.MoveStart wdCharacter, 1
            Dim numRng As Range
            Set numRng = rng.Duplicate
            numRng.End = rng.End
            ' 设置数字为上标
            numRng.Font.Superscript = True
            counter = counter + 1
            ' 重置范围继续搜索
            rng.Collapse wdCollapseEnd
        Loop
    End With
```

#### 算法步骤【规格】
1. `Set rng = ActiveDocument.Range`（`:6`），`counter = 0`（`:7`）
2. 设置 Find：清格式、`text = "m[2-9]{1,}"`、`MatchWildcards = True`、`Forward = True`、`Wrap = wdFindStop`（`:9-14`）
3. `Do While .Execute`（`:16`）：
   - `rng.MoveStart wdCharacter, 1` 把起点右移 1 个字符（跳过 `m`）（`:18`）
   - 复制 `numRng`（`:21-23`），`numRng.Font.Superscript = True`（`:26`）
   - `counter++`（`:27`）
   - `rng.Collapse wdCollapseEnd`（`:30`）
4. `MsgBox "单位上标转换完成！共处理 N 处"`（`:34`）

#### 写死的规则/常量【规格】
```vb
' 单位上标.bas:11   ← 只认 m2..m9（**不含 m1、m0**）
.text = "m[2-9]{1,}"  ' 匹配m后1个及以上数字

' 单位上标.bas:12-14
.MatchWildcards = True
.Forward = True
.Wrap = wdFindStop
```
- 计数上限：无。
- 交互：仅 1 个 MsgBox（`:34`）。

#### 边界与坑
- **`:18` 的 `MoveStart` 有副作用（推演）**：`Execute` 每次会把 `rng` 重置为"刚找到的匹配范围"，紧接着 `MoveStart +1` **永久改变了这个 Range 对象的起点**；下一次 `Execute` 会从被改过的起点继续查找——这正是它不会原地死循环的原因，但也意味着**匹配被"截头"后可能漏掉紧随其后的匹配**（如 `m2m3` 连写：第一次 `m2` 处理后起点落在 `2`，其后搜 `m` 时 `m3` 仍在，通常能命中；但边界情形需实测）。
- 作者注释 `' 定位到数字部分（不移动结束位置！）`（`:17`）说明这是**刻意**的写法，且是修 bug 后的结果。
- **`Wrap = wdFindStop`**：到文档末尾就停（不是绕回开头）——比 `格式规范化.bas:45` 的 `wdFindContinue` 更安全。
- **与 `格式规范化.bas` 冲突**：那边的模式是 `m[0-9]{1,}`（`:42`），**`m1` 会被它上标**。两者不能同时当规格（见 §0.3）。
- **全字匹配缺失**：`m2` 的 `m` 前面没有边界约束，`km2` 里的 `m2` 也会被命中（这一条**与 `智能上下标` 的期望行为相反**：那里 `km2` 整体规则是 `NNS`，`km2` 的 `k`+`m` 正常、`2` 上标 —— 结果其实一致；但 `cm2` 在 `智能上下标` 是 `NNS`（`c`、`m` 正常），`单位上标` 只上标 `2`，**结果同样一致**。真正不同的是 `m[2-9]` 对 `xm2`/`Rm2` 这类噪声的误伤）。

#### 在 OOXML 里怎么落地（待验证）
- 目标：给匹配到的数字字符所在 run 追加 `w:rPr/w:vertAlign w:val="superscript"`；若 `m` 与数字同在一个 run，需**在 `m` 与数字之间拆 run**。
- **待验证**：`MatchWildcards` 的 `[2-9]{1,}` 是 Word 专有通配符语法（不是标准正则），`{1,}` 在 Word 里等价"1 次及以上"。XML 层建议用 Python `re` 的 `m[2-9]+` 等价实现，但要**实测两者在 `m` 前有字母**时是否给出同样结果（Word 的 Find 是流式扫描，`km2` 会命中 `m2` 子串）。
- **难点**：域与文本框（`w:txbxContent`）里的 `m2` —— `ActiveDocument.Range` **包含**文本框内容吗？Word 的 `ActiveDocument.Range` 只覆盖主文本故事，**不含**文本框/页眉页脚（对比 `格式规范化` 用的 `StoryRanges`）。XML 层要**明确只处理 `document.xml` 正文**才是对齐的。

---

### 2.5 `智能上下标.bas`（394 行）— FSM 版

**入口**：`Sub 智能上下标调整()` — `智能上下标.bas:23`
**文件头自述**（`:2-9`）："智能上下标工具 - **修复等号边界版**"，特点三条：等号作为明确边界字符 / 精确符号识别 / **修复错误格式化问题**。

#### 它到底改什么【规格】
- 只在**选中范围**内工作（`:24-27` 没选中就退出）。
- 对识别出的符号按**规则串**逐字符设 `Superscript` / `Subscript` / 都取消。
- 规则串语义（`:366-384`）：
```vb
' 智能上下标.bas:369-383
    For i = 1 To Len(rule)
        If i <= rng.Characters.Count Then
            Select Case Mid(rule, i, 1)
                Case "S"  ' 上标
                    rng.Characters(i).Font.Superscript = True
                    rng.Characters(i).Font.Subscript = False
                Case "B"  ' 下标
                    rng.Characters(i).Font.Subscript = True
                    rng.Characters(i).Font.Superscript = False
                Case "N"  ' 正常
                    rng.Characters(i).Font.Superscript = False
                    rng.Characters(i).Font.Subscript = False
            End Select
        End If
    Next i
```
→ **`N` 是"主动清除上下标"，不是"不动"**。这条对 XML 复刻很关键（要删 `w:vertAlign`）。

#### 算法步骤【规格】
1. 选区检查（`:24-27`）
2. `charCount = selRange.ComputeStatistics(wdStatisticCharacters)`（`:35`）
3. **超过 `MAX_CHAR_LIMIT = 1230`** → 提示并退出（`:38-42`）
4. `selectedText = selRange.text`，若以 `vbCr` 结尾则去掉最后一个字符（`:46-51`）
5. 初始化两个字典（`:56-58`）
6. `识别符号(selectedText, ...)` 用状态机产出符号集合（`:62`）
7. 集合非空 → `应用格式规则`（`:65-67`）
8. `MsgBox "处理完成"`（`:70`）

**状态机 `识别符号`（`:114-228`）**：
- 状态：`STATE_START(0)` / `STATE_IN_SYMBOL(1)` / `STATE_IN_NUMBER(2)` / `STATE_IN_SPECIAL(3)` / `STATE_IN_UNIT(4)`（`:12-18`）
- 逐字符循环 `For i = 1 To Len(text)`（`:124`）：
  - **边界字符优先**（`:128-150`）：命中 → 按当前状态收尾（`IN_SYMBOL`/`IN_UNIT` 产出符号，`IN_NUMBER` 只回 START 不产出），`GoTo Continue`（边界字符本身被跳过）
  - `STATE_START`（`:153-175`）：单字符特殊符号 → 产出；否则看 2 字符特殊符号 → 产出并 `i = i + 1`；否则数字 → `IN_NUMBER`；字母 → `IN_SYMBOL`
  - `IN_NUMBER`（`:177-185`）：遇字母 → `IN_UNIT` 并记 `unitStartPos`；遇非字母数字 → 回 `START`
  - `IN_UNIT`（`:187-194`）：遇非字母数字 → 产出 `(unitStartPos, i-1)`
  - `IN_SYMBOL`（`:196-203`）：遇非字母数字 → 产出 `(startPos, i-1)`
- 循环结束后兜底（`:209-225`）：`IN_UNIT`/`IN_SYMBOL`/`IN_NUMBER` 各自产出到 `Len(text)`

**规则查询 `获取格式规则`（`:278-343`）**，三级：
1. 特殊符号走 `specialRules`，否则走 `commonRules`（精确匹配）
2. `清除格式`（只留字母数字，`:345-363`）后再查一次同一字典
3. **仅对非特殊符号**的"智能单位检测"（`:310-338`）：
   - `cleanText = "m2"` → `NS`；`"m3"` → `NS`
   - 首字符 `m` 且长度 2 且第 2 位是数字 → `NS`
   - 尾字符 `2` 且长度 > 1 且首字符是字母 → `String(Len-1,"N") & "S"`
   - 尾字符 `3` 且长度 > 1 且首字符是字母 → 同上
   - 都写死了"严格限制"（作者注释 `:322` `:328` `:334`）

#### 写死的规则/常量【规格】

**(a) 字符数上限**
```vb
' 智能上下标.bas:21
Const MAX_CHAR_LIMIT As Long = 1230 ' 固定字符数限制
' 判定在 :35-42，用 selRange.ComputeStatistics(wdStatisticCharacters)
```

**(b) 边界字符表（**含等号**，这是"修复等号边界版"的核心）**
```vb
' 智能上下标.bas:100-111
Private Function IsBoundaryChar(char As String) As Boolean
    Dim boundaryChars As String
    boundaryChars = "，。！？；：""""=-+*×/÷（）【】《》＝％"
    If InStr(boundaryChars, char) > 0 Then
        IsBoundaryChar = True
    Else
        IsBoundaryChar = False
    End If
End Function
```
**逐字符展开（实测 20 个字符，`:103` 的字节已核对）**：
`，` `。` `！` `？` `；` `：` `"`(半角) `=`(半角) `-` `+` `*` `×`(U+00D7) `/` `÷`(U+00F7) `（` `）` `【` `】` `《` `》` `＝`(全角) `％`(全角)
（注：VB 里 `""""` 表示一个半角双引号；表里**没有**空格、`~`、`°`、半角 `%`、`;`、`,`、`.`）

**(c) 普通符号字典 `commonRules`（`:76-92`）**
```vb
' 智能上下标.bas:74-97
Private Sub 初始化规则字典(commonRules As Object, specialRules As Object)
    ' 普通符号规则（字母开头），N正常、S上标、B下标
    commonRules.Add "Vmax", "NBBB"   ' Vmax
    commonRules.Add "Qpl", "NBB"    ' Qpl (pl下标)
    ' 单位符号规则（数字+单位）
    commonRules.Add "m2", "NS"      ' m2
    commonRules.Add "m3", "NS"      ' m3
    commonRules.Add "m4", "NS"      ' m?
    commonRules.Add "m5", "NS"      ' m?
    commonRules.Add "cm2", "NNS"    ' cm2
    commonRules.Add "km2", "NNS"    ' km2
    commonRules.Add "qm", "NB" ' kg/m3
    commonRules.Add "Qm", "NB" ' kg/m3
    commonRules.Add "KP", "NB" ' kg/m3
    commonRules.Add "CV", "NB" ' kg/m3
    commonRules.Add "H24P", "NBBB" ' kg/m3
    commonRules.Add "hR", "NB" ' kg/m3
    commonRules.Add "H24", "NBB" ' kg/m3
    ' 特殊符号规则（非字母开头）
    specialRules.Add "°C", "SN"    ' 摄氏度（°C）
    specialRules.Add "mm", "NN"    ' 毫米
End Sub
```
> 注意 `:82-92` 的行末注释**全是复制粘贴残留**（`:86-92` 都写 `' kg/m3`，`:82-83` 写 `' m?`），不含信息。
> 已实测核对：`' m?` 里的 `?` 是**半角 `?`（0x3F）**，疑为原文的 `m⁴`/`m⁵` 上标字符在保存链路里丢失——**只是注释，不影响功能**。

**(d) 特殊符号字典 `specialRules`（`:95-96`）** —— 仅 2 个键：
| 键 | 规则 | 逐字符含义 | 实测字节 |
|---|---|---|---|
| `°C` | `SN` | `°` 上标，`C` 正常 | `°` = U+00B0（cp936 双字节 `A1E3`） |
| `mm` | `NN` | 两个字符都"清除上下标" | 全半角 ASCII |

**(e) 辅助判定（`:387-393`）**
```vb
' 智能上下标.bas:387-393
Private Function IsLetter(char As String) As Boolean
    IsLetter = (char >= "a" And char <= "z") Or (char >= "A" And char <= "Z")
End Function
Private Function IsDigit(char As String) As Boolean
    IsDigit = (char >= "0" And char <= "9")
End Function
```
→ **只认 ASCII 字母数字**；全角数字、希腊字母、`²` 等都不算。

#### 边界与坑
- **`:155` 的单字符特殊符号分支是死代码（推演）**：`specialRules.Exists(currentChar)` 永远为假，因为两个键（`°C`、`mm`）都是 2 字符。真正生效的是 `:158-163` 的 2 字符分支。
- **`°C` 在 FSM 版里几乎永不生效（推演，需实测）**：`°` 若紧跟在数字后（如 `25°C`），此时状态已是 `IN_NUMBER`，`:155-175` 的 `STATE_START` 分支根本不会执行 → `:159-163` 的 2 字符匹配机会也没了。只有当 `°` 恰好落在**状态为 START 的位置**（例如前一个字符是非字母数字的**非边界**字符，如空格）时 `°C` 才命中。→ **这是 FSM 版相对正则版的功能缺失**。
- **`i < Len(text)` 守卫（`:158`）**：当 `i` 指向最后一个字符时，`STATE_START` 分支里的 2 字符/进入状态逻辑被整段跳过；末字符靠 `:209-225` 的兜底处理。因此**末字符若是字母，会到现在才被判为 `IN_SYMBOL` 并产出**（能工作，但路径不同）。
- **纯数字串被静默丢弃（`:145-146`）**：`IN_NUMBER` 撞上边界字符时只回 START、**不产出**。所以 `m2` 里的 `2` 不算独立符号；`2m` 这种也会因状态切换被丢。
- **`:258` 的 `On Error Resume Next` 掩盖了什么**：`应用格式规则` 在设 `symbolRange` 前开了 `On Error Resume Next`，若 `SetRange` 越界（例如 `selRange.text` 里含 `Chr(7)` 导致字符数与 Range 位置数不一致），错误被吞掉，该符号**静默漏处理**（`Err.Clear` 在 `:268`）。→ 排查困难。
- **`1230` 字符上限是性能护栏**：`:370` 用 `rng.Characters(i)` **逐字符**写格式，在 Word 里是 O(字符数) 次 COM 调用，超过约 1230 字符会明显卡顿。**XML 层完全不需要这个上限**（可如常处理全文），但要保留"结果等价"。
- **边界表含 `-`**：`m-3`、`A-1` 都会被 `-` 断开，`m-3` 的 `3` 不会被识别（作者注释说这是"修复错误格式化"，属于**有意**行为）。
- **`=` 被列为边界**（版本标题点明）：`Vmax=3.5` 里的 `=3.5` 会被断开，避免把 `Vmax=3` 之类整体误判。**复刻时必须照抄 `=` 在表里**。
- **作用域是 `selRange.text` 的字符位置**（`:261-263`）：
```vb
' 智能上下标.bas:260-263
            Set symbolRange = selRange.Duplicate
            symbolRange.SetRange _
                Start:=selRange.Start + startPos - 1, _
                End:=selRange.Start + endPos
```
→ 依赖"字符串下标 = Range 位置偏移"，**含域/表格标记时会错位**（同 §2.3）。

#### 在 OOXML 里怎么落地（待验证）
- 目标：`w:vertAlign` 三态（`superscript` / `subscript` / 删除）。
- **难点 1：位置映射**。FSM 是在"选中文本字符串"上跑偏移，XML 层应该直接对**逻辑文本**跑，然后映射回 run —— 这比 VBA 更准，但**必须与 Word 的 Range 偏移语义做差分测试**（域、`w:br`、`w:tab`、单元格标记处）。
- **难点 2：`Characters(i)` 的计数口径**。Word 的 `Range.Characters.Count` 在段末、表格、域附近与"字符串长度"有已知偏差；`应用规则到范围` 里的 `If i <= rng.Characters.Count`（`:370`）是防越界，XML 层写逻辑文本时不会遇到，但**结果要对齐**。
- **待验证**：`N`（清除上下标）在 XML 里要不要**新建** `w:rPr`？若原 run 无 `rPr`，Word 会创建；纯 XML 实现只删 `w:vertAlign` 即可（结果等价）。
- **难点 3：run 拆分**。规则串要求"符号内第 k 个字符上标"，几乎必然要把符号内的 run 切开（一个 `m2` 常被拼写检查拆成多个 `w:r`）。

---

### 2.6 `智能上下标（正则）.bas`（370 行）— 正则版

**入口**：同名 `Sub 智能上下标调整()` — `智能上下标（正则）.bas:15`
**注意**：文件第 1 行是 `Attribute VB_Name = "智能上下标"`（`：1`），与 FSM 版**同名** → **两个宏不可能同时导入同一个工程**（后导入的会覆盖/报重名）。这是"二选一"的硬证据。

#### 它到底改什么【规格】
同 §2.5（`w:vertAlign` 三态），但**识别机制完全不同**：用 `VBScript.RegExp` 的 **alternation** 一次性找出所有字典键，**没有边界字符概念**、**没有智能单位兜底**。

#### 算法步骤【规格】
1~5 同 §2.5 的 1~5（`:16-52`）
6. `正则表达式识别符号`（`:107-152`）：
   - `构建正则模式`（`:155-207`）：把**特殊符号键先加入**（`:167-169`），再把不在特殊表里的普通键加入（`:172-177`）；**按键长度降序冒泡排序**（`:186-194`）；每键经 `转义正则表达式` 后用 `|` 连接（`:199`），最后**整体包一层括号**（`:203`）
   - 设 `Global=True`、`IgnoreCase=False`、`MultiLine=True`（`:120-123`）
   - `regEx.Execute` 遍历匹配，每个匹配产出 `Array(FirstIndex+1, FirstIndex+Length, Value, 是否特殊符号)`（`:135-144`）
7. 集合非空 → `应用格式规则`（`:59-61`），**与 FSM 版是同一套逻辑**（`:235-279`）
8. `MsgBox "处理完成，共识别并格式化 N 个符号"`（`:64`）

#### 写死的规则/常量【规格】

**(a) 字符数上限**：与 FSM 版相同
```vb
' 智能上下标（正则）.bas:13
Const MAX_CHAR_LIMIT As Long = 1230 ' 固定字符数限制
```

**(b) 字典（`:72-104`）— 与 FSM 版**不同**，多 4 个键，且 `qm/Qm` 的注释已清理**
```vb
' 智能上下标（正则）.bas:74-103  ← 注意 :74-75 有 RemoveAll（FSM 版没有）
    commonRules.RemoveAll
    specialRules.RemoveAll
    specialRules.Add "°C", "SN"    ' 摄氏度（°C）
    specialRules.Add "mm", "NN"    ' 毫米
    commonRules.Add "Vmax", "NBBB"   ' Vmax
    commonRules.Add "Qpl", "NBB"    ' Qpl (pl下标)
    commonRules.Add "m2", "NS"
    commonRules.Add "m3", "NS"
    commonRules.Add "m4", "NS"
    commonRules.Add "m5", "NS"
    commonRules.Add "cm2", "NNS"
    commonRules.Add "km2", "NNS"
    commonRules.Add "qm", "NB"
    commonRules.Add "Qm", "NB"
    commonRules.Add "KP", "NB"
    commonRules.Add "CV", "NB"
    commonRules.Add "H24P", "NBBB"
    commonRules.Add "hR", "NB"
    commonRules.Add "H24", "NBB"
    commonRules.Add "Q4ml", "NBSS"    ' :100  FSM 版没有
    commonRules.Add "Q4el", "NBSS"    ' :101  FSM 版没有
    commonRules.Add "Pt", "NB"        ' :102  FSM 版没有
    commonRules.Add "Q4al+pl", "NBSSSSS" ' :103  含 + 的复合键，FSM 版没有
```

**(c) 正则元字符转义表（`:210-232`）**
```vb
' 智能上下标（正则）.bas:216-217
    Dim specialChars As String
    specialChars = "\+.?*|{}[]()^$"
```
→ **`-`、`#`、`/` 不在表内**（因为只在字符类外使用，`-` 无需转义，这里是正确的）。

**(d) 正则模式的推演结果（推演，需实测打印核对）**
按键长降序、等长保持插入顺序（`specialRules` 先、`commonRules` 后），得到：
```
(Q4al\+pl|Vmax|H24P|Q4ml|Q4el|Qpl|cm2|km2|H24|°C|mm|m2|m3|m4|m5|qm|Qm|KP|CV|hR|Pt)
```
（`:117` 有 `Debug.Print "构建的正则模式: " & pattern`，跑一次立即窗口就能核对。）

#### 边界与坑
- **完全没有边界字符概念**：`KP`、`CV`、`Pt`、`qm` 这些 2 字符短键会在**任何位置**命中 → `Pt100`、`CV值`、英文单词里的 `pt`（大小写敏感所以 `pt` 不命中，但 `Pt` 命中）都会被改格式。**这是正则版最大的误伤源**，也是 FSM 版要"修复错误格式化"的由来。
- **alternation 从左到右优先，长键优先已由排序保证，但 `mm` 排在 `m2` 前（推演）**：`mm2` 会命中 `mm`（规则 `NN` = 清除上下标），**`m2` 永远轮不到** → `mm2` 不产生上标。FSM 版同样不产生（走 2 字符特殊分支后 `2` 落进 `IN_NUMBER`）。**两版在这里结果一致**。
- **`°C` 在正则版能正常工作**（`25°C` 会在 `°` 处命中 `°C`），而 FSM 版几乎不行 → **两版行为相反**。
- **`m6`~`m9`**：正则版**没有**智能兜底（FSM 版 `:323` 有 `m+数字` 分支）→ 正则版只认识 `m2`~`m5`，`m6/m7` 不改。**两版结果不同**。
- **`:302` 的条件收紧了 `cleanText` 重试**：`If cleanText <> symbolText Then`（FSM 版没有这个守卫，无条件重试）。
- **`MultiLine = True` 但模式里没有 `^`/`$`** → 该开关**实际无影响**（推演）。
- **`On Error GoTo ErrorHandler`（`:16`）**：整宏一个错误处理，`MsgBox "发生错误：" & Err.Description`（`:68`）。比 FSM 版更"响"（FSM 版靠 `On Error Resume Next` 吞错）。

#### 在 OOXML 里怎么落地（待验证）
- **强烈建议用正则版做规格**：XML 层用 Python `re` 实现最简单、最可复现，且正则版**本来就是正则**。
- **但必须解决误伤**：直接照抄正则会引入 `Pt`/`KP`/`CV` 的无边界误匹配。建议向用户提出：**正则版 + FSM 版的边界表**（取长补短），并让用户在"误伤少"与"漏改少"之间裁决。
- 匹配偏移 → run 映射的难点同 §2.5。

---

### 2.7 `格式规范化.bas`（113 行）

**入口**：`Sub 格式规范化()` — `格式规范化.bas:2`
**这是"总入口"型宏**：4 个功能 + 2 次确认对话框，最后可能**保存并关闭文档**。

#### 它到底改什么【规格】
| 功能 | 动作 | 范围 | 行号 |
|---|---|---|---|
| 1 | `Font.Color = wdColorBlack`（全文变黑） | **全故事** | `:13-20` |
| 2 | `HighlightColorIndex = wdNoHighlight`（去高亮） | **全故事** | `:23-30` |
| 3 | `m[0-9]{1,}` 的数字部分 `Font.Superscript = True` | 正文（`ActiveDocument.Range`） | `:32-68` |
| 4 | `TablesOfContents(1).UpdatePageNumbers`（更新目录页码） | 第一个目录 | `:70-79` |
| 5 | 汇总 `MsgBox` | — | `:82-94` |
| 6 | **询问是否保存并关闭文档** | — | `:96-112` |

#### 算法步骤【规格】
1. 初始化 `counter = 0`、`bHasTOC = False`（`:9-10`）
2. **功能1**（`:13-20`）：`On Error Resume Next` → 遍历 `ActiveDocument.StoryRanges`，对每个故事用 `Do ... Set rng = rng.NextStoryRange ... Loop Until rng Is Nothing` 链式遍历，逐个 `rng.Font.Color = wdColorBlack`
3. **功能2**（`:23-30`）：同结构，`rng.HighlightColorIndex = wdNoHighlight`
4. **功能3**（`:32-68`）：
   - `Set findRng = ActiveDocument.Range`，记录 `originalStart` / `originalEnd`（`:33-37`）
   - Find：`text = "m[0-9]{1,}"`、`MatchWildcards = True`、`Forward = True`、**`Wrap = wdFindContinue`**（`:40-45`）
   - `Do While .Execute`（`:47`）：
     - `If findRng.Find.Found`（`:49`）
     - `Set rng = findRng.Duplicate`（`:51`）→ `rng.MoveStart wdCharacter, 1` 跳过 `m`（`:54`）→ `rng.Font.Superscript = True`（`:57`）
     - `counter++`（`:58`）
     - **防死循环**：`findRng.Collapse wdCollapseEnd`（`:61`）后 `findRng.End = originalEnd`（`:62`）
     - 否则 `Exit Do`（`:64`）
5. **功能4**（`:70-79`）：若 `TablesOfContents.Count > 0` → 取第 1 个 → `toc.UpdatePageNumbers` → `bHasTOC = True`
6. 汇总 `MsgBox`（`:82-94`），文案见下方常量块
7. **保存关闭确认**（`:96-112`）：`MsgBox "是否保存并关闭当前文档？"` 选"是" → 若 `ActiveDocument.Path = ""` 弹 `wdDialogFileSaveAs` 对话框，否则 `ActiveDocument.Save` → `ActiveDocument.Close SaveChanges:=wdSaveChanges` → 提示"文档已保存并关闭。"；选"否" → 提示"文档保持打开状态，修改已生效但未保存。"

#### 写死的规则/常量【规格】
```vb
' 格式规范化.bas:16   → 功能1：全文变黑（注意用 Font.Color，不是 ColorIndex）
            rng.Font.Color = wdColorBlack

' 格式规范化.bas:26   → 功能2：去高亮
            rng.HighlightColorIndex = wdNoHighlight

' 格式规范化.bas:42-45  → 功能3：正则是 [0-9]（含 m1！与 单位上标.bas 的 [2-9] 冲突）
        .text = "m[0-9]{1,}"  ' 匹配m后1个及以上数字
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindContinue  ' 改为继续查找，避免提前停止

' 格式规范化.bas:82-94   → 汇总文案（原文里的 "?" 是半角问号，疑似被替换掉的符号字符）
    msg = "格式化完成！" & vbNewLine & vbNewLine
    msg = msg & "执行结果汇总：" & vbNewLine
    msg = msg & "? 全文已设置为黑色" & vbNewLine
    msg = msg & "? 已移除所有高亮" & vbNewLine
    msg = msg & "? 单位上标处理：" & counter & " 处" & vbNewLine
    If bHasTOC Then
        msg = msg & "? 目录页码已更新"
    Else
        msg = msg & "? 未找到目录，跳过页码更新"
    End If

' 格式规范化.bas:97-106  → 保存并关闭
    If MsgBox("是否保存并关闭当前文档？", vbYesNo + vbQuestion, "保存确认") = vbYes Then
        If ActiveDocument.Path = "" Then
            Application.Dialogs(wdDialogFileSaveAs).Show
        Else
            ActiveDocument.Save
        End If
        ActiveDocument.Close SaveChanges:=wdSaveChanges
```
> **`?` 已实测核对**：`:85-91` 的 `?` 在源文件里**就是半角 `?`（0x3F）**，
> 疑为作者原文里的 `✔`/`●` 之类符号在某个环节被替换掉了。**它只影响提示文案，不影响文档**。

#### 边界与坑
- **这个宏会关掉文档（`:97-109`）**：批处理场景里极其危险（XML 引擎不要复刻"关闭"这一步）。
- **4 处 `On Error Resume Next`（`:13` `:23` `:39` `:71`）**：
  - `:13`/`:23` 吞掉"某些故事类型无内容"的错误（正常）；
  - `:39` 吞掉 Find 出错（**危险**：会静默少改）；
  - `:71` 吞掉"目录受保护/锁定"导致 `UpdatePageNumbers` 失败（**静默不更新页码**，而用户看到的是"已更新"或"未找到目录"两种文案，**无法区分这两者**）。
- **功能 1 用 `Font.Color` 而非 `ColorIndex`（`:16`）**：它把"自动色"也改成**显式黑色**。
  → XML 里等价于给每个 run 写 `w:color w:val="000000"`（**不是** `w:val="auto"`）。
  → **与交付版 `格式规范化claw.bas:7` 的 `ColorIndex = wdBlack` 不同**（那个等价 `w:val="auto"`）。
  → **这条差异会影响差分测试**（XML 里一个是新增 `w:color`，一个是新增 `w:color val="auto"`），必须选定哪个是规格。
- **功能 3 的 `wdFindContinue`（`:45`）**：查找到文档末尾会**绕回开头**继续，靠 `findRng.End = originalEnd`（`:62`）精确复位防止死循环。作者注释 `' 改为继续查找，避免提前停止`（`:45`）`' 重置查找范围，避免死循环`（`:60`）说明这两行是两次修 bug 的产物。
- **只处理第一个目录（`:73`）**：`TablesOfContents(1)` → 文档有多个目录时只更新第一个。
- **XML 路线做不到功能 4**：`UpdatePageNumbers` 需要**分页引擎**（`README.md` "明确不做"、`PLAN.md` §6 已定）。XML 引擎里功能 4 必须**显式列为"不支持"**并要求用户单独处理。

#### 在 OOXML 里怎么落地（待验证）
- 功能 1 → 遍历 `w:rPr` 插/改 `w:color w:val="000000"`；
  **待验证**：是否要处理 `w:rPrChange`（修订记录）里的旧属性、以及 `w:style` 定义里的 `w:color`（Word 的 `Font.Color` 只改**直接格式**，**不改样式**——XML 层照抄"只改 run 直接格式"才对）。
- 功能 2 → 遍历 `w:rPr` **删除** `w:highlight`（若用 `w:shd` 做高亮则不在此列，`HighlightColorIndex` 对应 `w:highlight`）。
- 功能 3 → 同 §2.4，但**模式不同（`[0-9]`）**。
- **"故事范围"要逐个部件处理**：`document.xml`（正文）+ `header*.xml` + `footer*.xml` + `footnotes.xml` + `endnotes.xml` + `comments.xml` + 文本框（`w:txbxContent` 在 document.xml 内嵌套）。
  **待验证**：`StoryRanges` 链覆盖的 11 种故事与上述部件的对应关系（`reference/delivered-macros/格式规范化claw.bas:5` 正好写成 `For st = 1 To 11`，可当覆盖清单的线索）。

---

### 2.8 `段落配方生成器.bas`（419 行）

> **这是用户最看重的功能，Excel 契约细节单列在 §3。本节只写流程与对象层。**

**入口**：`Sub 段落配方生成器()` — `段落配方生成器.bas:2`

#### 它到底改什么【规格】
1. **往当前文档末尾追加一段"配方"文本**（`ActiveDocument.Content.End`，`:90-91`）—— 这是它唯一改的文档内容；
2. 可选**新建一个 `.xlsx`**（写入表头 + 每个变量的前缀/初值）并保存到文档同目录（`:338-404`）。

#### 算法步骤【规格】
1. 空选区 → 提示退出（`:17-20`）
2. `GetUserSelection()`（`:62-79`）：`InputBox` 选模式 `1`（识别高亮）/ `2`（识别特定字符），默认 `"1"`，非法值弹错并返回 0
3. 模式 2 → `InputBox` 输入特定字符（`:26-29`），空则退出
4. `MsgBox "是否自动创建Excel表格？"`（`:32`）
5. 取 Excel 文件名（默认 `数据表.xlsx`，`:36-37` 或 `:48-49`）、Sheet 名（默认 `Sheet1`，`:39-40` 或 `:51-52`），补 `.xlsx` 后缀（`:43-45`）
6. 取配方名称（默认 `默认配方`，`:55-56`）
7. `ProcessSelectedRange`（`:81-132`）：
   - `Set recipeRng = ActiveDocument.Range` 后 `SetRange End, End`（`:90-91`）
   - 写标题 `=== 段落配方 [名] ===`（`:98`）
   - 按模式走 `ProcessHighlightedTextWithFix`（`:104-106`）或 `ProcessSpecialChars`（`:107-109`），**边扫描边拼 `recipeText` 与两个 Collection**
   - 把 `recipeText` 写到文档末尾（`:113`）
   - 依次追加 `EXCEL_FILE:`（`:117`）、`SHEET_NAME:`（`:119`）、`VARIABLE_COUNT:`（`:121`）
   - 追加 `=== 配方结束 ===` + 空行（`:124`）
   - `shouldCreateExcel` → `CreateExcelWithData`（`:127-129`）
   - 完成 `MsgBox`（带变量数，`:131`）
8. **高亮模式 `ProcessHighlightedTextWithFix`（`:134-270`）**：逐字符扫描选区，用 `doc.Range(pos, pos)` 的 `HighlightColorIndex <> wdNoHighlight` 判高亮（`:156-159`）；高亮状态**翻转**时切分：非高亮→高亮 时把前面的非高亮文本写成 `TEXT:`（`:165-174`）；高亮→非高亮 时**"隐形修正"**（`:177-213`，见坑）；末尾单独收尾（`:219-267`）
9. **特定字符模式 `ProcessSpecialChars`（`:272-336`）**：按 `Split(fullText, vbCrLf)` 分段（`:284`），段内 `Do ... InStr(currentPos, paraText, specChar)` 循环找占位符（`:298-328`），每命中一次产出一个 `VAR:`
10. **`CreateExcelWithData`（`:338-404`）**：`GetObject(, "Excel.Application")` 复用已开 Excel，失败则 `CreateObject`（`:345-350`）；`Workbooks.Add`（`:358`）→ `Worksheets(1)`（`:359`）→ `.Name = sheetName`（`:360`）→ 写表头与数据（`:363-370`）→ `AutoFit`（`:373-374`）→ `SaveAs fullPath`（`:385`）→ `Close` + `Quit`（`:395-396`）

#### 写死的规则/常量【规格】
```vb
' 段落配方生成器.bas:36-37   → 默认文件名（两处重复）
        excelPath = InputBox("请输入Excel文件名称（不含路径）：", "Excel文件", "数据表.xlsx")
        If excelPath = "" Then excelPath = "数据表.xlsx"

' 段落配方生成器.bas:39-40   → 默认 Sheet 名
        sheetName = InputBox("请输入Sheet名称：", "Sheet名称", "Sheet1")
        If sheetName = "" Then sheetName = "Sheet1"

' 段落配方生成器.bas:43-45   → 后缀归一
        If Not (Right(excelPath, 4) = ".xls" Or Right(excelPath, 5) = ".xlsx") Then
            excelPath = excelPath & ".xlsx"
        End If

' 段落配方生成器.bas:56      → 默认配方名
    If recipeName = "" Then recipeName = "默认配方"

' 段落配方生成器.bas:65-67   → 模式选择对话框
    result = InputBox("请选择识别模式：" & vbCrLf & _
                     "1 - 识别高亮文本" & vbCrLf & _
                     "2 - 识别特定字符", "模式选择", "1")

' 段落配方生成器.bas:98 / :117-124  → 配方文本骨架
    recipeRng.text = vbCrLf & "=== 段落配方 [" & recipeName & "] ===" & vbCrLf
    ...
    recipeRng.text = "EXCEL_FILE:" & excelPath & vbCrLf
    recipeRng.text = "SHEET_NAME:" & sheetName & vbCrLf
    recipeRng.text = "VARIABLE_COUNT:" & varCount & vbCrLf
    recipeRng.text = "=== 配方结束 ===" & vbCrLf & vbCrLf

' 段落配方生成器.bas:189/204/231/251  → 变量行格式（4 处一模一样）
                    recipeText = recipeText & "VAR:" & varCount & "|" & excelPath & "!" & sheetName & "!B" & (varCount + 1) & vbCrLf

' 段落配方生成器.bas:363-364  → Excel 表头
    excelWorksheet.Cells(1, 1).Value = "项目"
    excelWorksheet.Cells(1, 2).Value = "数值"

' 段落配方生成器.bas:406-418  → cleanText（TEXT: 的内容清洗）
Function cleanText(text As String) As String
    Dim result As String
    result = Trim(text)
    If Right(result, 1) = Chr(13) Then
        result = Left(result, Len(result) - 1)
    End If
    If Right(result, 1) = Chr(10) Then
        result = Left(result, Len(result) - 1)
    End If
    cleanText = result
End Function
```
- **高亮判定的阈值**：`wdNoHighlight = 0`（作者注释在 `:158`）。
- **模式 2 的变量值**：不是输入的占位符文字被替换，而是**把该占位符本身当变量值存进 Excel**：
```vb
' 段落配方生成器.bas:315
                varTexts.Add specChar
```
- **保存目录**：
```vb
' 段落配方生成器.bas:376-381
    fullPath = ActiveDocument.Path
    If fullPath = "" Then
        fullPath = Environ("USERPROFILE") & "\Documents"
    End If
    fullPath = fullPath & "\" & excelPath
```

#### 边界与坑
- **"隐形修正"（`:177-213`）是本宏最反直觉的地方【规格 + 推演】**：
  当高亮区结束于 `currentPos` 时，代码认为**最后一个高亮字符实际属于后面的正文**，于是把变量范围缩到 `correctedEnd - 1`（`:180` `:186`）：
  ```vb
  ' 段落配方生成器.bas:177-198
                ' 从高亮切换到非高亮 - 应用隐形修正
                ' 将高亮区间缩减一个字符（最后一个字符划归到后面的文本）
                Dim correctedEnd As Long
                correctedEnd = currentPos - 1
                ' 确保高亮区间至少有一个字符
                If highlightStart < correctedEnd Then
                    ' 缩减高亮区间（去掉最后一个字符）
                    Dim varRng As Range
                    Set varRng = doc.Range(highlightStart, correctedEnd - 1)
                    varCount = varCount + 1
                    recipeText = recipeText & "VAR:" & varCount & "|" & excelPath & "!" & sheetName & "!B" & (varCount + 1) & vbCrLf
                    ...
                    rng.Start = correctedEnd
                Else
                    ' 如果高亮区间只有一个字符，不进行缩减
                    Set varRng = doc.Range(highlightStart, correctedEnd)
  ```
  **推演**：这是为了绕过 Word 在"高亮区末尾"对 `HighlightColorIndex` 的读取偏差（边界字符常被读成"仍高亮"）。**变量实际内容 = `lastTextSegment` 之外的那段高亮文字**，并被写进 Excel 的 B 列。
  → **待验证**：这一步的净效果必须用"Word + 宏"跑一份金标准再与 XML 实现对比，不能靠读代码断言。
- **变量前缀的用途**：`varPrefixes.Add lastTextSegment`（`:193` 等）把"该变量**前面**那段 TEXT"存进 A 列。**`段落重配` 完全不读 A 列**（§3.5）→ A 列只是**给人看的上下文**。
- **配方不是自包含的**：`VAR:` 行里**没有**变量原来的文字，变量值只在 Excel 里 → 丢了 xlsx 就无法重配。
- **`recipeRng.text = ...` 是"插入"不是"覆盖"**：`:90-91` 把范围折叠到文档末尾，之后每次赋值都在末尾追加（这是 Word 的 Range 语义：赋值时替换范围内容；折叠到零长即插入）。XML 层要复刻"追加到 `document.xml` 末尾"。
- **`On Error Resume Next` 只在 `CreateExcelWithData` 里局部使用**（`:345-350` `:384-392`），主流程**没有**全局错误处理 → Excel 未安装时靠 `:352-355` 显式判断 `excelApp Is Nothing`。
- **`GetObject(, "Excel.Application")`（`:346`）会连上用户**正在编辑**的 Excel**：随后 `Workbooks.Add` 会在用户的 Excel 里新建工作簿、`Quit` 会**关掉用户的 Excel 进程**（含未保存工作）→ **有丢数据风险**（推演；`Quit` 在 `:396`）。
- **`:359` 只改第一个工作表的 Name（`:360`）**：若 Excel 新建工作簿的默认 sheet 数与名称冲突，`Name` 赋值可能报错（未处理）。
- **`:367` 循环用 `varPrefixes(i)` / `varTexts(i)`**（`Collection` 是 1-based）。

#### 在 OOXML 里怎么落地（待验证）
- **写 xlsx**：xlsx 也是 zip + XML → 纯标准库可写（`xl/workbook.xml`、`xl/worksheets/sheet1.xml`、`xl/sharedStrings.xml`）。
  **待验证**：`AutoFit` 列宽在 XML 里对应 `<cols>` 的 `width`/`customWidth`（Word 的 AutoFit 是"按内容计算"，XML 层要么不写宽度、要么自己算个字宽近似值——**建议第一版不写宽度**，但要向用户说明与 Word 的差异）。
- **改文档**：向 `document.xml` 末尾追加 `w:p` 段落，每行一个 `w:p`（或一个 `w:p` 里用 `w:br` 分行 —— **必须选定**，因为 `段落重配` 是按 `vbCrLf` 切分的，即**每个 `TEXT:`/`VAR:` 行都是独立段落**才对得上）。

---

### 2.9 `段落重配.bas`（281 行）

> Excel 读写契约细节见 §3。

**入口**：`Sub 段落重配()` — `段落重配.bas:2`

#### 它到底改什么【规格】
- **不改选区**：把重建出来的文本**追加到文档末尾**：
```vb
' 段落重配.bas:274-277
    Set outputRng = ActiveDocument.Range
    outputRng.SetRange ActiveDocument.Content.End, ActiveDocument.Content.End
    outputRng.text = vbCrLf & vbCrLf & "=== 重建段落 ===" & vbCrLf & outputText & vbCrLf & "=== 结束 ===" & vbCrLf
```
- 从选中范围读"配方文本"，从 `.xlsx` 读变量值，用纯文本拼出段落（**不带任何格式**）。

#### 算法步骤【规格】
1. 空选区 → 提示退出（`:13-16`）
2. `recipeText = Selection.Range.text`（`:18-19`）
3. `ParseRecipeInfo`（`:22`）解析 3 个头部字段，失败 → 提示退出（`:23-26`）
4. `ReDim excelData(1 To variableCount)`（`:29-30`）
5. `GetExcelData`（`:32`）读 Excel；失败 → 提示退出（`:33-35`）
6. `ReconstructParagraph`（`:38`）重建并写文档

#### 写死的规则/常量【规格】
```vb
' 段落重配.bas:47-50 / :196-199  → 换行统一成一个"?"再切分（两处重复）
    recipeText = Replace(recipeText, vbCrLf, "?")
    recipeText = Replace(recipeText, vbCr, "?")
    recipeText = Replace(recipeText, vbLf, "?")
    lines = Split(recipeText, "?")

' 段落重配.bas:60-64  → 头部字段识别（前缀长度写死：11 / 11 / 15）
        If Left(line, 11) = "EXCEL_FILE:" Then
            excelPath = Trim(Mid(line, 12))
        ElseIf Left(line, 11) = "SHEET_NAME:" Then
            sheetName = Trim(Mid(line, 12))
        ElseIf Left(line, 15) = "VARIABLE_COUNT:" Then
            countStr = Trim(Mid(line, 16))
            If IsNumeric(countStr) Then
                variableCount = Val(countStr)
            End If
        End If

' 段落重配.bas:75  → 三项齐备才算解析成功
    ParseRecipeInfo = (excelPath <> "" And sheetName <> "" And variableCount > 0)

' 段落重配.bas:219-224  → 段落区间标记
        If InStr(line, "=== 段落配方") > 0 Then
            inParagraphSection = True
        ElseIf InStr(line, "=== 配方结束") > 0 Then
            inParagraphSection = False
        End If

' 段落重配.bas:237-268  → TEXT:/VAR: 两种行的处理
        If Left(line, 5) = "TEXT:" Then
            textContent = Mid(line, 6)
            If textContent = "" Then
                outputText = outputText & vbCrLf     ' 空 TEXT: 表示换行
                previousLineWasText = False
            Else
                ' 前一行不是 TEXT 且当前以数字开头 → 先换行（如 "2、土壤改良工程"）
                If Not previousLineWasText And IsNumeric(Left(textContent, 1)) Then
                    outputText = outputText & vbCrLf
                End If
                outputText = outputText & textContent
                previousLineWasText = True
            End If
        ElseIf Left(line, 4) = "VAR:" Then
            currentVar = currentVar + 1
            If currentVar <= variableCount Then
                outputText = outputText & excelData(currentVar)
            Else
                outputText = outputText & "#数据缺失#"
            End If
            previousLineWasText = False
        Else
            ' 没有 TEXT:/VAR: 前缀的行 → 换行 + 原文
            If line <> "" Then
                outputText = outputText & vbCrLf & line
                previousLineWasText = True
            End If
        End If

' 段落重配.bas:116 / :122-126  → Excel 取值位置与类型转换
        cellAddress = "B" & (i + 1)
        cellValue = excelWorksheet.Range(cellAddress).Value
        If IsError(cellValue) Then
            excelData(i) = "#错误#"
        ElseIf IsNull(cellValue) Or cellValue = "" Then
            excelData(i) = ""
        Else
            excelData(i) = CStr(cellValue)
        End If

' 段落重配.bas:260  → 变量多于 Excel 行数时的占位
                outputText = outputText & "#数据缺失#"

' 段落重配.bas:279  → 完成提示
    MsgBox "段落重建完成！共使用了 " & currentVar & " 个变量。", vbInformation
```

#### 边界与坑
- **分隔符就是一个半角 `?`（**已实测核对字节**）**：`:47-50` 与 `:196-199` 的 `"?"` 在源文件里是 `0x3F`。
  → **推演**：作者原意大概是一个不会出现在正文里的特殊分隔符（如 `Chr(1)`），在某个复制/保存环节被替换成了 `?`。
  → **后果**：配方正文里只要出现一个**半角问号**，切分就会错位。**XML 引擎要重做这条契约**（建议改用 `\n` 或显式 `\x1F`），并且**必须与用户确认是否要兼容这个 `?` 行为**。
- **`ParseRecipeInfo` 只认 3 个头部行，且行首匹配**：`Trim` 之后 `Left(...)` 比较 → 前面有空格也 OK（`:57`）。
- **`VAR:` 行里的负载字段（`|路径!Sheet!B2`）**：**没有任何代码读取**。`ParseRecipeInfo` 只看 3 个 `:` 头，`ReconstructParagraph` 只数 `VAR:` 出现次数（`:255` `Left(line,4)`）→ **`段落重配` 完全依赖头部的 `EXCEL_FILE:`/`SHEET_NAME:`**。这是一处"生成端写得详细、消费端不用"的冗余（关系到 §3 的第三方读写契约）。
- **`variableCount` 必须 > 0**（`:75`）：纯文本配方（无变量）无法重配。
- **`excelData` 只有 `variableCount` 个**，靠 `currentVar <= variableCount` 防越界（`:257`）。
- **数字经 `CStr` 会本地化**（`:126`）：`1.5` → `"1.5"`；大数可能变成 `"1.23457E+11"`，千分位与科学计数法都会**改变原文**。→ XML 侧建议**读单元格原始值**而不是格式化文本。
- **`GetFullExcelPath` 的三级查找（`:153-184`）**：① 含 `\` 且 `Dir()` 存在 → 原样；② `ActiveDocument.Path & "\" & fileName`；③ 当前工作目录 `Dir(fileName)`；都失败 → `""`。
- **`GetExcelData` 的错误处理（`:141-151`）**：`ErrorHandler` 里关工作簿、退出 Excel，返回 `False`；但**没有释放对象引用到 `Nothing`**（`:134-136` 只在成功路径）→ 残留 COM 引用（推演）。
- **写出来的是纯文本段落**：不继承任何样式/编号/缩进（**这是用户可能最不满意的地方**，见 §3.7）。

#### 在 OOXML 里怎么落地（待验证）
- **读 xlsx**：直接解 `xl/worksheets/sheetN.xml` + `xl/sharedStrings.xml`，取 **B2..B(n+1)**。定义"B 列"用的是 `Range("B2")` 这种 **A1 风格地址** → XML 层要把行号/列号映射成 `<c r="B2"><v>…</v></c>`（**注意**：若单元格是共享字符串，`<v>` 是索引、文本在 `sharedStrings.xml`；也可能被 `<is><t>` 内联存储）。
- **写文档**：追加 `w:p`；`outputText` 里的 `vbCrLf` 要变成**新段落**（`w:p`）或 `w:br` —— **必须与生成端一致**（生成端每行是一个 `w:p`，因为它是 `recipeRng.text = "... & vbCrLf"` 插入的）。

---

### 2.10 `目录修改.bas`（43 行）

**入口**：`Sub 删除低于当前选定目录层级的目录()` — `目录修改.bas:2`
（**注意**：Sub 名与文件名不同，这也是交付版把它改名成 `静默更新目录页码_claw` 的原因，见附录 A.1）

#### 它到底改什么【规格】
- 读选中段落的 `LeftIndent` 判断"选中了第几级目录"，然后**删除**全文里缩进达到更低层级的**段落**。
- 动作是 `para.Range.Delete`（`:37`）—— **真删文字**。

```
目录修改.bas:9-39
    If Selection.Type = wdNoSelection Then
        MsgBox "请先选中一个目录项！"
        Exit Sub
    End If
    indentValue = Selection.paragraphs(1).LeftIndent
    Select Case indentValue
        Case 0
            selectedLevel = 1 ' 一级目录
        Case 28
            selectedLevel = 2 ' 二级目录
        Case 56
            selectedLevel = 3 ' 三级目录
        Case Else
            MsgBox "未识别的目录层级"
            Exit Sub
    End Select
    For Each para In ActiveDocument.paragraphs
        indentValue = para.LeftIndent
        If (selectedLevel = 1 And indentValue >= 28) Or _
           (selectedLevel = 2 And indentValue >= 56) Then
            para.Range.Delete
        End If
    Next para
```

#### 算法步骤【规格】
1. `Selection.Type = wdNoSelection` → 提示退出（`:9-12`）
2. `indentValue = Selection.paragraphs(1).LeftIndent`（`:15`）
3. `Select Case`：`0`→1 级、`28`→2 级、`56`→3 级、其它→提示退出（`:18-28`）
4. 遍历 `ActiveDocument.paragraphs`，`indentValue >= 28`（选 1 级时）或 `>= 56`（选 2 级时）→ `para.Range.Delete`（`:31-39`）

#### 写死的规则/常量【规格】
```vb
' 目录修改.bas:18-28   → 层级 ⇄ LeftIndent 的映射（**单位是 pt**，见下）
    Select Case indentValue
        Case 0
            selectedLevel = 1 ' 一级目录
        Case 28
            selectedLevel = 2 ' 二级目录
        Case 56
            selectedLevel = 3 ' 三级目录

' 目录修改.bas:35-36   → 删除条件（**只有 1 级、2 级有分支**）
        If (selectedLevel = 1 And indentValue >= 28) Or _
           (selectedLevel = 2 And indentValue >= 56) Then
```

#### 边界与坑
- **选中 3 级目录 → 静默什么都不做（推演，实为 bug）**：`Case 56` 把 `selectedLevel` 设为 `3`，但 `:35-36` 的两个条件**都不涉及 `selectedLevel = 3`** → 循环体永不执行，宏跑完没有任何提示。**必须作为规格缺陷报给用户**。
- **`LeftIndent` 用 `Single` 精确等于 28/56 判定**：不同模板/手工拖拽会产生 `27.9`、`28.5` 等值 → 落到 `Case Else` 报"未识别的目录层级"。**建议改成区间判定**（推演）。
- **只取选中段的第一个段落**（`Selection.paragraphs(1)`，`:15`）：选中多段时只看第一段。
- **`para.Range.Delete` 在 `For Each` 中改集合**：Word 允许，但可能跳段落（推演）。
- **误删风险**：条件只看缩进，**不看样式**→ 正文里任何 `LeftIndent >= 28pt` 的段落（如引文块、缩进的列表）都会被删。**建议加 `para.Style` 判定**。
- **这是唯一给出"文档版式常数"的宏**：28 pt / 56 pt 的步长。对 XML 层是一个**有用的旁证**（说明该文档体系里目录每级缩进 28pt = `w:ind/@w:left="560"` twips），但**它不是** §4 那个"表头空格居中"的公式。

#### 在 OOXML 里怎么落地（待验证）
- 目标：读 `w:pPr/w:ind/@w:left`（twips ÷ 20 = pt），删除整段（删 `w:p`）。
- **待验证**：Word 的 `Paragraph.LeftIndent` 读数来自"样式 + 直接格式 + 文档默认"的**合成值**；XML 层直接读 `w:ind` 会**漏掉样式里定义的缩进** → 必须实现样式继承解析（`w:pStyle` → `styles.xml` → `w:basedOn` 链 → `docDefaults`）。
- 目录段落若在 `w:sdt`（内容控件）里，删除 `w:p` 还要考虑控件边界。

---

### 2.11 `表格空格回车删除.bas`（202 行）

**入口（两个）**：
- `表格空格回车删除_全文修改()` — `表格空格回车删除.bas:9`
- `表格空格回车删除_选定修改()` — `表格空格回车删除.bas:21`

#### 它到底改什么【规格】
- 对每个目标表格的每个单元格：把单元格内容里的**空格 / 不间断空格 / 回车 / 手动换行 / 表格标记**按用户选的档位**全部删除**，然后 `cellRange.text = newText` 写回。
- **是"全量删除"，不是"裁首尾"**（与交付版相反，见附录 A.3）。

#### 算法步骤【规格】
1. `GetUserRemoveOption()`（`:175-201`）：`InputBox` 选 `1` 仅空格 / `2` 仅回车 / `3` 两者，**默认值 `"3"`**；空串→`0`（取消）；非数字→弹错→`0`；越界(非 1~3)→弹错→`0`
2. `ProcessTablesInDocument removeOption, processSelection`（`:35-121`）：
   - **步骤 1 收集表格**（`:56-68`）：选定模式下用区间相交判定
     ```vb
     ' 表格空格回车删除.bas:59
             If Not (Selection.Start > tbl.Range.End Or Selection.End < tbl.Range.Start) Then
     ```
     全文模式全部加入。无表格 → 提示后 `GoTo ExitSub`（`:71-78`）
   - **步骤 2 逐单元格**（`:81-101`）：选定模式再判单元格与选区是否相交（`:86-87`），调用 `ProcessSingleCell`
   - **步骤 3 汇总 `MsgBox`**：`"扫描了 N 个单元格，其中 M 个单元格的内容被清理。"`（`:104-112`）
3. `ProcessSingleCell`（`:125-172`）：
   - `Set cellRange = targetCell.Range` → **`cellRange.MoveEnd wdCharacter, -1`**（`:136`，排除单元格结束标记）
   - `If Len(originalText) > 0`（`:142`）才处理
   - 按 `removeOption` 做 `Select Case` 替换（`:144-158`）
   - 有变化才写回（`:161-164`）：`If newText <> originalText Then cellRange.text = newText`
   - 出错 → `ErrorHandler` 返回 `False`（`:170-171`，**静默跳过该单元格**）

#### 写死的规则/常量【规格】
```vb
' 表格空格回车删除.bas:144-158   → 三档删除的字符集（**权威表**）
        Select Case removeOption
            Case 1  ' 仅删除空格
                newText = Replace(newText, " ", "")
                newText = Replace(newText, Chr(160), "") ' 不间断空格
            Case 2  ' 仅删除回车
                newText = Replace(newText, Chr(13), "")   ' 段落标记
                newText = Replace(newText, Chr(11), "")   ' 手动换行符
                newText = Replace(newText, Chr(7), "")    ' 表格标记
            Case 3  ' 删除空格和回车
                newText = Replace(newText, " ", "")
                newText = Replace(newText, Chr(160), "")
                newText = Replace(newText, Chr(13), "")
                newText = Replace(newText, Chr(11), "")
                newText = Replace(newText, Chr(7), "")
        End Select

' 表格空格回车删除.bas:179-182   → 交互文案与默认值 "3"
    userInput = InputBox("请选择要删除的内容：" & vbCrLf & _
                       "1 - 仅删除空格" & vbCrLf & _
                       "2 - 仅删除回车" & vbCrLf & _
                       "3 - 删除空格和回车", "去除空格和回车", "3")
```
**字符码对照（Word/VBA 语义）**：

| 代码 | Unicode | 在 Word 里是什么 | 在 OOXML 里对应什么 |
|---|---|---|---|
| `" "` | U+0020 | 半角空格 | `w:t` 里的普通空格（**注意 `xml:space="preserve"`**） |
| `Chr(160)` | U+00A0 | 不间断空格 NBSP | `w:t` 里的 U+00A0（也可写成实体） |
| `Chr(13)` | U+000D | 段落标记 | **`w:p` 边界**（不是字符） |
| `Chr(11)` | U+000B | 手动换行 | `w:r/w:br`（若无 `@w:type`） |
| `Chr(7)` | U+0007 | 单元格结束标记 | **`w:tc` 边界**（不是字符） |

#### 边界与坑
- **不含全角空格（U+3000）**：中文报告里最常见的是全角空格 `　`，**本宏删不掉**。→ 待用户确认是否补上（推演：很可能用户以为它删了）。
- **`Chr(7)` 的 `Replace` 在写回时会发生什么**：`:136` 已经用 `MoveEnd -1` 排除了单元格末标记，正常情况 `Chr(7)` 不会出现；但若单元格内有嵌套表格，`Cells` 的 `.Range.Text` 可能带 `Chr(7)`，此时 `Replace` 会**删掉嵌套结构标记** → 写回时 Word 自行修复/可能报错。`On Error` 会把它变成**静默跳过**（`:170-171`）。
- **整格文本写回 = 丢弃 run 级格式**（`:162`）：单元格里的字体/加粗/上下标/颜色全被压成一种。**这是与 XML 实现最大的语义分歧点**——XML 层若"只删字符、保留 run 属性"，结果会比 Word 宏**更保真**，差分测试会不通过。**必须让用户裁决保真度**。
- **`On Error GoTo ErrorHandler`（`:132`）+ 返回 False**：任何单元格出错都无声跳过，只在最后汇总里体现为"M 小于 N"。用户**无法知道哪些单元格失败了**。
- **两个入口共用核心**，唯一区别是 `processSelection` 布尔 → **XML 引擎天然可以只做"全文"一档**，把"选定"换成规则（如"只处理某个表格编号范围"）。
- **`Selection.Start/End` 与表格范围相交**（`:59`）：这是 Word 里判"选区与表格相交"的标准做法；XML 层没有选区，要改成"按索引/编号筛选表格"。

#### 在 OOXML 里怎么落地（待验证）
- 目标：清理 `w:tc` 内所有 `w:t` 的文本。
- **难点 1：`Chr(13)` 的删除 = 合并段落**。一个单元格里有 2 个 `w:p`，要"删回车"就必须把两个段落**合并成一个 `w:p`**（后段的 run 追加到前段）——这是**结构性**改动，且要处理前段末尾的 `w:rPr/w:rPr`（段落标记格式）、`w:pPr`（缩进/对齐）取舍。**这是本宏在 XML 层最难的一步**。
- **难点 2：`Chr(11)` = `w:br`**，直接删 `w:br` 即可（相对简单）。
- **难点 3：`Chr(7)` 没有字符对应**，XML 层**什么都不用做**（结构本身就是 `w:tc`）→ 需要向用户说明这条"自动等价"。
- **难点 4：空格删了要处理 `xml:space="preserve"`**：`w:t` 首尾空格靠该属性保留，删掉中间空格后若首尾出现空格，要确保属性正确（反之可能出现空格被"吃掉"的假象）。
- **难点 5：合并单元格**：`w:tcPr/w:gridSpan`（横合并）与 `w:vMerge`（纵合并）→ 遍历时要按 `w:tblGrid` 展开列，不能按 `w:tr/w:tc` 计数。

---

### 2.12 `规划报告一键宏.bas`（105 行）

**入口**：`Sub 规划报告一键宏_重构()` — `规划报告一键宏.bas:2`
**文件头自述**（`:3-6`）："重构版：执行文档格式标准化 / 功能：1.替换文本 2.设置段落格式 3.设置字体格式"

#### 它到底改什么【规格】
四个动作，全部作用在 **`ActiveDocument.Content`（全文）**：

| # | 动作 | Word 属性 | 行号 |
|---|---|---|---|
| 1 | `其它` → `其他` | Find/Replace 文本 | `:17-27` |
| 2 | `东流流经` → `向东流经` | Find/Replace 文本（**新增功能**） | `:30-40` |
| 3 | 段落对齐 两端对齐 → 左对齐 | `ParagraphFormat.Alignment` | `:44-55` |
| 4 | 中文字体 `仿宋` → `宋体` | `Font.NameFarEast` | `:67-78` |

#### 算法步骤【规格】
1. `On Error GoTo ErrorHandler`（`:8`）
2. **保存原选区**：`Set bkm = Selection.Range`（`:10-11`，注释"保存原始位置，处理完恢复"）
3. 关屏更（`:14`）
4. **替换 1**（`:17-27`）：`Find` 用 `其它`→`其他`，`Forward=True`、`Wrap=wdFindContinue`、`MatchCase=False`、`MatchWholeWord=False`、`Execute Replace:=wdReplaceAll`
5. **替换 2**（`:30-40`）：同上，`东流流经`→`向东流经`
6. **对齐**（`:44-55`）：`Find.text = ""` + `.ParagraphFormat.alignment = wdAlignParagraphJustify`（找"两端对齐"的段落），`Replacement.text = ""` + `.Replacement.ParagraphFormat.alignment = wdAlignParagraphLeft`，**`.Format = True`**（关键），`Execute Replace:=wdReplaceAll`
7. （`:57-64` 是注释掉的"遍历段落改对齐"备选方案，**未执行**）
8. **字体**（`:67-78`）：`Find.text = ""` + `.Font.NameFarEast = "仿宋"` → `.Replacement.Font.NameFarEast = "宋体"`，`.Format = True`，`Execute Replace:=wdReplaceAll`
9. （`:80-83` 是注释掉的"英文字体"代码，`ActiveDocument.Content.Font.Name = "Times New Roman"`，**未执行**）
10. `bkm.Select` 恢复选区（`:86`），开屏更（`:87`）
11. 完成 `MsgBox`（`:89-94`）

#### 写死的规则/常量【规格】
```vb
' 规划报告一键宏.bas:20-21    → 替换 1
        .text = "其它"
        .Replacement.text = "其他"

' 规划报告一键宏.bas:33-34    → 替换 2（"新增功能"）
        .text = "东流流经"
        .Replacement.text = "向东流经"

' 规划报告一键宏.bas:47-51    → 对齐：两端 → 左
        .text = ""  ' 空文本匹配任意内容
        .ParagraphFormat.alignment = wdAlignParagraphJustify
        .Replacement.text = ""  ' 不改变文本
        .Replacement.ParagraphFormat.alignment = wdAlignParagraphLeft
        .Format = True  ' 重要：应用格式查找

' 规划报告一键宏.bas:70-74    → 字体：仿宋 → 宋体（**中文字体，NameFarEast**）
        .text = ""  ' 空文本匹配任意内容
        .Font.NameFarEast = "仿宋"  ' 查找中文字体为仿宋
        .Replacement.text = ""  ' 不改变文本
        .Replacement.Font.NameFarEast = "宋体"  ' 改为宋体
        .Format = True  ' 重要：应用格式查找

' 规划报告一键宏.bas:89-94    → 完成文案（**权威的"这个宏做了什么"清单**）
    MsgBox "文档格式标准化完成！" & vbCrLf & _
           "1. '其它'已替换为'其他'" & vbCrLf & _
           "2. '东流流经'已替换为'向东流经'" & vbCrLf & _
           "3. 段落对齐：两端对齐→左对齐" & vbCrLf & _
           "4. 字体：仿宋→宋体", _
           vbInformation, "操作完成"
```
> **字体名注意**：原件用 `仿宋`；`reference/delivered-macros/规划报告一键宏claw.bas:11` 用的是 `仿宋_GB2312`。
> 两者是**不同字体名**（`仿宋_GB2312` 是旧字库名）→ 必须让用户确认模板里到底是哪个。

#### 边界与坑
- **`.text = ""` + `.Format = True` 的"空文本格式替换"是 Word 技巧**：等价于"把所有**满足该格式**的 run 改属性"。在 XML 层等价于遍历 run 改 `w:jc` / `w:rFonts`。
  - **但有一个关键差别**：段落对齐是**段落属性**（`w:pPr/w:jc`），Word 的 Find 是按"段落格式"匹配的；XML 层直接遍历 `w:p` 更准确。
  - **格式查找会漏掉"继承自样式"的内容吗？** 不会——Word 的 Find 匹配的是**显示出来的有效格式**（含样式继承）。→ **XML 层必须实现样式继承解析**（否则漏改）。
- **替换"其它"→"其他"是跨 run 的**：Word 的 Find **能跨 run 匹配**（如 `其` 和 `它` 在两个 run 里）。XML 层若按 run 逐个替换会**漏掉** → 必须在逻辑文本层做替换再拆回 run。**这是本项目文本层的核心难点**（`PLAN.md` §2 已点到）。
- **`bkm.Select`（`:86`）在 `ErrorHandler` 里也调用（`:103`）**：若 `bkm` 未成功赋值（`Set bkm = Selection.Range` 在 `:11` 且 `On Error` 在 `:8` 之后，正常应有值），错误路径里的 `bkm.Select` 可能二次报错。
- **无 `ActiveDocument.Save`**：本宏不保存（与 `格式规范化.bas` 相反）。
- **`MatchWholeWord = False`**（`:25` `:38`）→ `其它` 是子串匹配（中文无词边界，实际等价）。
- **注意它**不**处理表头**：`PLAN.md` §5.4 猜它实现了"表头格式"，**实际没有**（见 §4）。

#### 在 OOXML 里怎么落地（待验证）
- 动作 1/2 → 逻辑文本层替换（跨 run），再按新文本重建 run 边界。
- 动作 3 → 遍历 `w:p`，把 `w:pPr/w:jc` 为 `both` 的改成 `left`（**待验证**：`w:jc` 缺失表示默认左对齐，所以只需处理显式 `both`；但"样式里是 both、直接格式没写"的段落要**在 `w:pPr` 里显式写 `left`** 才能盖住样式）。
- 动作 4 → 遍历 `w:r/w:rPr/w:rFonts`，`@w:eastAsia="仿宋"` 改成 `"宋体"`；**样式层的 `w:rFonts` 要不要一起改？** Word 的 Find 不改样式定义（只改直接格式）→ **XML 层照抄只改 run**，但要向用户说明"样式定义的仿宋不动"。

---

## 3. 「段落配方」这套：Excel 契约（**单列细写**）

> 涉及两个宏：`段落配方生成器.bas`（写配方 + 写 xlsx）与 `段落重配.bas`（读配方 + 读 xlsx）。
> 本节是**第三方照着重写的契约**，每一条都带行号。

### 3.1 文件与位置

| 项 | 规则 | 出处 |
|---|---|---|
| 文件名 | **用户输入**；`InputBox` 默认填 `数据表.xlsx`；留空则用默认 | `段落配方生成器.bas:36-37`、`:48-49` |
| 后缀归一 | 若结尾不是 `.xls`/`.xlsx` → 追加 `.xlsx` | `段落配方生成器.bas:43-45` |
| 存放目录（**生成时**） | `ActiveDocument.Path`；若文档未保存（`Path = ""`）→ `Environ("USERPROFILE") & "\Documents"`；再拼 `"\" & excelPath` | `段落配方生成器.bas:376-381` |
| Sheet 名 | **用户输入**；默认 `Sheet1` | `段落配方生成器.bas:39-40`、`:51-52` |
| 工作簿来源（生成时） | `Workbooks.Add`（**新建**，不是打开已有文件）→ 因此**每次生成都会覆盖同名文件**（`SaveAs` 到已有路径会触发 Excel 覆盖询问，`DisplayAlerts` 未设为 False，**可能弹窗卡住批处理**） | `段落配方生成器.bas:358`、`:385` |
| Excel 复用 | `GetObject(, "Excel.Application")` 优先复用已开的 Excel，失败才新建 | `段落配方生成器.bas:346-349` |

**`GetFullExcelPath`（读取时的定位规则）** — `段落重配.bas:153-184`，按顺序：

1. 若 `fileName` 含 `\` **且** `Dir(fileName) <> ""` → 直接当完整路径返回（`:158-163`）
2. `ActiveDocument.Path`（末尾补 `\`）+ `fileName`，`Dir()` 存在则返回（`:166-174`）
3. `Dir(fileName)`（**当前工作目录**，即 Word 进程的 CurDir）存在则返回（`:177-180`）
4. 都不行 → 返回 `""` → `GetExcelData` 返回 `False` → 主流程弹"无法从Excel获取数据！请检查文件路径和Sheet名称。"并退出（`段落重配.bas:89-92`、`:32-35`）

> **对 XML 引擎的含义**：生成端写"文档同目录"，读取端找"文档同目录 / 当前目录"。
> 外置工具应当**统一为"输入 `.docx` 的同目录"**，并把这层查找规则写进 CLI 参数（`--data-dir`）。

### 3.2 工作表结构（**权威**：写端的 6 行代码就是全部契约）

```vb
' 段落配方生成器.bas:359-374
    Set excelWorksheet = excelWorkbook.Worksheets(1)
    excelWorksheet.Name = sheetName
    ' 写入表头
    excelWorksheet.Cells(1, 1).Value = "项目"
    excelWorksheet.Cells(1, 2).Value = "数值"
    ' 写入变量前缀和初始值
    For i = 1 To varPrefixes.Count
        excelWorksheet.Cells(i + 1, 1).Value = varPrefixes(i)
        excelWorksheet.Cells(i + 1, 2).Value = varTexts(i)
    Next i
    ' 自动调整列宽
    excelWorksheet.Columns(1).AutoFit
    excelWorksheet.Columns(2).AutoFit
```

| 位置 | 内容 | 语义 |
|---|---|---|
| 工作表 | 工作簿**第 1 个**表，名字 = 用户输入的 Sheet 名 | 只有一个表 |
| `A1` | `项目` | 表头（固定字面量） |
| `B1` | `数值` | 表头（固定字面量） |
| **`A(n+1)`** | 第 n 个变量的**前缀**：该变量**前面紧邻的那段 `TEXT:` 文本** | **给人看的上下文，`段落重配` 完全不读** |
| **`B(n+1)`** | 第 n 个变量的**值** | **`段落重配` 唯一读取的列** |
| 列宽 | A、B 两列 `AutoFit` | 只影响观感 |

- **变量编号 n 从 1 开始，固定落在 Excel 第 n+1 行、B 列**。
  生成端把行号写进了配方：`"!B" & (varCount + 1)`（`:189` `:204` `:231` `:251`）；
  读取端也是同一个公式：`cellAddress = "B" & (i + 1)`（`段落重配.bas:116`）→ **两边独立算出来的，必须一致**。

### 3.3 三类识别模式分别往表里放什么

| 模式 | 怎么切变量 | `varTexts.Add(...)`（→ **B 列**） | `varPrefixes.Add(...)`（→ **A 列**） |
|---|---|---|---|
| **1 高亮** | 逐字符扫选中范围，`HighlightColorIndex <> wdNoHighlight` 的连续片段 | 高亮片段的原文（**经"隐形修正"少一个字符**，见 §2.8 坑） | 该高亮片段**前面**那段非高亮文本（`lastTextSegment`） |
| **2 特定字符** | `InputBox` 输入的字符串（如 `xx`）每次出现算一个变量 | **就是这个特定字符本身**（`段落配方生成器.bas:315` `varTexts.Add specChar`），**不是**该位置原有的文字 | 该占位符之前的文本片段 |

出处：高亮 → `段落配方生成器.bas:193-194`、`:208-209`、`:235-236`、`:255-256`；
特定字符 → `段落配方生成器.bas:314-315`。

**用户问的"选中范围/高亮文本/特殊字符三类各是什么"，落点是：**
- **选中范围** = 生成配方时**必须先选中**（`:17-20`），选区是唯一的"处理范围"来源（外置工具无选区 → 需换成规则）；
- **高亮文本** = 模式 1 里被判定为高亮的片段 → **成为变量**（B 列的值）；
- **特殊字符** = 模式 2 里的占位符（一个用户输入的字符串）→ **每次出现成为一个变量**（B 列存的就是它本身，A 列存它前面的文本）。

### 3.4 单元格里存什么格式的文本

- **`B` 列存的是"纯文本"**：`Cells(r, c).Value` 直接赋值字符串，**不带任何单元格格式**（没有字体、颜色、批注、数据验证）。
- **换行**：`cleanText`（`段落配方生成器.bas:406-418`）会把 `TEXT:` 内容的**首尾 `Chr(13)`/`Chr(10)` 去掉** → 写进 Excel 的文本**不含换行**。多段落被拆成**多行**（每个 `TEXT:` 一个变量行 → 各占一行）。
- **空段落的表示**：`TEXT:`（后面什么都不写）表示一个换行；在`ProcessSpecialChars` 里空段落写成 `TEXT:`（`:292-295`），段间还会补一个 `TEXT:`（`:332-334`）。
- **占位符/ID**：**没有独立的"配方编号/ID"字段**。唯一的编号就是 `VAR:n` 里的序号，与 Excel 行号 `n+1` 绑定。配方名只出现在文档标题行 `=== 段落配方 [名] ===`（`:98`），**不写进 Excel**。
- **数字/日期**：若 `varTexts` 存的是 `"1.5"` 这类字符串，Excel 会**按数字存**（`Value` 赋值会做类型推断）→ 读回来时 `CStr(cellValue)` 可能改变写法（千分位、科学计数法）。

### 3.5 配方文本格式（**实例**）

**生成端写入顺序（`:98` → `:113` → `:117-124`）**，即文档末尾出现：

```
（空行）
=== 段落配方 [土方计算] ===
TEXT:1、
VAR:1|数据表.xlsx!Sheet1!B2
TEXT:本期
VAR:2|数据表.xlsx!Sheet1!B3
TEXT:万m
VAR:3|数据表.xlsx!Sheet1!B4
TEXT:
TEXT:2、
VAR:4|数据表.xlsx!Sheet1!B5
TEXT:设计工程量
EXCEL_FILE:数据表.xlsx
SHEET_NAME:Sheet1
VARIABLE_COUNT:4
=== 配方结束 ===
（空行）
```

**解析规则（`段落重配.bas:41-76`）**：
- 换行统一成 `?` 后按 `?` 切行（`:47-50`），逐行 `Trim`（`:57`），空行跳过（`:58`）
- `Left(line,11) = "EXCEL_FILE:"` → 取 `Mid(line,12)` 作路径（`:60-61`）
- `Left(line,11) = "SHEET_NAME:"` → 取 `Mid(line,12)` 作表名（`:62-63`）
- `Left(line,15) = "VARIABLE_COUNT:"` → 取 `Mid(line,16)`，`IsNumeric` 才 `Val`（`:64-69`）
- 三项都非空且变量数 > 0 → 成功（`:75`）

**`ReconstructParagraph` 的还原规则（`段落重配.bas:186-280`）**：
1. 只有落在 `=== 段落配方` 与 `=== 配方结束` **之间**的行才处理（`:219-229`）；区间外跳过
2. 三个头部行**显式跳过**（`:232-234`）
3. `TEXT:` 开头（`:237`）：
   - 内容为空 → **输出一个换行**（`:242-245`）
   - 内容非空 → 若**前一行不是 TEXT** 且**本行首字符是数字** → **先补一个换行**（`:248-250`，为 `2、土壤改良工程` 这种编号段准备）
   - 追加内容，标记"上一行是 TEXT"（`:252-253`）
4. `VAR:` 开头（`:255`）：**只数序号**，`currentVar++`，取 `excelData(currentVar)` 追加；超出范围 → 追加 `#数据缺失#`（`:256-262`）
5. **其它行**（`:263-268`）：`outputText & vbCrLf & line`（换行 + 原文），并标记"上一行是 TEXT"
6. 最终写文档：`vbCrLf & vbCrLf & "=== 重建段落 ===" & vbCrLf & outputText & vbCrLf & "=== 结束 ===" & vbCrLf` 追加到**文档末尾**（`:277`）

> **样式/编号/缩进从哪来？答案：不来自任何地方。**
> `ReconstructParagraph` 只拼**纯文本**，写回时是 `outputRng.text = ...`（`:277`）→ 新段落用**当前插入点/默认样式**。
> **`TEXT:` 里原本的样式信息（标题级别、编号、加粗、上下标）全部丢失**。
> → 这是用户"最看重的功能"里**最大的规格缺口**，必须在方案阶段提出（见 §3.7）。

### 3.6 Excel COM 调用清单（**逐属性/方法与参数**）

> 需求方要求列全——因为我们要用纯 XML 直接读写 `.xlsx`，必须知道它碰了哪些单元格与格式。
> **好消息：它只碰单元格 `Value` 和两列的 `AutoFit`，没有字体/颜色/边框/条件格式/批注/数据验证。**

**A. 写（`段落配方生成器.bas:338-404`，`Sub CreateExcelWithData`）**

| 行 | 调用 | 参数/用法 | 对应 xlsx 部件 |
|---|---|---|---|
| `:346` | `GetObject(, "Excel.Application")` | 复用已开 Excel；失败进 `Err` | — |
| `:348` | `CreateObject("Excel.Application")` | 新建 | — |
| `:352-355` | `If excelApp Is Nothing` | Excel 未安装 → `MsgBox` 并退出 | — |
| `:358` | `excelApp.Workbooks.Add` | **新建工作簿**（无模板参数） | 整个包 |
| `:359` | `excelWorkbook.Worksheets(1)` | 取第 1 个表 | `xl/worksheets/sheet1.xml` |
| `:360` | `excelWorksheet.Name = sheetName` | 改名 | `xl/workbook.xml` 的 `<sheet name>` |
| `:363` | `Cells(1, 1).Value = "项目"` | 写 A1 | `<c r="A1" t="s"><v>索引</v></c>` + `sharedStrings.xml` |
| `:364` | `Cells(1, 2).Value = "数值"` | 写 B1 | 同上（B1） |
| `:368` | `Cells(i+1, 1).Value = varPrefixes(i)` | 写 A2…A(n+1) | 每一行一个 `<row>` |
| `:369` | `Cells(i+1, 2).Value = varTexts(i)` | 写 B2…B(n+1) | 数值型会被写成 `<c r="B2"><v>1.5</v></c>`（无 `t`） |
| `:373` | `excelWorksheet.Columns(1).AutoFit` | 自动列宽 | `<cols><col min=1 max=1 width=… customWidth=1/></cols>`（**第一版建议不写**） |
| `:374` | `excelWorksheet.Columns(2).AutoFit` | 同上 | 同上 |
| `:385` | `excelWorkbook.SaveAs fullPath` | **保存路径**（文档同目录） | 决定文件名/位置 |
| `:386-391` | `Err` 分支 → `MsgBox` + `Close False` + `Quit` | 保存失败处理 | — |
| `:395` | `excelWorkbook.Close` | 关闭（**不带参数** → 会保存提示；此时已 SaveAs 过，通常不弹） | — |
| `:396` | `excelApp.Quit` | **退出 Excel**（若复用了用户的 Excel → 关掉用户进程） | — |
| `:399-401` | `Set … = Nothing` | 释放 3 个对象 | — |

**B. 读（`段落重配.bas:78-151`，`Function GetExcelData`）**

| 行 | 调用 | 参数/用法 | 对应 xlsx 部件 |
|---|---|---|---|
| `:88` | `GetFullExcelPath(excelPath)` | 路径定位（见 §3.1） | — |
| `:95` | `CreateObject("Excel.Application")` | **总是新建**（不复用） | — |
| `:96-97` | `.Visible = False`、`.DisplayAlerts = False` | 后台打开、不弹警告 | — |
| `:100` | `excelApp.Workbooks.Open(fullPath)` | 打开工作簿 | 整个包 |
| `:104` | `excelWorkbook.Sheets(sheetName)` | **按名字取表**（`Sheets` 含图表表；`段落重配` 用 `Sheets`，写端用 `Worksheets`） | `xl/workbook.xml` 的 `<sheet name>` |
| `:105-110` | `If excelWorksheet Is Nothing` | 表不存在 → 关簿退出、返回 `False` | — |
| `:119` | `excelWorksheet.Range(cellAddress).Value` | `cellAddress = "B" & (i+1)` | `<c r="B2">…</c>` |
| `:121-127` | `IsError` → `"#错误#"`；`IsNull` 或 `""` → `""`；否则 `CStr` | 三档取值语义 | 共享字符串 / `<v>` / 空 |
| `:131` | `excelWorkbook.Close False` | **不保存关闭** | — |
| `:132` | `excelApp.Quit` | 退出 Excel | — |
| `:134-136` | `Set … = Nothing` | 释放 | — |
| `:141-151` | `ErrorHandler` | 关簿 + 退出 + 返回 `False`（**未释放对象**） | — |

**C. 需要给 XML 实现的对照结论**

- 只读/写 **A1:B(n+1)** 区域，**两张表头单元格**是固定字符串。
- **B 列是唯一有语义的列**（`段落重配` 只读它）；A 列是**只写不读**的上下文。
- **没有单元格格式**（无字体/颜色/边框/数字格式/公式/批注/合并）。
- **表名必须精确匹配**（`Sheets(sheetName)`，大小写按 Excel 规则不敏感，但**中文/空格必须一致**）。

### 3.7 XML 落地的三个决断点（**待用户裁决**）

1. **还原段落要不要带格式？** VBA 版是**纯文本**（`:277`）。XML 版可以做得更好（把 `TEXT:` 段落的 `pPr`/`rPr` 也一起存/还），但那样就**与 Word 宏不互通**。→ 需要用户在"互通"与"保真"之间选。
2. **分隔符**：VBA 用半角 `?`（已实测），XML 版应当换成一个不会冲突的分隔符（如 `\x1F` 或干脆按行），并明确**是否兼容旧配方**。
3. **A 列要不要用起来**：现在 A 列只写不读。XML 版可以让 A 列参与"变量定位"（例如按前缀把变量插回**正确的那一段**），但同样是"互通 vs 增强"的取舍。

---

## 4. 「表头格式」这条规则：**在 12 个宏里没有实现**

### 4.1 用户原话（唯一规格来源）

> 「标准的表头格式是段首空两格，然后 `表XX-X 表格名字`。后面的表格名字要通过增减与表 XX-X 之间的空格实现居中」

### 4.2 核查结论【规格：这是"没找到"，不是"没读懂"】

**12 个宏、5 个交付版 `*claw` 宏、`Normal.dotm` 的 6 个模块里，没有任何代码实现这条规则。**
`PLAN.md` §5.4 猜它在 `规划报告一键宏.bas`——**该文件全文 105 行里没有一处提到表格、表头、居中、间距**（它只做替换/对齐/字体，见 §2.12）。

**已执行的核查（逐条可复核）**：

| 核查范围 | 方法 | 结果 |
|---|---|---|
| 12 个 `.bas` | 全文搜 `表`/`居中`/`Center`/`Indent`/`Space`/`Len(`/`表XX` | **无**表头居中代码；`Indent` 只出现在 `目录修改.bas:15,32`（TOC 层级） |
| 12 个 `.bas` | 搜 `alignment` | 只有 `规划报告一键宏.bas:48,50,61,62`（两端→左） |
| `reference/delivered-macros/` 5 个 `*claw.bas` | 同关键词 | **无**（`表格边框调整claw.bas` 是边框，不是表头） |
| `Normal.dotm` 模块清单 | 读 `MacroToolbox/data/state/template_snapshot.json` | 6 个模块：`表格边框调整claw`、`表格空格回车删除claw`、`格式规范化claw`、`规划报告一键宏claw`、`目录修改_claw`、`替代宏_替换宏1` → **无表头宏** |
| `MacroToolbox` 库（6 个 `m-*.bas`） | 搜关键词 | 无表头宏 |
| 整个 `E:\Zspace`（`.py/.md/.bas/.json/.txt`） | 搜 `表XX`/`表头`/`居中`/`空两格` | 命中的都是**别的语义**（`EngiBlock` 的"表头行检测"、MacroToolbox 的"居中对齐枚举"）→ **无实现** |
| `D:\百度网盘\宏相关（主机）` 全目录 | 搜 `表XX`/`居中`/`表头` | 只有 CAD 插件的 `表头`（无关）、Excel 宏的 `空格`（无关） |
| `D:\百度网盘` 是否有 `.docm/.dotm` | 按扩展名枚举 | **无**（该盘只有这 12 个 Word 宏） |

### 4.3 因此：这条规则**必须由用户补齐**——以下是缺的参数

用户的描述里，只有"段首空两格"是明确的；其余都得问：

| 待确认项 | 为什么要问 |
|---|---|
| **"两格"是什么** | 两个**半角空格**（U+0020 + U+0020）？两个**全角空格**（U+3000 ×2）？还是 `w:ind/@w:firstLine = 2 字符`（即 `firstLineChars="200"`）？三种在 XML 里完全不同 |
| **`表XX-X` 的编号规则** | `XX` 是**章号**（如 `表5-1`）还是**表序号**？`-X` 是章内序号？与哪个标题级别联动？是否全文重排？"表"字后有没有空格（原话里写 `表XX-X` 还是 `表 XX-X`）？ |
| **空格的"总宽度"取多少** | 原话说"通过增减空格实现居中"——居中需要知道**目标版心宽度**（= 页面宽 - 左右页边距，单位是**字符**还是 **pt**）与**当前行内容宽度**。这两个数从哪来？ |
| **怎么算需要几个空格** | 公式是什么？`(总宽 - 编号宽 - 名字宽) / 2 / 单空格宽` 取整？还是固定配比？**小数怎么办（左右各半，总宽奇数）**？ |
| **用什么单位"量宽度"** | `表5-1` 是 4 个字符宽（半角）/ 4 个全角宽？中文字与数字的宽度比是否按"1 汉字 = 2 半角"折算？**这是公式能不能落地的关键** |
| **段落本身的对齐方式** | 该段落是 `居中`（`w:jc="center"`）还是 `左对齐 + 空格`？如果是"左对齐 + 空格"，那"段首空两格"和"居中空格"是**两段独立的空格**？ |
| **作用范围** | 全文所有 `表X-Y` 段落？还是只处理表格**紧邻上方**的那一段？如何识别一个段落是"表头"（正则 `^表\s*\d+-\d+`？） |
| **幂等** | 已经居中过的段落再跑一次，会不会越加越多？（用户说"增减"，说明要实现**重算**而不是追加） |

### 4.4 我**不**编的东西

- 我**不给**"公式的量级猜测"（例如"按 40 字宽算"），因为没有任何代码或文档支持这个数。
- 我**不**把 `目录修改.bas` 的 `28/56 pt` 当作表头公式——那是**目录缩进**的常数（§2.10），与表头空格无关。
- 我**不**把 `MacroToolbox/toolbox/core/generator.py:107` 的 `u"居中": 1` 当作线索——那是**对齐枚举的中文名映射**（`左/居中/右` → 0/1/2），与表头排版无关。

**→ 建议的推进方式**：让用户在 Word 里**手工做一个样例**（一段居中的表头 + 一张表），把那个 `.docx` 给出来，
我们用 `wordfactory inspect` 读它的 `w:ind` / `w:jc` / 空格字符码 —— **从样本反推公式**比继续找宏可靠得多。

---

## 5. 三类规则：替换表 / 去空格 / 格式规范化

### 5.1 文本替换对照表（**字符 → 字符**）

**A. `规划报告一键宏.bas`（全文，`wdReplaceAll`）**

| 查找 | 替换 | 大小写 | 整词 | 出处 |
|---|---|---|---|---|
| `其它` | `其他` | 不敏感（`MatchCase=False`） | 否（`MatchWholeWord=False`） | `规划报告一键宏.bas:20-21` |
| `东流流经` | `向东流经` | 不敏感 | 否 | `规划报告一键宏.bas:33-34` |

> 交付版把这两条**变成可配置常量且默认关闭**：`规划报告一键宏claw.bas:5-10`
> （`CFG_替换文本_1 = "旧文本A"`、`CFG_启用文本替换_1 = False`）→ **交付版默认不替换任何文本**。

**B. `MarkDown语言清除.bas`（仅选中范围；按正则顺序执行）**

| 序 | 匹配 | 替换为 | 说明 | 出处 |
|---|---|---|---|---|
| 1 | `` `([^`]+)` `` | `“$1”` | 行内代码 → **中文左右双引号**（`U+201C`/`U+201D`） | `MarkDown语言清除.bas:37-38` |
| 2 | `\*\*([^*]+)\*\*` | `$1` | 粗体标记去掉 | `:41-42` |
| 3 | `\*([^*]+)\*` | `$1` | 斜体标记去掉 | `:44-45` |
| 4 | 段首所有 `#` | 删除 | `Do While` 循环剥干净 | `:55-62` |
| 5 | 段首 1 个 `*` / `-` / `+` | 删除（只一个） | `LTrim` 后判断 | `:64-67` |
| 6 | 段内任何 `*` | 删除 | `Replace(para,"*","")` | `:69` |
| 7 | `#.`/`##.` 开头的整段 | 删掉首个 `.` 及之前 | **死代码**（`#` 已在序 4 被剥） | `:71-81` |
| 8 | 连续 `"  "`（两空格） | `" "`（一空格） | 循环折叠到没有双空格 | `:83-85` |

**C. 只改格式、不改字符的"替换"（列在这里以免混淆）**

| 宏 | 查找条件 | 改成 | 出处 |
|---|---|---|---|
| `规划报告一键宏.bas` | 段落对齐 = 两端对齐 | 左对齐 | `:47-51` |
| `规划报告一键宏.bas` | 中文字体 = `仿宋` | `宋体` | `:70-74` |
| `格式规范化.bas` | `m[0-9]{1,}` 的数字部分 | 上标 | `:42`、`:57` |
| `单位上标.bas` | `m[2-9]{1,}` 的数字部分 | 上标 | `:11`、`:26` |
| `智能上下标*.bas` | 见 §2.5/§2.6 字典 | 上标/下标/清除 | `智能上下标.bas:76-96` 等 |

**D. 上下标符号字典（复刻用的权威表）**

| 键 | 规则串 | 逐字符 | FSM 版 | 正则版 |
|---|---|---|---|---|
| `Vmax` | `NBBB` | V正常 m下 a下 x下 | ✔ `:76` | ✔ `:85` |
| `Qpl` | `NBB` | Q正常 p下 l下 | ✔ `:77` | ✔ `:86` |
| `m2` | `NS` | m正常 2上 | ✔ `:80` | ✔ `:87` |
| `m3` | `NS` | m正常 3上 | ✔ `:81` | ✔ `:88` |
| `m4` | `NS` | m正常 4上 | ✔ `:82` | ✔ `:89` |
| `m5` | `NS` | m正常 5上 | ✔ `:83` | ✔ `:90` |
| `cm2` | `NNS` | c正常 m正常 2上 | ✔ `:84` | ✔ `:91` |
| `km2` | `NNS` | k正常 m正常 2上 | ✔ `:85` | ✔ `:92` |
| `qm` | `NB` | q正常 m下 | ✔ `:86` | ✔ `:93` |
| `Qm` | `NB` | Q正常 m下 | ✔ `:87` | ✔ `:94` |
| `KP` | `NB` | K正常 P下 | ✔ `:88` | ✔ `:95` |
| `CV` | `NB` | C正常 V下 | ✔ `:89` | ✔ `:96` |
| `H24P` | `NBBB` | H正常 2下 4下 P下 | ✔ `:90` | ✔ `:97` |
| `hR` | `NB` | h正常 R下 | ✔ `:91` | ✔ `:98` |
| `H24` | `NBB` | H正常 2下 4下 | ✔ `:92` | ✔ `:99` |
| `Q4ml` | `NBSS` | Q正常 4下 m上 l上 | ✘ | ✔ `:100` |
| `Q4el` | `NBSS` | Q正常 4下 e上 l上 | ✘ | ✔ `:101` |
| `Pt` | `NB` | P正常 t下 | ✘ | ✔ `:102` |
| `Q4al+pl` | `NBSSSSS` | Q正常 4下 a上 l上 +上 p上 l上 | ✘ | ✔ `:103` |
| `°C`（特殊） | `SN` | °上 C正常 | ✔ `:95`（**几乎不生效**，见 §2.5） | ✔ `:79`（生效） |
| `mm`（特殊） | `NN` | 两字符都清除上下标 | ✔ `:96` | ✔ `:80` |

**FSM 版的额外"智能检测"（字典之外，`:310-338`）**：

| 条件 | 规则 | 出处 |
|---|---|---|
| `cleanText = "m2"` | `NS` | `:312-315` |
| `cleanText = "m3"` | `NS` | `:317-320` |
| 首字符 `m` 且长度 = 2 且第 2 位是数字 | `NS`（→ 覆盖 `m4`~`m9`） | `:323-326` |
| 尾字符 `2` 且长度 > 1 且首字符是字母 | `N`×(len-1) + `S` | `:329-332` |
| 尾字符 `3` 且长度 > 1 且首字符是字母 | `N`×(len-1) + `S` | `:335-338` |

**边界字符表（只有 FSM 版有）**：`，` `。` `！` `？` `；` `：` `"` `=` `-` `+` `*` `×` `/` `÷` `（` `）` `【` `】` `《` `》` `＝` `％` — `智能上下标.bas:103`

### 5.2 去空格：「无意义空格」的判定条件

**宏 1：`表格空格回车删除.bas`（表格单元格内）**

判定条件 = **位置无关，单元格内所有匹配字符都删**（不是"仅首尾"）：

| 档位 | 删除的字符 | 出处 |
|---|---|---|
| 1 仅空格 | `" "`(U+0020)、`Chr(160)`(NBSP) | `:146-147` |
| 2 仅回车 | `Chr(13)`(段落标记)、`Chr(11)`(手动换行)、`Chr(7)`(单元格标记) | `:149-151` |
| 3 两者 | 上述 5 个字符 | `:153-157` |
| **默认档位** | **3**（`InputBox` 默认值 `"3"`） | `:182` |
| 前置守卫 | `If Len(originalText) > 0` 才处理 | `:142` |
| 写回条件 | `If newText <> originalText`（无变化不写） | `:161` |

**不删的（重要缺口）**：

| 字符 | Unicode | 是否删除 | 说明 |
|---|---|---|---|
| **全角空格 `　`** | U+3000 | **不删** | 中文报告里最常见→**待用户确认是否补** |
| 制表符 `Chr(9)` | U+0009 | **不删** | 交付版 `表格空格回车删除claw.bas:14` **补上了 `vbTab`** |
| `Chr(10)` | U+000A | 不删 | Word 里通常只出现 `Chr(13)`，但**外部生成的 docx 可能只有 `Chr(10)`** |
| 其它 Unicode 空白（`U+2002` 等） | — | 不删 | — |

**宏 2：`MarkDown语言清除.bas`（选中范围内）**

| 判定 | 动作 | 出处 |
|---|---|---|
| 段首 `Trim` 后的空白 | `Trim` 掉 | `:57`、`:64`（`LTrim`）、`:87`（`Trim`） |
| 连续的**两个**半角空格 | 折叠成 1 个（循环，直到没有双空格） | `:83-85` |
| 列表标记后 | `Trim(Mid(para,2))` | `:66` |
| 结果 | **保留单个空格**（不删光） | — |

### 5.3 格式规范化：动作清单（`格式规范化.bas`）

**完整动作清单（就这 6 条，没有别的）**：

| # | 动作 | 作用范围 | Word 属性/方法 | 出处 |
|---|---|---|---|---|
| 1 | **全文变黑** | **全故事**（`StoryRanges` 全链） | `rng.Font.Color = wdColorBlack` | `:13-20` |
| 2 | **去高亮** | **全故事** | `rng.HighlightColorIndex = wdNoHighlight` | `:23-30` |
| 3 | **`m` 后数字变上标** | 正文（`ActiveDocument.Range`） | `rng.Font.Superscript = True` | `:32-68` |
| 4 | **更新目录页码** | 第 1 个目录 | `TablesOfContents(1).UpdatePageNumbers` | `:70-79` |
| 5 | 汇总提示 | — | `MsgBox` | `:82-94` |
| 6 | **询问保存并关闭** | 整篇文档 | `Save` / `SaveAs` 对话框 / `Close wdSaveChanges` | `:96-112` |

**"还有没有别的"——答案：没有。** 该宏**不**做：
- 字号（没有 `Font.Size`）
- 字体名（没有 `Name`/`NameFarEast`）
- 行距/段间距（没有 `LineSpacing`/`SpaceBefore`/`SpaceAfter`）
- 缩进（没有 `Indent`）
- 对齐（没有 `Alignment`）
- 表格边框（没有 `Borders`）
- 清理空格（那是 `表格空格回车删除.bas` 的事）

**交付版 `格式规范化claw.bas` 的差异（`reference/delivered-macros/格式规范化claw.bas`，24 行）**：

| 差异 | 原件 | 交付版 |
|---|---|---|
| 故事遍历 | `For Each rng In StoryRanges` + `NextStoryRange` 链 | `For st = 1 To 11` 遍历 + 链（`:5-12`） |
| 变黑属性 | `Font.Color = wdColorBlack`（`:16`） | **`Font.ColorIndex = wdBlack`**（`:7`） |
| Find 范围 | `wdFindContinue`（`:45`） | `wdFindStop`（`:14`） |
| 上标写法 | 改 `rng` 自身（`:54-57`） | `ActiveDocument.Range(s, e).Font.Superscript`（`:19`），**有 `If s < e` 守卫**（`:19`） |
| 防死循环 | `Collapse` + `End = originalEnd`（`:61-62`） | `fr.Start = fr.End: fr.End = oe` + `If fr.Start >= oe Then Exit Do`（`:20-21`） |
| 目录页码 | **有**（`:70-79`） | **无** |
| 保存并关闭 | **有**（`:96-112`） | **无** |
| 完成提示 | `MsgBox` 汇总 | **无**（静默） |

→ **交付版是"清理过、去掉危险动作、静默执行"的版本**。两版都可当规格，但**必须选定一个**。

---

## 6. 未能确定（**不编造**）

按重要性排序。每条都注明"我猜的是什么"以及"怎么才能确定"。

| # | 未确定事项 | 我的猜测（**未经验证**） | 怎么确定 |
|---|---|---|---|
| 1 | **表头格式（段首两格 + `表XX-X` + 空格居中）的算法** | 12 个宏里没有；可能是用户在 Word 里**手工做的**，或写在**别的机器的模板**里 | 让用户**手工做一个样例 docx**，用 `wordfactory inspect` 反推（§4.4） |
| 2 | **`?` 为何是配方行的分隔符**（`段落重配.bas:47-50`、`:196-199`） | 原文是某个特殊控制字符（如 `Chr(1)`），在复制/保存链路里被替换成半角 `?` | 问用户；或看 Word 里实际生成的配方文本里那个字符长什么样 |
| 3 | **`智能上下标.bas` 的 `°C` 规则是否真的不生效** | 推演：几乎不生效（`°` 前是数字时状态已不在 START） | 在 Word 里实测 `25°C` / `温度 °C` 两种输入 |
| 4 | **"隐形修正"的精确语义**（`段落配方生成器.bas:177-213`） | 为绕过高亮区末尾的 `HighlightColorIndex` 读取偏差，把最后一个高亮字符划给后面的 TEXT | Word + 宏跑一份金标准，逐字符对比变量边界 |
| 5 | **`单位上标.bas:18` 的 `MoveStart` 副作用是否导致漏匹配** | 推演：连写场景（`m2m3`）可能漏 | Word 里实测 `m2m3`、`m2 m3`、`m10`（`{1,}` 会匹配 `m10` 吗？）、`m1` |
| 6 | **`单位上标` 的 `[2-9]` 与 `格式规范化` 的 `[0-9]` 哪个是正确规格** | 无法判断——两台机器上两个宏行为不同 | **用户裁决** |
| 7 | **两份 `智能上下标` 哪一份是规格**（两文件同名 `Attribute VB_Name = "智能上下标"`，不可能共存） | 正则版是"重构版"，可能更新；但 FSM 版标题说"修复边界"，也可能更新 | **用户裁决**（可问"你平时按哪个"） |
| 8 | **`全选表格.bas` 的净效果** | 推演：最终不落盘任何 `w:permStart`（自己加的权限自己删了） | 在 Word 里跑一次，另存为 docx，解包看有没有 `w:permStart` |
| 9 | **要不要删全角空格 U+3000** | 推演：用户**以为**它删了（中文报告里全是全角空格） | **用户确认**；建议默认补上 |
| 10 | **`目录修改.bas` 选三级目录时什么都不做，是不是 bug** | 是 bug（`Case 56` 设了 3，但删除条件里没有 3 的分支） | **用户确认**（是修还是保留） |
| 11 | **`Chr(7)` 的 `Replace` 在真实文档里会不会删掉嵌套结构** | 通常已被 `MoveEnd -1` 排除，理论上不发生 | 造一个含嵌套表格的夹具实测 |
| 12 | **`段落重配` 丢格式是否可接受** | 用户"最看重"这个功能，而它**只还原纯文本**（编号/样式/缩进全丢）→ 很可能不满意 | **用户必须裁决**（是"互通优先"还是"保真优先"，见 §3.7） |
| 13 | **`规划报告一键宏` 的字体到底是 `仿宋` 还是 `仿宋_GB2312`** | 原件 = `仿宋`（`:71`），交付版 = `仿宋_GB2312`（`交付版:11`） | **用户确认**模板里的实际字体名 |
| 14 | **配方的 A 列到底该不该用** | 现在 A 列只写不读；用起来可增强定位 | **用户裁决**（§3.7） |
| 15 | **`VAR:` 行里 `|路径!Sheet!B2` 字段无人读取** | 生成端冗余，消费端只靠头部 3 行 | 若要第三方互通，**以头部 3 行为准**（这样最省） |
| 16 | **Word 的 `Range` 位置 ↔ 字符串下标的换算细节**（影响 §2.3/§2.5 的偏移） | 简单文本 1:1；含域/表格标记时不等价 | 造夹具实测（域、表格、`w:br`） |
| 17 | **`ActiveDocument.Range` 是否覆盖文本框/页眉** | 推演：**否**（只有 `StoryRanges` 才覆盖）→ 与 `格式规范化` 的范围不同 | 实测（在文本框里放 `m2`，跑 `单位上标` 看改不改） |

---

## 附录 A：交付版 `*claw` 宏（`reference/delivered-macros/`，只读参考）

> 这 5 个文件是从**真 `Normal.dotm`** 导出的（`MacroToolbox/data/state/template_snapshot.json` 记录了模板路径与 sha256）。
> 编码是 **UTF-8**（与 `reference/vba-macros/` 的 cp936 不同）。它们是**清理过的执行版**，不是原始规格。

### A.1 `目录修改_claw.bas`（8 行）— 与文件名的宏完全无关
```
目录修改_claw.bas:3-8
Sub 静默更新目录页码_claw()
Dim toc As TableOfContents
For Each toc In ActiveDocument.TablesOfContents
    toc.UpdatePageNumbers
Next
End Sub
```
→ **文件名 `目录修改_claw` 里装的是"静默更新所有目录页码"**（对每个目录都更新，而 `格式规范化.bas:73` 只更新第 1 个）。
→ **这是 XML 路线做不到的功能**（需分页引擎），但它是**"用户实际在用的东西"**，值得在方案里明确"不做 + 替代路线"（`PLAN.md` §6 已列四条路）。

### A.2 `表格边框调整claw.bas`（13 行）— **不在 12 个原件里的第 13 个功能**
```
表格边框调整claw.bas:4-12
Dim tbl As Table: Application.ScreenUpdating = False
For Each tbl In ActiveDocument.Tables
    With tbl.Borders(wdBorderTop): .LineStyle = wdLineStyleSingle: .LineWidth = wdLineWidth150pt: End With
    With tbl.Borders(wdBorderBottom): .LineStyle = wdLineStyleSingle: .LineWidth = wdLineWidth150pt: End With
    With tbl.Borders(wdBorderLeft): .LineStyle = wdLineStyleSingle: .LineWidth = wdLineWidth150pt: End With
    With tbl.Borders(wdBorderRight): .LineStyle = wdLineStyleSingle: .LineWidth = wdLineWidth150pt: End With
    With tbl.Borders(wdBorderInsideHorizontal): .LineStyle = wdLineStyleSingle: .LineWidth = wdLineWidth050pt: End With
    With tbl.Borders(wdBorderInsideVertical): .LineStyle = wdLineStyleSingle: .LineWidth = wdLineWidth050pt: End With
Next: Application.ScreenUpdating = True
```
**写死的边框规则**：外框 4 边 = 单线 **1.5 pt**；内横/内竖 = 单线 **0.5 pt**。
→ OOXML：`w:tblPr/w:tblBorders/{top,left,bottom,right}` 用 `w:sz="12"`（1/8 pt 单位 → 1.5pt×8=12），
`{insideH,insideV}` 用 `w:sz="4"`（0.5pt×8=4），`w:val="single"`。（**待验证**：Word 的 `wdLineWidth150pt` 是否为 1.5pt，以及 Word 实际写出的 `w:sz` 值。）

### A.3 `表格空格回车删除claw.bas`（23 行）— **语义与原件相反**
```
表格空格回车删除claw.bas:7-21
For Each tbl In ActiveDocument.Tables
    For r = 1 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count
            Dim txt As String
            txt = tbl.Cell(r, c).Range.Text
            Do While Len(txt) > 0
                Dim ch As String: ch = Right(txt, 1)
                If ch = Chr(13) Or ch = Chr(7) Or ch = " " Or ch = vbTab Then
                    txt = Left(txt, Len(txt) - 1)
                Else: Exit Do: End If
            Loop
            If Len(txt) > 0 Then tbl.Cell(r, c).Range.Text = txt
        Next
    Next
Next
```
| 维度 | 原件 `表格空格回车删除.bas` | 交付版 `claw` |
|---|---|---|
| 删除位置 | **全格所有位置**（`Replace`） | **只裁尾部**（循环 `Right(txt,1)`） |
| 字符集 | 空格、NBSP、`Chr(13)`、`Chr(11)`、`Chr(7)` | 空格、**`vbTab`**、`Chr(13)`、`Chr(7)`（**没有 `Chr(11)`、没有 NBSP**） |
| 遍历 | `tbl.Range.Cells`（含合并单元格的正确集合） | `For r / For c` + `tbl.Cell(r,c)`（**合并单元格会报错/重复**） |
| 档位 | 3 档可选 | 无（只有一档） |
| 空单元格 | 一起处理 | `If Len(txt) > 0` 才写回（`:18`） |

→ **两版的"删除判据"完全不同**（"全删" vs "只裁尾"）。**必须让用户指定哪个是他要的**（推演：报告里单元格首尾常有多余回车，交付版的"只裁尾"更可能是用户的真实需求）。

### A.4 `规划报告一键宏claw.bas`（81 行）
见 §2.12 与 §5.1.A 的差异说明。结构是"CONFIG 配置区 + 执行体"，**5 个功能各有 `CFG_启用_*` 开关，全部默认 `False`**（`:7` `:10` `:13` `:16` `:19`）→ 默认行为是"什么都不做"。
配置项：`CFG_原字体="仿宋_GB2312"`、`CFG_新字体="宋体"`、`CFG_原对齐=3`、`CFG_新对齐=0`、`CFG_原字号=12`、`CFG_新字号=14`（**字号功能在原件里没有**）。

### A.5 `格式规范化claw.bas`（24 行）
见 §5.3 的差异表。

### A.6 第 6 个模块（未导出为文件）
`Normal.dotm` 里还有 `替代宏_替换宏1`（82 行，`MacroToolbox/data/state/template_snapshot.json`），**没有导出到 `reference/delivered-macros/`**。
→ 名字像是"替换类宏"，**未读到内容**（列在 §6 之外，仅作提示：若需要，可从 `MacroToolbox/data/backup/*/template.dotm` 里导出核对）。

---

## 附录 B：原件、副本、交付版三者对照

| 项 | 路径 | 编码 | 用途 |
|---|---|---|---|
| **原件（规格来源）** | `D:\百度网盘\宏相关（主机）\word宏\*.bas`（12 个） | cp936 | **只读**，本文行号的依据 |
| 项目内副本 | `E:\Zspace\projects\word-factory\reference\vba-macros\*.bas`（12 个） | cp936（同字节） | 只读参考，行号一致 |
| 交付版 | `E:\Zspace\projects\word-factory\reference\delivered-macros\*.bas`（5 个） | UTF-8 | 从真 `Normal.dotm` 导出，**已交付并验收通过** |
| 前身项目 | `E:\Zspace\projects\MacroToolbox\` | — | 宏管理工具箱（Python + VBE 后端），**本项目不改动它** |

---
