---
name: md-to-docx
description: 把 Markdown(.md) 文件导出为 Word(.docx) 文件。当用户要求把 markdown 转成 docx/Word、导出为党政机关公文格式、导出为党政公文格式的红头/正式文稿、或导出为浙江大学学位论文/学术文档格式时使用本 skill。内置两套版式 preset：official（GB/T 9704-2012 党政公文，仿宋三号正文、黑体一级标题、楷体加粗二级标题、五号宋体表格、四号黑体表题、黑色字体、奇偶页码）与 zju（浙江大学研究生学位论文，仿宋小四正文 20 磅行距、每章另起页、页眉 STYLEREF）。支持 md 标题层级到 docx 标题样式的映射（md # → docx Title，md ## → Heading 1，以此类推）；支持在 md 中独占一行写 HTML 注释形式的指令插入 Word 分节符（`<!-- 分节符 -->` 下一页、`<!-- 分节符: 连续 -->`、`: 偶数页`、`: 奇数页`）。
agent_created: true
allowed-tools: Bash, Read, Write, Edit, Glob, Grep
---

# md-to-docx

将 Markdown 转换为符合中文排版规范的 `.docx`，内置两套版式。

## 何时使用

- 用户要求「把 md 转成 docx / Word」
- 用户要求输出**党政机关公文格式**
- 用户要求输出**浙江大学学位论文 / 学术文档格式**

## 使用方式

用受管 Python 运行（已装 `python-docx`）：

```bash
PY="C:/Users/hongq/.workbuddy/binaries/python/envs/default/Scripts/python.exe"

# 党政机关公文格式（默认）
"$PY" scripts/md2docx.py 输入.md -o 输出.docx --mode official

# 浙江大学学位论文格式
"$PY" scripts/md2docx.py 输入.md -o 输出.docx --mode zju
```

若当前 Python 缺少依赖：`"$PY" -m pip install python-docx`。

`scripts/md2docx.py` 可以换成 SKILL 安装后的绝对路径；`-o` 省略时输出 `<同名>_<模式>.docx`。

## 标题层级映射（固定规则，不要改动）

| Markdown | Word 样式（中文界面） | 内置样式名 | 含义（公文）               | 含义（浙大）                   |
| :------- | :-------------------- | :--------- | :------------------------- | :----------------------------- |
| `#`      | 标题                  | Title      | 文件标题（二号小标宋居中） | 论文题目（小二仿宋加粗居中）   |
| `##`     | 标题 1                | Heading 1  | 一级标题（黑体三号）       | 章标题（三号仿宋加粗居中）     |
| `###`    | 标题 2                | Heading 2  | 二级标题（楷体三号加粗）   | 一级节标题（四号仿宋加粗顶左） |
| `####`   | 标题 3                | Heading 3  | 三级标题（仿宋三号加粗）   | 二级节标题（小四仿宋加粗顶左） |
| `#####`  | 标题 4                | Heading 4  | 四级标题                   | 三级节标题（小四仿宋顶左）     |
| `######` | 标题 5                | Heading 5  | 五级标题                   | 四级节标题                     |

> 中文版 Word 把内置样式显示成「标题」「标题 1」「题注」等中文名，而 `.docx` 文件里的
> `w:name` 仍然存 `title` / `heading 1`。脚本写入时用内置名，解析时英文与中文
> （`标题 1`、`标题1`）都能命中；页眉 STYLEREF 也会按文档里实际存在的样式名输出。

## 两套版式速查

### official（党政机关公文，GB/T 9704-2012）

- 页面：A4，上 3.7 / 下 3.5 / 左 2.8 / 右 2.6 cm（版心 156 × 225 mm）
- 正文：仿宋三号（16 pt），两端对齐，首行缩进 2 字符，行距固定值 28.8 磅（每面 22 行）
- 一级标题黑体三号；二级标题楷体三号**加粗**；三级标题仿宋三号加粗
- 表格正文字体：**宋体五号（10.5 pt）**；**表格标题：黑体四号**，居中置于表上方
- 注释/附注：仿宋三号，左缩进 2 字符
- **所有文字颜色为黑色**（含超链接）
- 页码：宋体四号，左右加一字线，**单页码居右空一字、双页码居左空一字**（奇偶页不同）
- **不生成**红头（发文机关标志、发文字号、签发人、红色分隔线），**不生成**版记（抄送机关、印发机关、印发日期），**不生成**封面/封底

### zju（浙江大学研究生学位论文）

- 页面：A4，上下 2.54 cm，左右 3.17 cm，页眉 1.5 cm，页脚 1.75 cm
- 正文：仿宋小四（12 pt），20 磅行距，首行缩进 2 字符，两端对齐
- 章标题居中三号仿宋加粗（段前 24 磅、段后 18 磅）；一级节 四号加粗；二级节 小四加粗；三级节 小四不加粗
- **每一章另起页**（自动在第二个及以后的 Heading 1 前分页）
- 页眉：宋体小五号，左侧「浙江大学学位论文」，右侧用 STYLEREF 取当前章名，带单横线
- 页码：宋体小五号，页脚居中
- 表题在表上方、图题在图下方，均仿宋五号

完整取值表见 `references/official-format.md` 与 `references/zju-format.md`，需要向用户解释或微调时先读这两份。

## Markdown 书写约定

- **表题**：在表格**紧上方**独占一行写 `表 1 xxx` 或 `Table 1 xxx`，脚本会自动识别为表题并应用表题字体。
  写在表格下方也能识别（自动转为表题）。不支持带字母的编号请用 `表 A.1` 形式，已支持。
- **图题**：图片 `![说明](路径)` 之后独占一行写 `图 1 xxx`，置于图下方。
- 表格对齐 `| :--- | :---: | ---: |` 会被解析（列对齐目前统一居中，如需改请调整 `cfg.table_cell_align`）。
- 引用块 `>` 渲染为「注释/附注」样式。
- 列表支持有序/无序与嵌套，编号支持 `1.` / `（1）` / `①` 等。
- YAML front matter 会自动跳过。
- **HTML 注释** `<!-- ... -->`（含行内）一律不写入正文。

### 分节符（=== 独占一行的 HTML 注释指令 ===）

在 md 里独占一行写注释形式的指令，导出时在该位置插入 Word 分节符：

```markdown
<!-- 分节符 -->                    下一页（默认，新节从下一页开始）
<!-- 分节符: 连续 -->              连续（不另起页，仅开始新节，可用于换页边距/纸张方向/分栏）
<!-- 分节符: 偶数页 -->            新节从下一个偶数页开始
<!-- 分节符: 奇数页 -->            新节从下一个奇数页开始

<!-- section-break -->             英文写法，等价「分节符」
<!-- section-break: continuous --> 英文写法，类型可用 continuous / even-page / odd-page / next-page
```

- 类型名中文英文都认（`下一页|新页|next-page`、`连续|continuous`、`偶数页|even`、`奇数页|odd`），
  大小写与 `-`/`_` 不敏感；写错或不写按「下一页」处理。
- 分节符**之前**那一节会完整保留页面设置、页眉页脚与奇偶页页码；分节符**之后**的内容沿用全文
  页面设置与页眉页脚（body 级 sectPr），页码连续编号。
- 承载分节符的段落被压到最小行高（1 磅、字号 1 磅），不会多出空行；文档末尾的分节符不会被
  「清理尾部空段落」误删。
- 与 zju 模式的「每章另起页」是两套机制：前者是分节符，后者是段落 `pageBreakBefore`。

样例见 `examples/sample-section.md`。

## 常用参数

| 参数                                               | 作用                                                            |
| :------------------------------------------------- | :-------------------------------------------------------------- |
| `-m, --mode`                                       | `official`（默认）或 `zju`                                      |
| `--line-spacing`                                   | 覆盖正文行距，可填磅值（`28.8`）或倍数（`1.5`）                 |
| `--zju-line-spacing`                               | `fixed20`（默认 20 磅）或 `multi15`（1.5 倍）                   |
| `--title-font`                                     | 公文标题字体，例如 `--title-font 黑体`                          |
| `--official-title-style`                           | `standard`=二号小标宋居中（默认）；`heading`=按正文一级标题处理 |
| `--prefer-fangsong-gb2312`                         | 正文仿宋改用「仿宋\_GB2312」                                    |
| `--table-cell-font` / `--table-cell-size`          | 表格正文字体/字号（默认宋体、五号）                             |
| `--no-page-number` / `--no-odd-even`               | 关页码 / 页码不区分奇偶                                         |
| `--no-header` / `--header-left` / `--header-right` | 页眉控制（zju 模式）                                            |
| `--no-chapter-page-break`                          | zju 模式下每一章不另起页                                        |
| `--bullet-style`                                   | `dash`（默认）/ `dot` / `disc`                                  |
| `--ordinal-style`                                  | `arabic`（默认）/ `chinese` / `circle`（①②③）                   |
| `--auto-number`                                    | 为 md 二级及以下标题自动生成 `一、/（一）/1.` 序号              |
| `--no-font-check`                                  | 关闭「字体未安装」提示                                          |

## 字体缺失处理

脚本在运行时检测所需中文字体是否安装（读注册表 + 字体目录），缺失时在 stderr 提示。
常见情况：

- 缺「方正小标宋简体」→ 用 `--title-font 华文中宋` 或 `--title-font 黑体`
- 缺「仿宋\_GB2312」→ 保持默认「仿宋」即可，GB/T 9704-2012 本身只要求「仿宋体」

## 回归验证（改动脚本后必跑）

`tests/baseline/` 下的 6 个 docx（2 个样例 md + 分节符样例，各 official / zju 两种模式）是
**预期输出基线**，用来验证 skill 是否仍然有效：

```bash
"$PY" scripts/verify.py            # 重新生成并与基线逐项比对，全部一致退出码 0
"$PY" scripts/verify.py --update   # 版式 preset 有意调整后，重新生成基线
"$PY" scripts/verify.py --keep     # 保留本次生成的 docx，便于人工打开核对
"$PY" scripts/check_output.py      # 轻量检查：r:id 引用与部件完整性
```

比对项：段落样式名、文本、对齐、缩进、行距、段前后间距、每个 run 的字体 XML（字体/字号/
加粗）、表格结构与单元格文本、页面与页边距、页眉页脚 XML（含 STYLEREF）、图片部件、
r:id 引用完整性。

流程：改 `scripts/md2docx.py` → 跑 `verify.py` → 出现 FAIL 时确认是有意调整还是回归缺陷；
有意调整则确认差异无误后执行 `--update` 刷新基线。基线是二进制，靠脚本比对属性，不要靠
肉眼 diff。

## 交付前检查

生成后用 `python-docx` 抽查段落样式的字体、字号、加粗、缩进、行距是否落到位（参见
`examples/` 下的样例 md）。如用户对某一处不满意，优先改 `PStyle` 预设而不是临时打补丁。

## 文件结构

```
md-to-docx/
├── SKILL.md
├── scripts/
│   ├── md2docx.py            # 转换器（自研 markdown 解析 + python-docx 渲染）
│   ├── verify.py             # 回归验证：重跑样例并与基线逐项比对
│   └── check_output.py       # 轻量检查：r:id 引用与部件完整性
├── references/
│   ├── official-format.md    # 党政公文格式取值表
│   └── zju-format.md         # 浙大学位论文格式取值表
├── examples/                 # 输入样例（md + 图片）
│   ├── sample-official.md    # 公文样例
│   ├── sample-figure.md      # 含图片/图题/表题的样例
│   ├── sample-section.md     # 分节符样例（下一页/连续/偶数页/奇数页 + 注释）
│   └── fig-testing.png
└── tests/baseline/           # 样例产物基线（用于回归验证，可由 verify.py --update 重建）
```

> 用户实际生成的 docx 写到 `-o` 指定的路径，不落在 skill 目录内。
