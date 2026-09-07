# MDTT — Markdown Trans Tool

## 一、基础信息

MDTT 是一款 Markdown 与 Word（docx）互转命令行工具。支持双向转换：

- **Markdown → docx**：`MDTT file.md [参数]`，可深度定制版式（页面、字体、页眉页脚等）

- **docx → Markdown**：`MDTT file.docx [参数]`，保留标题、列表、表格、加粗/斜体、超链接、图片等结构；接受修订并记录原文，提取批注，线性化文本框和分节内容

### 技术栈

- [markdown-it](https://github.com/markdown-it/markdown-it) — Markdown 解析（md → docx）

- [docx](https://github.com/dolanmiu/docx) — docx 文档生成（md → docx）

- [image-size](https://github.com/image-size/image-size) — 图片尺寸读取（用于按比例缩放）

- [mammoth](https://github.com/mwilliamson/mammoth.js) — docx → HTML 提取（docx → md）

- [node-html-parser](https://github.com/taoqf/node-html-parser) — HTML DOM 预处理（表格结构重建与合并单元格检测）

- [turndown](https://github.com/mixmark-io/turndown) + [turndown-plugin-gfm](https://github.com/mixmark-io/turndown-plugin-gfm) — HTML → Markdown（docx → md）

- [jszip](https://github.com/Stuk/jszip) + [xml-js](https://github.com/nashwaan/xml-js) — docx 解包与 OOXML 解析（版式提取）

### 项目结构

```
MDTransTool/
├── .github/workflows/ci.yml  # CI：push/PR 时在 Ubuntu/Windows × Node 18/20/22 上自动测试
├── src/
│   ├── cli.js                # 命令行入口：双向转换路由、位置参数校验
│   ├── converter.js          # md → docx 编排层：组装 Document
│   ├── blocks.js             # md → docx 块级解析（标题/列表/引用/代码块/表格）
│   ├── inline.js             # md → docx 行内解析（强调/链接/图片）与图片加载
│   ├── header-footer.js      # md → docx 页眉页脚（文字/图片布局/页码/渐变色带）
│   ├── styles.js             # md → docx 样式表、字体对象、间距换算、编号
│   ├── page.js               # md → docx 页面属性（尺寸/边距/垂直对齐/页码格式）
│   ├── docx-to-md.js         # docx → md 编排层
│   ├── docx-import/          # OOXML 预处理、富结构降级与 Markdown 渲染
│   │   ├── extractors/       # 修订、批注、文本框、浮动对象、分节处理
│   │   └── renderers/        # HTML → Markdown 与合并表格回退
│   ├── preset-extract.js     # docx 版式提取为自定义预设（解包 OOXML 逆向映射）
│   ├── args.js               # 参数解析：-- 参数规格表与校验
│   ├── presets.js            # 预设方案（sundy：圣典法律文书）
│   └── options.js            # 默认配置、单位换算、深合并
└── test/
    ├── sample.md             # 测试样例（覆盖全部支持的语法）
    ├── unit-test.mjs         # 单元测试（换算、参数、预设名、修订/批注注释等）
    ├── run-test.mjs          # 端到端校验（九组用例：转换、预设、富结构降级等）
    ├── openxml-validator/    # Microsoft Open XML SDK Schema 验证器
    └── assets/               # 测试图片
```

### 系统要求

- Node.js >= 18

- 仅运行完整测试时：.NET 8 SDK（Microsoft Open XML SDK Schema 验证）

完整测试使用 `npm test`；也可在端到端测试生成样例 docx 后，使用 `npm run test:openxml` 单独执行 Schema 验证。

## 二、本地安装

在本项目根目录执行：

```bash
npm install
npm link
```

`npm link` 会将 `MDTT` 命令注册为全局命令，之后即可在任意目录、任意终端中使用。

## 三、使用方法

### 基本用法

```bash
MDTT <文件名>.md [参数]                  Markdown 转 docx
MDTT <文件名>.docx [参数]                docx 转 Markdown
MDTT <文件名>.docx --save-preset <预设名>  提取 docx 版式为自定义预设并保存
```

文件名可省略后缀：MDTT 会自动匹配同目录同名的 `.md` / `.docx` 文件（两者同时存在时需写明后缀，因为后缀决定转换方向）。

文件名包含空格或特殊符号时，请用英文引号将「文件名+后缀」整体包裹（后缀也必须在引号内），如 `MDTT "我的 文档.md"`。文件名输错、空格未加引号、新旧语法混用等场景均有中文提示引导修正。

### 常用命令

日常工作中高频使用的命令：

```bash
MDTT <文件名>.md                              # 普通文档快速转换（默认格式：与 sundy 排版相同，仅无页眉页脚页码）
MDTT <文件名>                                  # 省略后缀：自动匹配同目录同名的 .md / .docx
MDTT "<含空格的文件名>.md"                      # 文件名含空格/特殊符号时，用英文引号整体包裹（含后缀）
MDTT <文件名>.md -p bottom                    # 在页脚添加居中纯数字的页码
MDTT <文件名>.md -p bottom --page-num-format 第X页/共Y页  # 页脚改为「第X页/共Y页」式页码
MDTT <文件名>.md -o <输出路径>.docx           # 输出到指定路径（如直接存进案件文件夹）
MDTT <文件名>.md --preset sundy               # 出法律文书（最常用：圣典排版，页眉页脚页码齐备）
MDTT <文件名>.md --preset sundy --font 黑体   # 在法律文书预设基础上设置字体
MDTT <文件名>.md --preset sundy --overwrite   # 如有重名文件，直接覆盖旧文件
MDTT <文件名>.md --preset sundy --no-first-page-number  # 法律文书首页（封面）不显示页码
MDTT <文件名>.md -p bottom --page-num-start 2 --no-first-page-number  # 封面不计页码，正文从第 2 页起（合同常用）
MDTT <文件名>.md --header "保密文件"          # 页眉居中显示文字（密级标识、单位名称）
MDTT <文件名>.md --header-left "委托代理合同" --header-right "2026-09"  # 页眉左右分布（左：文件标题，右：日期）
MDTT <文件名>.md --orientation landscape      # 横向页面（宽表格、时间轴、证据清单）
MDTT <文件名>.md -m 2.54,3.18,2.54,3.18       # 四边分别设置页边距（cm，顺序：上,右,下,左）
MDTT <文件名>.md --indent 0                   # 取消首行缩进（英文文档、清单式材料）
MDTT <文件名>.md --line-height 1.5            # 行距调整为 1.5 倍
MDTT <文件名>.md --font-size 小四             # 单独调整正文字号（支持中文字号名）
MDTT <文件名>.docx                            # 收到 Word 文档转回 Markdown 编辑
MDTT <文件名>.docx --overwrite                # 重新转换时覆盖已存在的 md 文件
MDTT <文件名>.docx -o <输出路径>.md           # docx 转 Markdown 并指定输出路径
MDTT <文件名>.docx --save-preset <预设名>     # 将 Word 文档的版式提取为自定义预设（如律所官方模板）
MDTT <文件名>.md --preset <预设名>            # 用提取的自定义预设转换
MDTT --help                                   # 忘记参数时查帮助
```

### 支持的 Markdown 语法

#### 块级元素

| 语法                    | 转换效果                                |
| --------------------- | ----------------------------------- |
| 六级标题（`#` \~ `######`） | Word 内置标题样式 Heading 1-6，字号/对齐/加粗可定制 |
| 有序 / 无序列表（含嵌套）        | Word 原生编号列表，层级缩进                    |
| 引用（`>`，含嵌套）           | 左侧竖线 + 左缩进样式                        |
| 围栏代码块                 | 等宽字体 + 灰色底纹（字体字号可定制）                |
| 表格（含对齐方式）             | Word 表格，表头加粗 + 灰色底纹，支持左/中/右对齐       |
| 分隔线（`---`）            | 段落底部横线                              |

#### 行内元素

| 语法                        | 转换效果                    |
| ------------------------- | ----------------------- |
| `**加粗**`、`*斜体*`、`~~删除线~~` | 对应 Word 字体样式（可分别关闭）     |
| `` `行内代码` ``              | 等宽字体 + 灰色底纹             |
| `[链接](url)`               | 可点击的超链接（蓝色下划线）          |
| `![图片](路径)`               | 嵌入图片，自动缩放至不超过页宽，独立成段时居中 |
| 行尾两个空格（硬换行）               | Word 内换行                |

#### 其他特性

- **中英文混排**：中西文字体分别设置（如仿宋 + Times New Roman）

- **图片容错**：图片缺失时以 `[图片缺失: ...]` 文本占位，不中断转换

- **首页差异化**：`--first-*` / `--no-first-*` 系列参数实现首页页眉页脚独立设置

### docx → Markdown 的富结构降级规则

- **修订**：默认接受修订；在修订后内容后紧接 `<!-- 此处系修订；修订前原文：…… -->`

- **批注**：在被批注内容后紧接 Markdown HTML 注释，记录批注范围和批注正文；范围超过 20 个字符时保留首尾各 10 个字符

- **文本框**：移除文本框外形，将内部段落和表格按锚定位置线性排列

- **浮动图片/图形**：丢弃位置、环绕和层级信息，按普通行内对象提取；没有可提取图片资源的纯矢量外形可能被忽略

- **多栏/多节**：丢弃分栏和分节版式，正文内容按 OOXML 中的源顺序线性排列

- **复杂直接格式**：尽可能保留 mammoth 可识别的语义格式，其余格式允许丢失且不额外警告

## 四、参数列表

### 页面设置

| 参数              | 别名     | 取值                                              | 说明                                                  |
| --------------- | ------ | ----------------------------------------------- | --------------------------------------------------- |
| `--page-size`   | `-s`   | `A4` / `A3` / `A5` / `letter` / `legal` / `宽,高` | 页面尺寸预设或自定义宽高（cm），如 `-s 21,29.7`                     |
| `--orientation` | <br /> | `portrait` / `landscape`                        | 纵向（默认）/横向                                           |
| `--margin`      | `-m`   | `上,右,下,左` 或单值                                   | 页边距（cm），如 `-m 2.54,3.18,2.54,3.18` 或 `-m 2.5`（四边统一） |
| `--v-align`     | <br /> | `top` / `center` / `bottom`                     | 页面内容垂直对齐（重心），默认 top                                 |

### 字体与字号

| 参数               | 别名     | 取值                           | 说明                                                  |
| ---------------- | ------ | ---------------------------- | --------------------------------------------------- |
| `--font`         | `-f`   | 字体名                          | 正文字体（中西文统一），如 `-f 宋体`                               |
| `--font-heading` | <br /> | 字体名                          | 标题字体（默认跟随 `--font`）                                 |
| `--font-code`    | <br /> | 字体名                          | 代码块字体（默认 Consolas）                                  |
| `--font-size`    | <br /> | pt 或中文字号名                    | 正文字号，如 `12` 或 `小四`                                  |
| `--heading-size` | <br /> | `auto` / 单值 / 六级逗号分隔         | 标题字号，如 `22,16,14,14,14,14`                          |
| `--code-size`    | <br /> | pt 或中文字号名                    | 代码块字号                                               |
| `--line-height`  | <br /> | 倍数 / 固定 pt                   | 行距：`auto` 时为倍数（如 `1.5`）；`exact`/`atLeast` 时为固定行高 pt |
| `--line-rule`    | <br /> | `auto` / `exact` / `atLeast` | 行距规则（默认 `auto`），配合 `--line-height` 使用               |

支持的中文字号名：初号、小初、一号、小一、二号、小二、三号、小三、四号、小四、五号、小五、六号、小六。

### 页眉

| 参数                                           | 说明                                     |
| -------------------------------------------- | -------------------------------------- |
| `--header <文本>`                              | 页眉文字（默认居中），文本中的 URL 自动转为超链接，支持 `\n` 多行 |
| `--header-left <文本>` / `--header-right <文本>` | 页眉左右分布（与 `--header` 互斥）                |
| `--header-align <left\|center\|right>`       | 页眉对齐，默认 center                         |
| `--header-font <字体名>`                        | 页眉字体，默认跟随正文                            |
| `--header-size <字号>`                         | 页眉字号，默认 9pt                            |
| `--header-image <路径>`                        | 页眉右侧放置图片（相对 md 文件或绝对路径），文字自动左置         |
| `--header-line`                              | 显示页眉下横线（默认不显示）                         |
| `--first-header <文本>`                        | 首页页眉文字                                 |
| `--no-first-header`                          | 首页不显示页眉                                |

### 页脚

与页眉基本对称：`--footer`、`--footer-left`、`--footer-right`、`--footer-align`、`--footer-font`、`--footer-size`（含页码字号）、`--footer-line`（页脚上横线）、`--first-footer`、`--no-first-footer`。页码位于页脚时使用页脚字体字号，位于页眉（`-p top`）时使用页眉字体字号。

### 页码

| 参数                                       | 别名                                                                                                | 说明           |
| ---------------------------------------- | ------------------------------------------------------------------------------------------------- | ------------ |
| `--page-number <none\|top\|bottom>`      | `-p`                                                                                              | 页码位置，默认 none |
| `--page-num-align <left\|center\|right>` | 页码对齐，默认 center                                                                                    | <br />       |
| `--page-num-format <模板>`                 | 模板中 `X`=当前页码、`Y`=总页数，如 `第X页`、`第X页/共Y页`、`X / Y`；特殊值 `1` / `i` / `I` / `一` 表示数字样式（阿拉伯/小写罗马/大写罗马/中文） | <br />       |
| `--page-num-start <数字>`                  | 起始页码                                                                                              | <br />       |
| `--total-pages`                          | 页码后附加总页数（格式无 `Y` 时追加 `/共Y页`）                                                                      | <br />       |
| `--no-first-page-number`                 | 首页不显示页码                                                                                           | <br />       |

### 正文对齐与间距

| 参数                                       | 说明                                              |
| ---------------------------------------- | ----------------------------------------------- |
| `--align <left\|center\|right\|justify>` | 正文对齐，默认左对齐                                      |
| `--heading-align <left\|center\|right>`  | 标题对齐（各级统一）                                      |
| `--para-spacing <pt>`                    | 段后间距                                            |
| `--heading-spacing <pt>`                 | 标题段前及段后间距                                       |
| `--indent <字符数>`                         | 首行缩进（仅普通正文段落；标题、列表、引用、代码块、表格不缩进），如 `--indent 2` |

### 样式开关

| 参数                    | 说明           |
| --------------------- | ------------ |
| `--no-bold`           | 标题不加粗        |
| `--no-italic`         | 斜体语法按普通文字渲染  |
| `--no-strike`         | 删除线语法按普通文字渲染 |
| `--heading-uppercase` | 标题英文转大写      |

### 输出控制

| 参数                    | 别名                 | 说明                                                       |
| --------------------- | ------------------ | -------------------------------------------------------- |
| `--output <路径>`       | `-o`               | 指定输出文件路径                                                 |
| `--overwrite`         | 覆盖已存在的输出文件（默认报错提示） | <br />                                                   |
| `--preset <预设名>`      | <br />             | 预设方案：内置（见下节）或自定义预设名                                      |
| `--save-preset <预设名>` | <br />             | 将 docx 的版式提取为自定义预设并保存（仅 .docx，不执行转换），重名时需加 `--overwrite` |

## 五、预设方案

| 预设                | 说明                            |
| ----------------- | ----------------------------- |
| `--preset sundy`  | **圣典法律文书**（详见下文完整规范）          |
| `--preset <自定义名>` | `--save-preset` 提取的自定义预设（见下节） |

预设可与单项参数混用，**单项参数覆盖预设中的对应项**：

```bash
MDTT doc.md --preset sundy --font 黑体 --header "保密文件"
```

参数优先级：命令行单项参数 > `--preset` 预设 > 默认值。

### 自定义预设（从 docx 提取版式）

将任意 Word 文档的版式提取为可复用的自定义预设，保存于 `~/.mdtt/presets/`：

```bash
MDTT 模板.docx --save-preset firm     # 提取版式保存为自定义预设 firm
MDTT notes.md --preset firm           # 用自定义预设 firm 转换
MDTT 模板.docx --save-preset firm --overwrite  # 覆盖同名自定义预设
```

**提取范围**：页面（尺寸/方向/页边距/垂直对齐）、正文字体与字号、段落（首行缩进/行距/段后距/对齐）、六级标题（字体/字号/加粗/对齐/间距）、页眉页脚（文字/对齐/字体字号/图片）、页码（位置/格式/对齐/起始页码）。行距支持倍数（auto）、固定值（exact）、最小值（atLeast）三种规则。

**说明**：

- 预设名仅限中英文、数字、下划线、连字符，且不以连字符开头（≤64 字符）

- 页眉图片会提取并落盘到预设旁（复用预设时自动还原）；渐变色带等无法映射的元素自动跳过，并在提取结果中逐项注明

- 同一级标题在正文中存在多种对齐（如一个居中、一个左对齐）时，会明确反馈不一致及具体分布，并保留样式表定义

- 提取后命令行会输出版式摘要，便于核对

- 自定义预设与单项参数可混用，优先级同内置预设

### sundy 预设规范（圣典法律文书）

| 项目   | 设定                                                                                                                                             |
| ---- | ---------------------------------------------------------------------------------------------------------------------------------------------- |
| 纸张   | A4 纵向；上下页边距 2.54cm，左右 3.18cm                                                                                                                   |
| 标题字体 | 中文宋体 / 西文 Times New Roman                                                                                                                      |
| 标题字号 | H1 二号、H2 三号、H3-H6 四号                                                                                                                           |
| 标题对齐 | H1 居中，其余左对齐                                                                                                                                    |
| 标题加粗 | H1-H3 加粗，H4-H6 不加粗                                                                                                                             |
| 标题间距 | 段前 0.5 行，段后 0 行                                                                                                                                |
| 正文字体 | 中文仿宋 / 西文 Times New Roman                                                                                                                      |
| 正文字号 | 四号                                                                                                                                             |
| 正文段落 | 首行缩进 2 字符；段后 0.5 行；行距 1.28 倍                                                                                                                   |
| 页眉   | 距页面顶端 0.85cm；三行左对齐：①圣典律师事务所 ②圣典官网：<https://www.sundylawyer.com/（超链接）③总所地址：南京市建邺区奥体大街68号新城科技园4A栋6楼、7楼；右端放置律所> logo；页眉底端红→橙→金渐变色带（以红为主）；仿宋/Times New Roman 小五 |
| 页脚   | 居中页码「第X页/共Y页」，仿宋/Times New Roman 五号                                                                                                            |
