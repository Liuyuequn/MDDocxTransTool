# MDDTT — Markdown 与 docx 互转命令行工具

## 触发条件

当用户要求将 Markdown（.md）文件转换为 Word（.docx）文档，或将 Word（.docx）文件转换为 Markdown 时，使用此 Skill。

## 工具位置

全局命令：`mddtt`（通过 `npm link` 注册，可在任意目录下使用）

## 基本用法

```bash
mddtt <文件名>.md [参数]                  Markdown 转 docx
mddtt <文件名>.docx [参数]                docx 转 Markdown
mddtt <文件名>.docx --save-preset <预设名>  提取 docx 版式为自定义预设并保存
```

输出文件默认与源文件同目录、同名（仅扩展名变化）。文件名可省略后缀，将自动匹配同目录同名的 `.md` / `.docx`（两者同时存在时需写明后缀）。

转化的文档是法律相关文档时请添加预设参数：`mddtt notes.md --preset sundy`。

## 常用示例

### Markdown → docx

```bash
mddtt notes.md                              # 默认格式（与 sundy 排版相同：仿宋四号、首行缩进两字符、行距1.28、H1居中；仅无页眉页脚页码）
mddtt notes.md --preset sundy               # 圣典法律文书预设（圣典律师事务所专用排版）
mddtt notes.md -o output/result.docx        # 指定输出路径
mddtt notes.md --overwrite                  # 覆盖已有文件
mddtt notes.md --preset sundy --overwrite   # 使用法律文书预设并覆盖
```

### docx → Markdown

```bash
mddtt report.docx                           # 转为同名 .md 文件
mddtt report.docx -o output.md              # 指定输出路径
mddtt report.docx --overwrite               # 覆盖已存在的文件
```

### 版式提取（docx → 自定义预设）

```bash
mddtt 模板.docx --save-preset firm          # 提取 Word 模板版式保存为自定义预设 firm
mddtt notes.md --preset firm                # 用自定义预设 firm 转换
mddtt 模板.docx --save-preset firm --overwrite  # 覆盖同名自定义预设
```

## 预设方案（仅 md → docx 时有效）

| 预设                | 说明                                                             |
| ----------------- | -------------------------------------------------------------- |
| `--preset sundy`  | 圣典法律文书：A4、宋体标题、仿宋正文四号、首行缩进2字符、页眉（圣典律师事务所 logo + 渐变色带）、页脚页码（五号） |
| `--preset <自定义名>` | 用 `--save-preset` 从 docx 提取的自定义预设（存于 `~/.mddtt/presets/`）       |

自定义预设提取范围：页面（尺寸/方向/页边距）、正文字体字号、段落（首行缩进/行距/段后距/对齐）、六级标题样式、页眉页脚文字与字体字号（含图片）、页码（位置/格式/起始页）；行距支持倍数（auto）/固定值（exact）/最小值（atLeast）；同一级标题存在多种对齐时会明确反馈不一致；渐变色带等无法映射的元素自动跳过并注明。

## 常用参数速查（仅 md → docx 时有效）

| 参数                                 | 说明                                              |
| ---------------------------------- | ----------------------------------------------- |
| `-s A4 / A3 / A5 / letter / legal` | 页面尺寸（也可自定义宽,高，单位 cm）                            |
| `-m 上,右,下,左`                       | 页边距（cm）                                         |
| `-f 宋体`                            | 正文字体                                            |
| `--font-size 四号`                   | 正文字号（支持中文字号名）                                   |
| `--line-height 1.5`                | 行距：`auto` 为倍数；`exact`/`atLeast` 为固定行高 pt        |
| `--line-rule exact`                | 行距规则：auto / exact / atLeast（配合 `--line-height`） |
| `--indent 2`                       | 首行缩进字符数                                         |
| `-p bottom`                        | 页码位置（top/bottom/none）                           |
| `--page-num-format 第X页/共Y页`        | 页码格式                                            |
| `-o 路径`                            | 指定输出文件路径                                        |
| `--save-preset <预设名>`              | 将 docx 版式提取为自定义预设（仅 .docx）                      |
| `--overwrite`                      | 覆盖已有输出文件 / 同名自定义预设                              |
| `--help`                           | 查看全部参数                                          |

## 注意事项

- 输出文件已存在时必须加 `--overwrite`，否则报错并退出

- 文件名含空格或特殊符号时，必须用英文引号将「文件名+后缀」整体包裹（如 `mddtt "我的 文档.md"`），否则参数会被终端拆散或丢弃

- 支持的中文字号名：初号、小初、一号、小一、二号、小二、三号、小三、四号、小四、五号、小五、六号、小六

- 图片缺失时以占位文本提示，不中断转换

- docx → md 会保留标题、列表（含嵌套）、表格、加粗/斜体/删除线、超链接；图片提取到输出目录的 `MDPictures/` 文件夹（hash 命名，相对路径引用）；**含合并单元格的表格以 HTML** **`<table>`** **形式嵌入**；修订默认接受并在修订后文后用 Markdown 注释记录原文；批注在被批注内容后用 Markdown 注释记录范围与正文；文本框、浮动对象、多栏和多节内容降级后线性排列；代码块和引用因 Word 格式限制无法自动识别

- 完整参数列表见 `mddtt --help`
