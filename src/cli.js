#!/usr/bin/env node
// mddtt 命令行入口
// 格式：mddtt <文件>.md [参数]                       → md 转 docx
//       mddtt <文件>.docx [参数]                     → docx 转 md
//       mddtt <文件>.docx --save-preset <预设名>      → 提取 docx 版式为自定义预设
// 参数优先级：命令行单项参数 > 预设（--preset）> 默认值（options.js）

import fs from "node:fs";
import path from "node:path";
import { convertMarkdownFile } from "./converter.js";
import { convertDocxFile } from "./docx-to-md.js";
import { defaultOptions, mergeOptions } from "./options.js";
import { presets } from "./presets.js";
import { parseArgs, argsHelpText } from "./args.js";
import {
  extractAndSavePreset,
  loadCustomPreset,
  listCustomPresets,
  validPresetName,
  presetSummaryLines,
} from "./preset-extract.js";

const USAGE = `MDDTT - Markdown Docx Trans Tool

用法:
  mddtt <文件名>.md [参数]                  Markdown 转 docx
  mddtt <文件名>.docx [参数]                docx 转 Markdown
  mddtt <文件名>.docx --save-preset <预设名>  提取 docx 版式为自定义预设并保存

示例:
  mddtt notes.md                            使用默认格式转为 docx
  mddtt notes.md --preset sundy             使用圣典法律文书预设转为 docx
  mddtt report.docx                         docx 转为 Markdown
  mddtt report.docx -o output.md            docx 转为 Markdown 并指定输出路径
  mddtt report.docx --save-preset firm      将 report.docx 的版式保存为自定义预设 firm
  mddtt notes.md --preset firm              用自定义预设 firm 转换

预设方案（仅 md → docx 时有效）:
  sundy    圣典法律文书（A4、宋体标题/仿宋正文、四号、首行缩进两字符、页眉页脚）
  自定义   --save-preset 提取保存于 ~/.mddtt/presets/，用 --preset <名> 调用

参数列表（--save-preset 适用于 .docx，其余仅 md → docx 时有效）:
${argsHelpText()}

说明:
  - 尺寸参数（页边距、自定义页面宽高）为纯数字，单位 cm
  - 字号参数接受数字（pt）或中文字号名（如 四号、小五）
  - 页码格式模板中 X 表示当前页码、Y 表示总页数
  - 文件名可省略后缀：自动匹配同目录同名的 .md / .docx（两者同时存在时须写明后缀）
  - --save-preset 提取页面/字体/字号/段落/标题/页眉页脚/页码为预设；页眉图片、渐变色带等无法映射的元素跳过并在结果中注明`;

const argv = process.argv.slice(2);

// 引号使用提示：文件名含空格或特殊符号时，须用英文引号将「文件名+后缀」整体包裹
// （半包裹如 "文件.md".docx 会被 PowerShell 解析为属性访问，整个参数被丢弃，程序收到的将是空参数）
const QUOTE_TIP = '提示: 文件名包含空格或特殊符号时，请用英文引号将「文件名+后缀」整体包裹（后缀也必须在引号内），例如: mddtt "我的 文档.md"';

if (argv.includes("--help") || argv.includes("-h")) {
  console.log(USAGE);
  process.exit(0);
}
if (argv.length === 0) {
  console.error("错误: 请指定要转换的文件");
  console.error(QUOTE_TIP);
  console.log(USAGE);
  process.exit(1);
}

// 位置参数必须在最前，遇到第一个 - 开头的 token 后其余全部视为参数项
const firstOption = argv.findIndex((t) => t.startsWith("-"));
const positional = firstOption === -1 ? argv : argv.slice(0, firstOption);
const rest = firstOption === -1 ? [] : argv.slice(firstOption);

if (positional.length < 1) {
  console.error("错误: 请指定要转换的文件");
  console.error(QUOTE_TIP);
  console.log(USAGE);
  process.exit(1);
}

// 旧语法提示：mddtt <文件>.docx to md → 扩展名即转换方向，无需 to md
if (positional.length === 3 && positional[1] === "to" && ["md", "docx"].includes(positional[2].toLowerCase())) {
  console.error(`错误: 已不再需要「${positional.slice(1).join(" ")}」，扩展名即转换方向，请直接使用: mddtt ${positional[0]}`);
  process.exit(1);
}

let inputPath = path.resolve(positional[0]);

// ============ 后缀自动匹配：未写后缀且文件不存在时，尝试补 .md / .docx ============
// 后缀即转换方向：两者同时存在时无法自动决定，要求用户写明
if (!fs.existsSync(inputPath) && path.extname(inputPath) === "") {
  const candidates = [".md", ".docx"].filter((e) => fs.existsSync(inputPath + e));
  if (candidates.length === 1) {
    inputPath += candidates[0];
    console.log(`提示: 已自动匹配 ${positional[0]}${candidates[0]}`);
  } else if (candidates.length === 2) {
    console.error(`错误: ${positional[0]}.md 与 ${positional[0]}.docx 同时存在，请写明后缀以指定转换方向（.md 转 docx、.docx 转 Markdown）`);
    process.exit(1);
  }
}

if (!fs.existsSync(inputPath)) {
  console.error(`错误: 找不到文件 ${inputPath}`);
  if (positional.length > 1) {
    // 文件名含空格未加引号时，会被终端拆成多段——重组原始输入并给出正确写法
    console.error(`提示: 文件名疑似包含空格但未加引号，被终端拆成了多段，请用英文引号将「文件名+后缀」整体包裹，如: mddtt "${positional.join(" ")}"`);
  } else {
    console.error("提示: 请检查文件名和路径是否输入正确");
  }
  process.exit(1);
}

const ext = path.extname(inputPath).toLowerCase();

// ============ 位置参数校验：扩展名即转换方向（.md→docx、.docx→md），只允许一个位置参数 ============
if (ext !== ".md" && ext !== ".docx") {
  if (ext === "") {
    console.error("错误: 请写明文件后缀（.md 转 docx、.docx 转 Markdown）");
  } else {
    console.error(`错误: 不支持的文件类型「${ext}」，仅支持 .md（转 docx）与 .docx（转 Markdown）`);
  }
  process.exit(1);
}
if (positional.length !== 1) {
  console.error(`错误: 无法识别的参数「${positional.slice(1).join(" ")}」，用法：mddtt <文件名>${ext} [选项参数]`);
  console.error(QUOTE_TIP);
  process.exit(1);
}

// 解析参数（位置参数已单独取出，参数项从 rest 中解析）
const parsed = parseArgs(rest);
if (parsed.errors.length) {
  for (const err of parsed.errors) console.error(`错误: ${err}`);
  process.exit(1);
}

// ============ --save-preset：从 docx 提取版式为自定义预设（动作，不执行转换） ============
if (parsed.savePreset != null) {
  if (ext !== ".docx") {
    console.error("错误: --save-preset 仅支持 .docx 文件（从 Word 文档提取版式）");
    process.exit(1);
  }
  try {
    const { file, options, notes } = await extractAndSavePreset(inputPath, parsed.savePreset, {
      overwrite: parsed.overwrite,
    });
    console.log(`预设已保存: ${parsed.savePreset}`);
    console.log(`位置: ${file}`);
    console.log("提取结果:");
    for (const line of presetSummaryLines(options)) console.log(`  ${line}`);
    if (notes.length) {
      console.warn("注意:");
      for (const n of notes) console.warn(`  - ${n}`);
    }
    console.log(`复用方式: mddtt <文件名>.md --preset ${parsed.savePreset}`);
  } catch (e) {
    console.error(`提取失败: ${e.message}`);
    process.exit(1);
  }
  process.exit(0);
}

if (ext === ".docx") {
  // ============ docx → Markdown 模式 ============
  const outputPath = parsed.output
    ? path.resolve(parsed.output)
    : inputPath.replace(/\.docx$/i, ".md");

  if (fs.existsSync(outputPath) && !parsed.overwrite) {
    console.error(`错误: 输出文件已存在 ${outputPath}，如需覆盖请加 --overwrite`);
    process.exit(1);
  }

  try {
    const out = await convertDocxFile(inputPath, outputPath);
    console.log(`转换完成: ${out}`);
  } catch (err) {
    console.error(`转换失败: ${err.message}`);
    process.exit(1);
  }
} else {
  // ============ Markdown → docx 模式 ============
  // 预设解析：先内置预设，再 ~/.mddtt/presets/ 自定义预设；单项参数 > 预设 > 默认值
  let presetPatch = null;
  if (parsed.preset != null) {
    if (presets[parsed.preset]) {
      presetPatch = presets[parsed.preset];
    } else if (validPresetName(parsed.preset)) {
      try {
        presetPatch = loadCustomPreset(parsed.preset);
      } catch (e) {
        console.error(`错误: ${e.message}`);
        process.exit(1);
      }
    }
    if (!presetPatch) {
      const custom = listCustomPresets();
      console.error(
        `错误: 未知预设「${parsed.preset}」，内置：${Object.keys(presets).join(" / ")}`
        + (custom.length ? `；自定义：${custom.join(" / ")}（位于 ~/.mddtt/presets/）` : "")
      );
      process.exit(1);
    }
  }
  const opts = mergeOptions(defaultOptions, presetPatch ?? {}, parsed.patch);

  const outputPath = parsed.output
    ? path.resolve(parsed.output)
    : inputPath.replace(/\.md$/i, ".docx");

  if (fs.existsSync(outputPath) && !parsed.overwrite) {
    console.error(`错误: 输出文件已存在 ${outputPath}，如需覆盖请加 --overwrite`);
    process.exit(1);
  }

  try {
    await convertMarkdownFile(inputPath, outputPath, opts);
    console.log(`转换完成: ${outputPath}`);
  } catch (err) {
    console.error(`转换失败: ${err.message}`);
    process.exit(1);
  }
}
