// 命令行参数解析：将 argv 中以 - 开头的参数解析为配置补丁
// 返回 { patch, output, overwrite, preset, errors }
// patch 与 defaultOptions 同形，交给 cli.js 按优先级合并

import { parseFontSize } from "./options.js";

const ALIGNMENTS = ["left", "center", "right", "justify"];

/** 参数规格表：name → { alias, value(取值) / flag(布尔), apply(补丁写入) } */
const SPECS = [
  // ---- 页面 ----
  {
    name: "--page-size", alias: "-s",
    help: "A4 / A3 / A5 / letter / legal 或自定义宽高如 21,29.7（cm）",
    apply: (v, p) => (p.page = { ...p.page, size: v }),
  },
  {
    name: "--orientation",
    help: "portrait（纵向，默认）/ landscape（横向）",
    validate: (v) => ["portrait", "landscape"].includes(v) || "应为 portrait 或 landscape",
    apply: (v, p) => (p.page = { ...p.page, orientation: v }),
  },
  {
    name: "--margin", alias: "-m",
    help: "上,右,下,左（cm）或单一数值四边统一，如 2.54,3.18,2.54,3.18",
    apply: parseMargin,
  },
  {
    name: "--v-align",
    help: "页面内容垂直对齐：top（默认）/ center / bottom",
    validate: (v) => ["top", "center", "bottom"].includes(v) || "应为 top、center 或 bottom",
    apply: (v, p) => (p.page = { ...p.page, vAlign: v }),
  },
  // ---- 字体与字号 ----
  {
    name: "--font", alias: "-f",
    help: "正文中文字体，如 宋体、仿宋、微软雅黑",
    apply: (v, p) => (p.fonts = { ...p.fonts, body: { eastAsia: v, ascii: v } }),
  },
  {
    name: "--font-heading",
    help: "标题中文字体（默认跟随 --font）",
    apply: (v, p) => (p.fonts = { ...p.fonts, heading: { eastAsia: v, ascii: v } }),
  },
  {
    name: "--font-code",
    help: "代码块字体（默认 Consolas）",
    apply: (v, p) => (p.fonts = { ...p.fonts, code: { eastAsia: v, ascii: v } }),
  },
  {
    name: "--font-size",
    help: "正文字号：数字（pt）或中文字号名（如 四号）",
    apply: (v, p) => (p.sizes = { ...p.sizes, body: parseFontSize(v) }),
  },
  {
    name: "--heading-size",
    help: "标题字号：auto（默认递减）/ 单一字号（六级相同）/ 逗号分隔六级，如 22,16,14,14,14,14",
    apply: parseHeadingSize,
  },
  {
    name: "--code-size",
    help: "代码块字号（pt 或中文字号名）",
    apply: (v, p) => (p.sizes = { ...p.sizes, code: parseFontSize(v) }),
  },
  {
    name: "--line-height",
    help: "行距值：lineRule=auto（默认）时为倍数（如 1.5）；exact/atLeast 时为固定行高（pt）",
    validate: (v) => (Number(v) > 0 || "应为正数"),
    apply: (v, p) => (p.paragraph = { ...p.paragraph, line: Number(v) }),
  },
  {
    name: "--line-rule",
    help: "行距规则：auto（倍数，默认）/ exact（固定行高）/ atLeast（最小行高），配合 --line-height 使用",
    validate: (v) => ["auto", "exact", "atLeast"].includes(v) || "应为 auto、exact 或 atLeast",
    apply: (v, p) => (p.paragraph = { ...p.paragraph, lineRule: v }),
  },
  // ---- 页眉 ----
  { name: "--header", help: "页眉文字（默认居中）", apply: (v, p) => (p.header = { ...p.header, text: v }) },
  { name: "--header-left", help: "页眉左侧文字（与 --header 互斥）", apply: (v, p) => (p.header = { ...p.header, left: v }) },
  { name: "--header-right", help: "页眉右侧文字", apply: (v, p) => (p.header = { ...p.header, right: v }) },
  {
    name: "--header-align",
    help: "页眉对齐：left / center（默认）/ right",
    validate: (v) => ALIGNMENTS.slice(0, 3).includes(v) || "应为 left、center 或 right",
    apply: (v, p) => (p.header = { ...p.header, align: v }),
  },
  {
    name: "--header-font",
    help: "页眉中文字体（默认跟随正文）",
    apply: (v, p) => (p.fonts = { ...p.fonts, header: { eastAsia: v, ascii: v } }),
  },
  {
    name: "--header-size",
    help: "页眉字号（默认 9pt，可用中文字号名如 小五）",
    apply: (v, p) => (p.sizes = { ...p.sizes, header: parseFontSize(v) }),
  },
  { name: "--header-line", flag: true, help: "显示页眉下横线", apply: (p) => (p.header = { ...p.header, line: true }) },
  {
    name: "--header-image",
    help: "页眉右侧图片路径（相对 md 文件或绝对路径），文字自动左置",
    apply: (v, p) => (p.header = { ...p.header, image: { ...(p.header.image || {}), path: v } }),
  },
  { name: "--first-header", help: "首页页眉文字", apply: (v, p) => (p.header = { ...p.header, firstText: v }) },
  { name: "--no-first-header", flag: true, help: "首页不显示页眉", apply: (p) => (p.header = { ...p.header, noFirst: true }) },
  // ---- 页脚 ----
  { name: "--footer", help: "页脚文字（默认居中）", apply: (v, p) => (p.footer = { ...p.footer, text: v }) },
  { name: "--footer-left", help: "页脚左侧文字（与 --footer 互斥）", apply: (v, p) => (p.footer = { ...p.footer, left: v }) },
  { name: "--footer-right", help: "页脚右侧文字", apply: (v, p) => (p.footer = { ...p.footer, right: v }) },
  {
    name: "--footer-align",
    help: "页脚对齐：left / center（默认）/ right",
    validate: (v) => ALIGNMENTS.slice(0, 3).includes(v) || "应为 left、center 或 right",
    apply: (v, p) => (p.footer = { ...p.footer, align: v }),
  },
  {
    name: "--footer-font",
    help: "页脚中文字体（默认跟随正文）",
    apply: (v, p) => (p.fonts = { ...p.fonts, footer: { eastAsia: v, ascii: v } }),
  },
  {
    name: "--footer-size",
    help: "页脚字号（含页码，默认 9pt，可用中文字号名）",
    apply: (v, p) => (p.sizes = { ...p.sizes, footer: parseFontSize(v) }),
  },
  { name: "--footer-line", flag: true, help: "显示页脚上横线", apply: (p) => (p.footer = { ...p.footer, line: true }) },
  { name: "--first-footer", help: "首页页脚文字", apply: (v, p) => (p.footer = { ...p.footer, firstText: v }) },
  { name: "--no-first-footer", flag: true, help: "首页不显示页脚", apply: (p) => (p.footer = { ...p.footer, noFirst: true }) },
  // ---- 页码 ----
  {
    name: "--page-number", alias: "-p",
    help: "页码位置：none（默认）/ top / bottom",
    validate: (v) => ["none", "top", "bottom"].includes(v) || "应为 none、top 或 bottom",
    apply: (v, p) => (p.pageNumber = { ...p.pageNumber, pos: v }),
  },
  {
    name: "--page-num-align",
    help: "页码对齐：left / center（默认）/ right",
    validate: (v) => ALIGNMENTS.slice(0, 3).includes(v) || "应为 left、center 或 right",
    apply: (v, p) => (p.pageNumber = { ...p.pageNumber, align: v }),
  },
  {
    name: "--page-num-format",
    help: "页码格式模板：X=当前页码，Y=总页数，如 第X页、第X页/共Y页；或数字样式 1 / i / I / 一",
    apply: (v, p) => (p.pageNumber = { ...p.pageNumber, format: v }),
  },
  {
    name: "--page-num-start",
    help: "起始页码（数字）",
    validate: (v) => (Number.isInteger(Number(v)) && Number(v) >= 1 || "应为正整数"),
    apply: (v, p) => (p.pageNumber = { ...p.pageNumber, start: Number(v) }),
  },
  {
    name: "--total-pages", flag: true,
    help: "页码后附加总页数（如格式无 Y 则追加 /共Y页）",
    apply: (p) => (p.pageNumber = { ...p.pageNumber, format: withTotal(p.pageNumber?.format ?? "X") }),
  },
  { name: "--no-first-page-number", flag: true, help: "首页不显示页码", apply: (p) => (p.pageNumber = { ...p.pageNumber, showOnFirstPage: false }) },
  // ---- 对齐与间距 ----
  {
    name: "--align",
    help: "正文对齐：left（默认）/ center / right / justify",
    validate: (v) => ALIGNMENTS.includes(v) || "应为 left、center、right 或 justify",
    apply: (v, p) => (p.paragraph = { ...p.paragraph, align: v }),
  },
  {
    name: "--heading-align",
    help: "标题对齐（各级统一）：left（默认）/ center / right",
    validate: (v) => ALIGNMENTS.slice(0, 3).includes(v) || "应为 left、center 或 right",
    apply: (v, p) => (p.heading = { ...p.heading, align: Array(6).fill(v) }),
  },
  {
    name: "--para-spacing",
    help: "段后间距（pt，如 8）",
    apply: (v, p) => (p.paragraph = { ...p.paragraph, afterPt: Number(v), afterLines: null }),
  },
  {
    name: "--heading-spacing",
    help: "标题段前及段后间距（pt，如 12）",
    apply: (v, p) => (p.heading = { ...p.heading, spacing: { beforePt: Number(v), afterPt: Number(v) } }),
  },
  {
    name: "--indent",
    help: "首行缩进字符数（仅正文段落），如 2",
    validate: (v) => (Number(v) >= 0 || "应为非负数字"),
    apply: (v, p) => (p.paragraph = { ...p.paragraph, firstLineChars: Number(v) }),
  },
  // ---- 样式开关 ----
  { name: "--no-bold", flag: true, help: "标题不加粗", apply: (p) => (p.heading = { ...p.heading, bold: Array(6).fill(false) }) },
  { name: "--no-italic", flag: true, help: "斜体按普通文字渲染", apply: (p) => (p.styles = { ...p.styles, italic: false }) },
  { name: "--no-strike", flag: true, help: "删除线按普通文字渲染", apply: (p) => (p.styles = { ...p.styles, strike: false }) },
  { name: "--heading-uppercase", flag: true, help: "标题英文转大写", apply: (p) => (p.styles = { ...p.styles, headingUppercase: true }) },
  // ---- 输出控制 ----
  { name: "--output", alias: "-o", help: "输出文件路径", cli: "output" },
  { name: "--overwrite", flag: true, help: "覆盖已存在的输出文件", cli: "overwrite" },
  { name: "--preset", help: "预设方案：sundy（圣典法律文书），或 ~/.mddtt/presets/ 中的自定义预设", cli: "preset" },
  {
    name: "--save-preset",
    help: "将 docx 的版式提取为自定义预设并保存到 ~/.mddtt/presets/（仅 .docx，不执行转换）",
    cli: "savePreset",
  },
];

function withTotal(format) {
  return format.includes("Y") ? format : `${format}/共Y页`;
}

/** 页边距取值解析：单值四边统一 / 四值（上,右,下,左），非法值抛错（导出供单元测试） */
export function parseMargin(v, patch) {
  const parts = v.split(",").map((x) => Number(x.trim()));
  if (parts.some((n) => Number.isNaN(n) || n < 0)) throw new Error(`页边距格式错误：「${v}」应为逗号分隔的非负数字（cm）`);
  let [top, right, bottom, left] = parts;
  if (parts.length === 1) right = bottom = left = top;
  if (parts.length !== 1 && parts.length !== 4) throw new Error(`页边距应为 1 个或 4 个数值：「${v}」`);
  patch.page = { ...patch.page, margin: { top, right, bottom, left } };
}

/** 标题字号取值解析：auto / 单值 / 六级逗号分隔（导出供单元测试） */
export function parseHeadingSize(v, patch) {
  if (v === "auto") {
    patch.heading = patch.heading || {};
    return;
  }
  const parts = v.split(",").map((x) => parseFontSize(x));
  let sizes6;
  if (parts.length === 1) sizes6 = Array(6).fill(parts[0]);
  else if (parts.length === 6) sizes6 = parts;
  else throw new Error(`标题字号应为单一值或逗号分隔的六级值：「${v}」`);
  patch.sizes = { ...patch.sizes, heading: sizes6 };
}

/** 解析 argv（已剔除位置参数）为 { patch, output, overwrite, preset, errors } */
export function parseArgs(argv) {
  const patch = { page: {}, fonts: {}, sizes: {}, heading: {}, paragraph: {}, header: {}, footer: {}, pageNumber: {}, styles: {} };
  const result = { patch, output: null, overwrite: false, preset: null, savePreset: null, errors: [] };
  const specByName = new Map();
  for (const s of SPECS) {
    specByName.set(s.name, s);
    if (s.alias) specByName.set(s.alias, s);
  }

  let i = 0;
  while (i < argv.length) {
    const token = argv[i];
    if (!token.startsWith("-")) {
      result.errors.push(`多余的位置参数「${token}」`);
      i++;
      continue;
    }
    const eq = token.indexOf("=");
    const name = eq === -1 ? token : token.slice(0, eq);
    const spec = specByName.get(name);
    if (!spec) {
      result.errors.push(`未知参数「${name}」（使用 --help 查看全部参数）`);
      i++;
      continue;
    }
    let value = eq === -1 ? null : token.slice(eq + 1);
    if (!spec.flag && value === null) {
      if (i + 1 >= argv.length) {
        result.errors.push(`参数 ${name} 缺少取值`);
        i++;
        continue;
      }
      value = argv[++i];
    }
    try {
      if (spec.cli) {
        if (spec.cli === "output") result.output = value;
        else if (spec.cli === "overwrite") result.overwrite = true;
        else if (spec.cli === "preset") result.preset = value;
        else if (spec.cli === "savePreset") result.savePreset = value;
      } else if (spec.flag) {
        spec.apply(patch);
      } else {
        if (spec.validate) {
          const err = spec.validate(value);
          if (err !== true && err !== undefined) throw new Error(`${name} ${err}`);
        }
        spec.apply(value, patch);
      }
    } catch (e) {
      result.errors.push(e.message);
    }
    i++;
  }

  // 互斥校验
  const h = patch.header || {};
  if (h.text != null && (h.left != null || h.right != null)) {
    result.errors.push("--header 与 --header-left/--header-right 互斥，请只选其一");
  }
  const f = patch.footer || {};
  if (f.text != null && (f.left != null || f.right != null)) {
    result.errors.push("--footer 与 --footer-left/--footer-right 互斥，请只选其一");
  }
  return result;
}

/** 生成帮助文本中的参数说明 */
export function argsHelpText() {
  return SPECS.map((s) => {
    const alias = s.alias ? `, ${s.alias}` : "";
    const value = s.flag ? "" : " <值>";
    return `  ${s.name}${alias}${value}\n      ${s.help}`;
  }).join("\n");
}
