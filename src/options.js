// 规范化配置结构与单位换算
// 尺寸类：页面/页边距用 cm，字号用 pt（内部换算为 twip / 半磅）
// 预设（presets.js）与命令行参数（args.js）均生成与本结构同形的补丁对象，
// 经 mergeOptions 深合并后交给 converter.js 使用。

// 中文字号名 → pt
export const CHINESE_FONT_SIZES = {
  初号: 42, 小初: 36, 一号: 26, 小一: 24, 二号: 22, 小二: 18,
  三号: 16, 小三: 15, 四号: 14, 小四: 12, 五号: 10.5, 小五: 9,
  六号: 7.5, 小六: 6.5,
};

// 页面尺寸预设（twip）
export const PAGE_SIZES = {
  A4: { width: 11906, height: 16838 },
  A3: { width: 16838, height: 23811 },
  A5: { width: 8391, height: 11906 },
  letter: { width: 12240, height: 15840 },
  legal: { width: 12240, height: 20160 },
};

// cm → twip
export const cmToTwip = (cm) => Math.round(cm * 566.9291339);
// pt → twip
export const ptToTwip = (pt) => Math.round(pt * 20);
// pt → 半磅（docx 库的 size 单位）
export const ptToHalfPoint = (pt) => Math.round(pt * 2);

/** 解析字号：数字视为 pt，中文字号名自动映射 */
export function parseFontSize(value) {
  if (typeof value === "number") return value;
  const num = Number(value);
  if (!Number.isNaN(num) && value.trim() !== "") return num;
  const cn = CHINESE_FONT_SIZES[value.trim()];
  if (cn) return cn;
  throw new Error(`无法识别的字号「${value}」（应为数字 pt 或中文字号名如“四号”）`);
}

/** 默认配置（不带任何参数与预设时的行为） */
export const defaultOptions = {
  page: {
    size: "A4",            // PAGE_SIZES 键名，或 "W,H"（cm）
    orientation: "portrait",
    margin: { top: 2.54, right: 3.18, bottom: 2.54, left: 3.18 }, // cm
    vAlign: "top",         // 页面垂直对齐（重心）：top / center / bottom
  },
  fonts: {
    body: { eastAsia: "仿宋", ascii: "Times New Roman" },
    heading: { eastAsia: "宋体", ascii: "Times New Roman" },
    code: { eastAsia: "仿宋", ascii: "Consolas" },
    header: { eastAsia: "仿宋", ascii: "Times New Roman" },
    footer: { eastAsia: "仿宋", ascii: "Times New Roman" },
  },
  sizes: {
    body: 14,              // pt（四号）
    heading: [22, 16, 14, 14, 14, 14], // pt，H1-H6（二号/三号/四号×4）
    code: 9,
    header: 9,             // 页眉字号（小五）
    footer: 10.5,          // 页脚字号（五号）
  },
  heading: {
    bold: [true, true, true, false, false, false],
    align: ["center", null, null, null, null, null], // null = 左对齐；H1 居中
    spacing: { beforeLines: 0.5, afterLines: 0 },    // 段前 0.5 行，段后 0（lines 优先于 pt）
  },
  paragraph: {
    align: null,           // null = 默认（左）
    firstLineChars: 2,     // 首行缩进字符数（仅普通正文段落）
    afterPt: 8,            // 段后间距（afterLines 优先于本值）
    afterLines: 1.5,       // 空行分段后的段后距，按正文行高计 1.5 行
    line: 1.28,            // 行距：lineRule=auto 时为倍数；exact/atLeast 时为固定行高（pt）
    lineRule: "auto",      // auto（倍数）/ exact（固定）/ atLeast（最小值）
  },
  header: {
    text: null, left: null, right: null,
    align: "center",
    firstText: null,       // 首页页眉文字；null=与默认页相同
    noFirst: false,
    line: false,           // 页眉下横线
    image: null,           // { path, heightCm } 页眉右侧图片；text 支持多行（\n 分隔）
    gradientLine: null,    // { stops: [[位置, "RRGGBB"], ...], heightCm } 页眉底端渐变色带
  },
  footer: {
    text: null, left: null, right: null,
    align: "center",
    firstText: null,
    noFirst: false,
    line: false,
  },
  pageNumber: {
    pos: "none",           // none / top / bottom
    align: "center",
    format: "X",           // 模板：X=当前页码，Y=总页数；特殊值 1/i/I/一 表示数字样式
    start: null,           // 起始页码
    showOnFirstPage: true,
  },
  styles: {
    italic: true,
    strike: true,
    headingUppercase: false,
  },
  markdown: {
    // 换行识别：true 时单个换行符即段内换行（配合双空格硬换行），false 时按 Markdown
    // 软换行处理（合并为同一段中的空格）。中文文档多以换行分行，故默认开启。
    breaks: true,
  },
};

/** 深合并：plain object 递归合并，其余类型（含数组）直接覆盖 */
export function mergeOptions(base, ...patches) {
  return patches.reduce(mergeTwo, base);
}

function mergeTwo(target, patch) {
  if (!patch) return target;
  const out = { ...target };
  for (const [key, value] of Object.entries(patch)) {
    if (value === undefined) continue;
    const prev = out[key];
    out[key] =
      isPlainObject(prev) && isPlainObject(value) ? mergeTwo(prev, value) : value;
  }
  return out;
}

function isPlainObject(v) {
  return v !== null && typeof v === "object" && !Array.isArray(v);
}
