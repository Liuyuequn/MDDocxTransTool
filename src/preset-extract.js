// 从 docx 提取版式为 MDDTT 预设：解包 OOXML 原始 XML，逆向映射为与 presets.js 同形的补丁
// 入口（cli.js 路由）：mddtt <文件>.docx --save-preset <预设名>
// 自定义预设保存于 ~/.mddtt/presets/<名>.json；MDDTT_HOME 环境变量可重定向（供测试隔离）

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";
import { xml2js } from "xml-js";
import { PAGE_SIZES } from "./options.js";

const TWIP_PER_CM = 566.9291339;

// pgNumType w:fmt → MDDTT 页码格式特殊值（纯数字样式）
const NUM_STYLE_FROM = {
  lowerRoman: "i",
  upperRoman: "I",
  chineseCounting: "一",
  chineseCountingThousand: "一",
  chineseLegalSimplified: "一",
};

// ================= xml-js 导航辅助 =================

function parseXml(xml) {
  return xml2js(xml, { compact: false });
}
function rootEl(doc) {
  return (doc.elements || []).find((e) => e.type === "element") ?? null;
}
function kids(el, name) {
  if (!el || !Array.isArray(el.elements)) return [];
  return el.elements.filter((e) => e.type === "element" && e.name === name);
}
function kid(el, name) {
  return kids(el, name)[0] ?? null;
}
function lastKid(el, name) {
  const all = kids(el, name);
  return all[all.length - 1] ?? null;
}
function attr(el, key) {
  return el?.attributes?.[key] ?? null;
}
function deepAll(el, name, out = []) {
  for (const c of el?.elements || []) {
    if (c.type !== "element") continue;
    if (c.name === name) out.push(c);
    deepAll(c, name, out);
  }
  return out;
}
function textOf(el) {
  return (el?.elements || []).filter((e) => e.type === "text").map((e) => e.text).join("");
}
const round2 = (v) => Math.round(v * 100) / 100;
const round1 = (v) => Math.round(v * 10) / 10;

// ================= rPr / pPr 读取 =================

/** rPr → 字体（中文 eastAsia / 西文 ascii） */
function fontFromRPr(rPr) {
  const rf = kid(rPr, "w:rFonts");
  if (!rf) return null;
  const eastAsia = attr(rf, "w:eastAsia");
  const ascii = attr(rf, "w:ascii") ?? attr(rf, "w:hAnsi");
  if (!eastAsia && !ascii) return null;
  return { eastAsia: eastAsia ?? ascii, ascii: ascii ?? eastAsia };
}
/** rPr → 字号（半磅值 → pt） */
function halfPtFromRPr(rPr) {
  const sz = attr(kid(rPr, "w:sz"), "w:val");
  return sz != null ? Number(sz) / 2 : null;
}
/** rPr → 是否加粗（w:val 为 false/0/none 视为不加粗） */
function isBold(rPr) {
  const b = kid(rPr, "w:b");
  if (!b) return false;
  return !["false", "0", "none"].includes(attr(b, "w:val"));
}
/** pPr → 间距（before/after/line twip 与 lineRule） */
function spacingOf(pPr) {
  const sp = kid(pPr, "w:spacing");
  if (!sp) return null;
  const num = (k) => (attr(sp, `w:${k}`) != null ? Number(attr(sp, `w:${k}`)) : null);
  return { before: num("before"), after: num("after"), line: num("line"), lineRule: attr(sp, "w:lineRule") };
}

// ================= 自定义预设存取 =================
// 存取函数均接受可选 dir 参数（预设目录），默认 ~/.mddtt/presets/。
// 命令行沿用默认值；VS Code 插件传入自己的存储目录（用户级 / 工作区级）。

/** 命令行默认的自定义预设目录：MDDTT_HOME 可重定向（测试用），默认 ~/.mddtt/presets/ */
export function presetsDir() {
  return process.env.MDDTT_HOME
    ? path.join(process.env.MDDTT_HOME, "presets")
    : path.join(os.homedir(), ".mddtt", "presets");
}

/** 预设名校验：中英文/数字/下划线/连字符，不以连字符开头，≤64 字符 */
export function validPresetName(name) {
  return /^[A-Za-z0-9_\u4e00-\u9fff][\w\u4e00-\u9fff-]{0,63}$/.test(String(name ?? ""));
}

export function customPresetPath(name, dir = presetsDir()) {
  return path.join(dir, `${name}.json`);
}

/** 列出自定义预设名（按名称排序）；目录不存在返回空数组 */
export function listCustomPresets(dir = presetsDir()) {
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir).filter((f) => f.endsWith(".json")).map((f) => f.slice(0, -5)).sort();
}

/** 读取自定义预设的 options 补丁；不存在返回 null，文件损坏抛错 */
export function loadCustomPreset(name, dir = presetsDir()) {
  const file = customPresetPath(name, dir);
  if (!fs.existsSync(file)) return null;
  let data;
  try {
    data = JSON.parse(fs.readFileSync(file, "utf-8"));
  } catch (e) {
    throw new Error(`自定义预设文件损坏「${file}」: ${e.message}`);
  }
  if (!data || typeof data !== "object" || !data.options || typeof data.options !== "object") {
    throw new Error(`自定义预设文件格式不正确「${file}」（缺少 options 字段）`);
  }
  return data.options;
}

// ================= 提取核心 =================

/**
 * 从 docx 二进制内容提取版式，返回 { options, notes, images }：
 * options 为与 presets.js 条目同形的补丁（经 mergeOptions 深合并后生效）
 * notes 为无法映射项的中文说明
 * images 为页眉/页脚图片的二进制（保存预设时落盘，见 savePresetFromExtraction）
 * 接受 Buffer / Uint8Array，供命令行之外的宿主（如 VS Code 插件）直接复用。
 */
export async function extractPresetOptionsFromBuffer(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const docXml = await zip.file("word/document.xml")?.async("string");
  if (!docXml) throw new Error("不是有效的 docx 文件（缺少 word/document.xml）");
  const stylesXml = (await zip.file("word/styles.xml")?.async("string")) ?? "";
  const relsXml = (await zip.file("word/_rels/document.xml.rels")?.async("string")) ?? "";
  const rels = parseRels(relsXml);

  const docRoot = rootEl(parseXml(docXml));
  const body = kid(docRoot, "w:body");
  const sectPr = body ? lastKid(body, "w:sectPr") : null;
  const notes = [];
  if (!sectPr) notes.push("未找到页面设置（sectPr），页面参数未提取");

  const opts = {};

  // ---- 页面：尺寸 / 方向 / 边距 / 垂直对齐 ----
  if (sectPr) {
    const pgSz = kid(sectPr, "w:pgSz");
    if (pgSz) Object.assign(opts, { page: { ...opts.page, ...extractPageSize(pgSz) } });
    const pgMar = kid(sectPr, "w:pgMar");
    if (pgMar) {
      const cm = (k) => round2(Number(attr(pgMar, `w:${k}`)) / TWIP_PER_CM);
      const margin = { top: cm("top"), right: cm("right"), bottom: cm("bottom"), left: cm("left") };
      // 页眉/页脚距页面边缘是可选属性；存在时一并保留，确保预设往返不丢失垂直位置。
      if (attr(pgMar, "w:header") != null) margin.header = cm("header");
      if (attr(pgMar, "w:footer") != null) margin.footer = cm("footer");
      opts.page = {
        ...opts.page,
        margin,
      };
    }
    const vAlign = attr(kid(sectPr, "w:vAlign"), "w:val");
    if (vAlign === "center" || vAlign === "bottom") opts.page = { ...opts.page, vAlign };
    if (kid(sectPr, "w:titlePg")) notes.push("文档启用了首页不同（titlePg），首页页眉页脚差异未提取");
  }

  // ---- 正文默认样式：docDefaults 优先，Normal 样式兜底 ----
  const stylesRoot = stylesXml ? rootEl(parseXml(stylesXml)) : null;
  const docDefaults = stylesRoot ? kid(stylesRoot, "w:docDefaults") : null;
  const normalStyle = stylesRoot
    ? kids(stylesRoot, "w:style").find((s) => attr(s, "w:styleId") === "Normal") ?? null
    : null;
  const ddRPr = kid(kid(docDefaults, "w:rPrDefault"), "w:rPr");
  const ddPPr = kid(kid(docDefaults, "w:pPrDefault"), "w:pPr");
  const nRPr = normalStyle ? kid(normalStyle, "w:rPr") : null;
  const nPPr = normalStyle ? kid(normalStyle, "w:pPr") : null;

  const bodyFont = fontFromRPr(nRPr) ?? fontFromRPr(ddRPr);
  const bodySize = halfPtFromRPr(nRPr) ?? halfPtFromRPr(ddRPr);
  const bodySpacing = spacingOf(nPPr) ?? spacingOf(ddPPr);
  // 正文行距：auto → 倍数；exact/atLeast → 固定行高（pt）
  let bodyLine = null;      // 倍数（lineRule=auto）
  let bodyLinePt = null;    // 固定行高 pt（lineRule=exact/atLeast）
  let bodyLineRule = null;  // 提取到的 lineRule（auto/exact/atLeast）
  if (bodySpacing?.line != null) {
    const rule = bodySpacing.lineRule ?? "auto";
    if (rule === "auto") {
      bodyLine = round2(bodySpacing.line / 240);
      bodyLineRule = "auto";
    } else if (rule === "exact" || rule === "atLeast" || rule === "exactly") {
      bodyLinePt = round2(bodySpacing.line / 20);
      bodyLineRule = rule === "exactly" ? "exact" : rule;
    } else {
      notes.push(`行距规则「${rule}」不受支持，未提取`);
    }
  }

  if (bodyFont) opts.fonts = { body: bodyFont };
  if (bodySize != null) opts.sizes = { body: bodySize };

  // 单行行高（pt）计算函数：auto 时 = 字号×倍数；exact/atLeast 时 = 固定值
  // 用于把标题段前/段后间距的"行"单位换算（auto 依赖各级标题字号，固定值则与字号无关）
  const lineHeightOf = bodyLine != null
    ? (sizePt) => (sizePt != null ? sizePt * bodyLine : null)
    : bodyLinePt != null ? () => bodyLinePt : null;

  // ---- 标题样式：六级字体 / 字号 / 加粗 / 对齐 / 间距 ----
  // 先按 styleId（HeadingN）匹配；未命中再按内置样式名 w:name（"heading N"/"标题 N"）兜底，
  // 兼容 WPS 等 styleId 非标准的文档
  const hStyles = new Map();
  for (const st of stylesRoot ? kids(stylesRoot, "w:style") : []) {
    const m = /^Heading([1-6])$/.exec(attr(st, "w:styleId") ?? "");
    if (m) hStyles.set(Number(m[1]), st);
  }
  if (!hStyles.size) {
    for (const st of stylesRoot ? kids(stylesRoot, "w:style") : []) {
      const nameVal = attr(kid(st, "w:name"), "w:val") ?? "";
      const m = /^(?:heading|标题)\s*([1-6])$/i.exec(nameVal.trim());
      if (m && !hStyles.has(Number(m[1]))) hStyles.set(Number(m[1]), st);
    }
    if (hStyles.size) notes.push("标题样式按样式名（而非标准 styleId）识别，提取结果可能不完整");
  }
  const headingSizes = [];
  const headingBold = [];
  const headingAlign = [];
  const hFonts = [];
  const beforeTw = [];
  const afterTw = [];
  for (let i = 1; i <= 6; i++) {
    const st = hStyles.get(i) ?? null;
    const rPr = st ? kid(st, "w:rPr") : null;
    const pPr = st ? kid(st, "w:pPr") : null;
    hFonts.push(fontFromRPr(rPr) ?? bodyFont ?? null);
    headingSizes.push(halfPtFromRPr(rPr) ?? bodySize ?? null);
    headingBold.push(st ? isBold(rPr) : false);
    const jc = attr(kid(pPr, "w:jc"), "w:val");
    headingAlign.push(jc === "center" ? "center" : jc === "right" ? "right" : null);
    const sp = spacingOf(pPr);
    beforeTw.push(sp?.before ?? null);
    afterTw.push(sp?.after ?? null);
  }
  if (hStyles.size) {
    const firstFont = hFonts.find(Boolean);
    if (firstFont) {
      const key = (f) => `${f?.eastAsia}/${f?.ascii}`;
      if (!hFonts.every((f) => f && key(f) === key(firstFont))) {
        notes.push("各级标题字体不一致，已采用一级标题字体");
      }
      opts.fonts = { ...opts.fonts, heading: firstFont };
    }
    opts.sizes = { ...opts.sizes, heading: headingSizes };
    opts.heading = {
      bold: headingBold,
      align: headingAlign,
      spacing: {
        beforeLines: null,
        afterLines: null,
        ...spacingPatch(beforeTw, headingSizes, lineHeightOf, "标题段前距", notes, "before"),
        ...spacingPatch(afterTw, headingSizes, lineHeightOf, "标题段后距", notes, "after"),
      },
    };
  }

  // 标题格式二义性检测：同一级标题在正文中出现多种对齐时，给出明确反馈
  detectHeadingAmbiguity(body, hStyles, notes);

  // ---- 段落：首行缩进 / 行距 / 段后距 / 对齐 ----
  const para = {};
  if (bodyLine != null) {
    para.line = bodyLine;
    para.lineRule = "auto";
  } else if (bodyLinePt != null) {
    para.line = bodyLinePt;
    para.lineRule = bodyLineRule === "atLeast" ? "atLeast" : "exact";
  }
  // 单行行高（pt）用于"行"单位段后距换算
  const lineHeightPtVal = bodyLine != null ? (bodySize != null ? bodySize * bodyLine : null) : bodyLinePt;
  if (bodySpacing?.after != null) {
    if (lineHeightPtVal != null) {
      para.afterLines = round2(bodySpacing.after / (lineHeightPtVal * 20));
    } else {
      para.afterLines = null;
      para.afterPt = round2(bodySpacing.after / 20);
    }
  } else {
    para.afterLines = null;
    para.afterPt = 0; // 未声明段后距 = 无段后距
  }
  para.firstLineChars = extractFirstLineChars(body, bodySize);
  const bodyAlign = extractBodyAlign(body, nPPr, ddPPr);
  if (bodyAlign) para.align = bodyAlign;
  opts.paragraph = para;

  // ---- 页眉页脚：文字 / 对齐 / 字体字号 / 页码字段 / 图片 ----
  const hdr = await parseHeaderFooter(zip, sectPr, rels, "header", notes);
  const ftr = await parseHeaderFooter(zip, sectPr, rels, "footer", notes);
  const images = {};
  if (hdr?.text) opts.header = { text: hdr.text, align: hdr.align ?? "left" };
  if (hdr?.image) images.header = hdr.image; // 二进制与高度，保存预设时落盘
  if (hdr?.font) opts.fonts = { ...opts.fonts, header: hdr.font };
  if (hdr?.size != null) opts.sizes = { ...opts.sizes, header: hdr.size };
  if (ftr?.text) opts.footer = { text: ftr.text, align: ftr.align ?? "left" };
  if (ftr?.image) images.footer = ftr.image;
  if (ftr?.font) opts.fonts = { ...opts.fonts, footer: ftr.font };
  if (ftr?.size != null) opts.sizes = { ...opts.sizes, footer: ftr.size };

  // ---- 页码：字段所在位置（页眉/页脚）+ 模板 + 数字样式 + 起始页 ----
  const pnFrom = ftr?.pn
    ? { ...ftr.pn, pos: "bottom" }
    : hdr?.pn
      ? { ...hdr.pn, pos: "top" }
      : null;
  const pgNumType = sectPr ? kid(sectPr, "w:pgNumType") : null;
  const styleVal = NUM_STYLE_FROM[attr(pgNumType, "w:fmt")] ?? null;
  const startVal = attr(pgNumType, "w:start");
  if (pnFrom) {
    const pn = { pos: pnFrom.pos, align: pnFrom.align ?? "left", format: pnFrom.template || "X" };
    if (styleVal) {
      if (pnFrom.template && pnFrom.template !== "X") {
        notes.push("页码数字样式与文字模板无法并存，已保留数字样式");
      }
      pn.format = styleVal;
    }
    if (startVal != null) pn.start = Number(startVal);
    opts.pageNumber = pn;
  } else if (styleVal != null || startVal != null) {
    notes.push("检测到页码设置（数字样式/起始页）但未找到页码字段，未提取");
  }

  return { options: opts, notes, images };
}

/** 从 docx 文件提取版式（extractPresetOptionsFromBuffer 的路径版） */
export async function extractPresetOptions(docxPath) {
  return extractPresetOptionsFromBuffer(fs.readFileSync(docxPath));
}

/**
 * 间距提取：优先换算为"行"单位（docx 的 before/after twip ÷ 单行行高×20），
 * 各级一致时返回 { xxxLines }；不一致或无法换算时退回 pt（取一级标题值）。
 * lineHeightOf(sizePt) 返回单行行高（pt），auto 依赖各级字号、exact/atLeast 为固定值。
 * prefix 为 "before" / "after"。
 */
function spacingPatch(twArr, sizeArr, lineHeightOf, label, notes, prefix) {
  const firstIdx = twArr.findIndex((t) => t != null);
  if (firstIdx === -1) return {};
  const key = `${prefix}Lines`;
  if (lineHeightOf != null) {
    const asLines = twArr.map((tw, i) => {
      if (tw == null) return null;
      const h = lineHeightOf(sizeArr[i]);
      return h != null ? tw / (h * 20) : null;
    });
    if (asLines.every((v) => v != null && Math.abs(v - asLines[firstIdx]) < 0.05)) {
      return { [key]: round2(asLines[firstIdx]) };
    }
  }
  const present = twArr.filter((t) => t != null);
  if (!present.every((t) => t === twArr[firstIdx])) {
    notes.push(`${label}各级不一致，已采用一级标题的值`);
  }
  return { [key]: null, [`${prefix}Pt`]: round2(twArr[firstIdx] / 20) };
}

/** 页面尺寸：匹配 PAGE_SIZES 命名尺寸；横向时先换算为纵向等效尺寸再匹配 */
function extractPageSize(pgSz) {
  const w = Number(attr(pgSz, "w:w"));
  const h = Number(attr(pgSz, "w:h"));
  if (!w || !h) return {};
  const orientation = attr(pgSz, "w:orient") ?? (w > h ? "landscape" : "portrait");
  const isLandscape = orientation === "landscape";
  const pw = isLandscape ? h : w;
  const ph = isLandscape ? w : h;
  for (const [name, dim] of Object.entries(PAGE_SIZES)) {
    if (dim.width === pw && dim.height === ph) return { size: name, orientation };
  }
  // 自定义尺寸："W,H"（cm，纵向等效）
  return { size: `${round1(pw / TWIP_PER_CM)},${round1(ph / TWIP_PER_CM)}`, orientation };
}

/** 正文首行缩进：统计正文顶层段落（跳过列表/样式/引用/代码块）最常见的缩进值 */
function extractFirstLineChars(body, bodySizePt) {
  const counts = new Map();
  const bump = (v) => counts.set(v, (counts.get(v) ?? 0) + 1);
  for (const p of kids(body ?? null, "w:p")) {
    const pPr = kid(p, "w:pPr");
    if (!pPr) { bump(0); continue; }
    // 列表（numPr）、标题/引用样式（pStyle）、代码块（shd 底纹）不计
    if (kid(pPr, "w:numPr") || kid(pPr, "w:pStyle") || kid(pPr, "w:shd")) continue;
    const ind = kid(pPr, "w:ind");
    if (!ind) { bump(0); continue; }
    const flc = attr(ind, "w:firstLineChars");
    if (flc) { bump(Number(flc) / 100); continue; }
    const fl = attr(ind, "w:firstLine");
    if (fl != null) { bump(bodySizePt ? round2(Number(fl) / (bodySizePt * 20)) : 0); continue; }
    // 仅有 left/right/hanging 缩进而无首行缩进：引用、代码块等 → 不计
  }
  if (!counts.size) return 0;
  let best = 0;
  let bestN = -1;
  for (const [v, n] of counts) {
    if (n > bestN) { best = v; bestN = n; }
  }
  return round2(best);
}

/** 正文对齐：顶层段落最常见 jc（需有过半数且至少 3 段），否则看 Normal/docDefaults */
function extractBodyAlign(body, nPPr, ddPPr) {
  const counts = new Map();
  for (const p of kids(body ?? null, "w:p")) {
    const pPr = kid(p, "w:pPr");
    if (!pPr || kid(pPr, "w:numPr") || kid(pPr, "w:pStyle")) continue;
    const jc = attr(kid(pPr, "w:jc"), "w:val");
    if (jc) counts.set(jc, (counts.get(jc) ?? 0) + 1);
  }
  let best = null;
  let total = 0;
  for (const [k, n] of counts) {
    total += n;
    if (!best || n > counts.get(best)) best = k;
  }
  const jc = total >= 3 && counts.get(best) >= total / 2 ? best
    : attr(kid(nPPr, "w:jc"), "w:val") ?? attr(kid(ddPPr, "w:jc"), "w:val");
  if (jc === "center" || jc === "right") return jc;
  if (jc === "both" || jc === "distribute") return "justify";
  return null; // left / 缺省 = 默认
}

/** 对齐值的中文名 */
function jcLabel(jc) {
  return { center: "居中", left: "左对齐", right: "右对齐", both: "两端对齐", distribute: "两端对齐" }[jc] ?? jc;
}

/**
 * 标题格式二义性检测：同一级标题在正文中存在多种对齐时，无法用单一格式标准表达，
 * 需要在结果中明确反馈（保留样式表定义，并告知不一致的具体分布）。
 */
function detectHeadingAmbiguity(body, hStyles, notes) {
  // 样式表各标题级的默认对齐（无 jc 时视为左对齐）
  const styleJc = new Map();
  for (let i = 1; i <= 6; i++) {
    const st = hStyles.get(i);
    const jc = st ? attr(kid(kid(st, "w:pPr"), "w:jc"), "w:val") : null;
    styleJc.set(i, jc ?? "left");
  }
  // 统计正文各标题级的"实际渲染对齐"分布（直接 jc 覆盖样式表，无则继承样式表）
  const dist = new Map(); // level -> Map(jc -> count)
  for (const p of kids(body ?? null, "w:p")) {
    const pPr = kid(p, "w:pPr");
    if (!pPr) continue;
    const styleId = attr(kid(pPr, "w:pStyle"), "w:val");
    const m = /^Heading([1-6])$/.exec(styleId ?? "");
    if (!m) continue;
    const level = Number(m[1]);
    const directJc = attr(kid(pPr, "w:jc"), "w:val");
    const actualJc = directJc ?? styleJc.get(level) ?? "left";
    if (!dist.has(level)) dist.set(level, new Map());
    const map = dist.get(level);
    map.set(actualJc, (map.get(actualJc) ?? 0) + 1);
  }
  for (const [level, map] of [...dist.entries()].sort((a, b) => a[0] - b[0])) {
    if (map.size <= 1) continue;
    const desc = [...map.entries()].map(([k, v]) => `${jcLabel(k)}×${v}`).join("、");
    notes.push(`H${level} 标题对齐不一致（${desc}），同一级标题只能采用一种对齐，已采用样式表定义「${jcLabel(styleJc.get(level))}」`);
  }
}

/** rels：rId → Target（相对 word/ 的路径） */
function parseRels(relsXml) {
  const map = new Map();
  if (!relsXml) return map;
  const root = rootEl(parseXml(relsXml));
  for (const rel of kids(root, "Relationship")) {
    map.set(attr(rel, "Id"), attr(rel, "Target"));
  }
  return map;
}

/**
 * 段落内容分段：文字段与页码字段段（X=当前页，Y=总页数）。
 * 字段缓存文本（separate 与 end 之间的 w:t）跳过，避免把 Word 缓存的页码值混入模板。
 */
function segmentsOf(p) {
  const segs = [];
  let fieldState = "outside"; // outside → instr → cached
  for (const c of p.elements || []) {
    if (c.type !== "element") continue;
    if (c.name === "w:r") {
      const instr = kid(c, "w:instrText");
      if (instr) {
        const t = textOf(instr).toUpperCase();
        if (t.includes("NUMPAGES")) segs.push({ field: "Y" });
        else if (t.includes("PAGE")) segs.push({ field: "X" });
        continue;
      }
      if (kid(c, "w:fldChar")) {
        const ty = attr(kid(c, "w:fldChar"), "w:fldCharType");
        if (ty === "begin") fieldState = "instr";
        else if (ty === "separate") fieldState = "cached";
        else if (ty === "end") fieldState = "outside";
        continue;
      }
      if (kid(c, "w:tab")) { segs.push({ text: "\t" }); continue; }
      if (fieldState === "cached") continue;
      for (const t of kids(c, "w:t")) segs.push({ text: textOf(t) });
    } else if (c.name === "w:hyperlink") {
      for (const t of deepAll(c, "w:t")) segs.push({ text: textOf(t) });
    } else if (c.name === "w:fldSimple") {
      const instr = (attr(c, "w:instr") ?? "").toUpperCase();
      if (instr.includes("NUMPAGES")) segs.push({ field: "Y" });
      else if (instr.includes("PAGE")) segs.push({ field: "X" });
    }
  }
  return segs;
}

/** 解析页眉/页脚 XML：文字行 / 对齐 / 字体字号 / 页码字段段落 / 图片 */
async function parseHeaderFooter(zip, sectPr, rels, kind, notes) {
  if (!sectPr) return null;
  const refs = kids(sectPr, `w:${kind}Reference`);
  const ref = refs.find((r) => attr(r, "w:type") === "default") ?? refs[0];
  if (!ref) return null;
  const target = rels.get(attr(ref, "r:id"));
  if (!target) return null;
  const xml = await zip.file(`word/${target.replace(/^\//, "")}`)?.async("string");
  if (!xml) return null;
  const root = rootEl(parseXml(xml));
  const label = kind === "header" ? "页眉" : "页脚";

  // 图片提取：沿 a:blip 的 r:embed 到该页眉/页脚自己的 rels 找媒体文件，读出二进制与显示高度
  const blips = deepAll(root, "a:blip");
  const drawingCount = deepAll(root, "w:drawing").length + deepAll(root, "w:pict").length;
  let image = null;
  if (blips.length) {
    const hfRelsXml = (await zip.file(`word/_rels/${target.replace(/^\//, "")}.rels`)?.async("string")) ?? "";
    const hfRels = parseRels(hfRelsXml);
    const mediaTarget = hfRels.get(attr(blips[0], "r:embed"));
    const buf = mediaTarget
      ? await zip.file(`word/${mediaTarget.replace(/^\//, "")}`)?.async("nodebuffer")
      : null;
    if (buf) {
      const extent = deepAll(root, "wp:extent")[0];
      const cy = extent ? Number(attr(extent, "cy")) : null;
      image = {
        data: buf,
        ext: path.extname(mediaTarget).slice(1).toLowerCase() || "png",
        heightCm: cy ? round2(cy / 360000) : undefined, // EMU → cm
      };
    }
  }
  if (drawingCount > (image ? 1 : 0)) {
    notes.push(`${label}含图形（如渐变色带），未提取`);
  } else if (!image && drawingCount > 0) {
    notes.push(`${label}含图片/图形，未提取`);
  }
  if (blips.length > 1) notes.push(`${label}含多张图片，仅提取第一张`);
  if (deepAll(root, "w:tbl").length) {
    notes.push(`${label}含表格布局，仅提取文字部分`);
  }

  const lines = [];
  let align = null;
  let font = null;
  let sizePt = null;
  let pn = null;
  for (const p of deepAll(root, "w:p")) {
    const segs = segmentsOf(p);
    if (segs.some((s) => s.field)) {
      const jc = attr(kid(kid(p, "w:pPr"), "w:jc"), "w:val");
      pn = { template: segs.map((s) => s.field ?? s.text).join(""), align: jcToAlign(jc) };
      font ??= firstFontIn(p);
      sizePt ??= firstSizeIn(p);
      continue;
    }
    const text = segs.map((s) => s.text).join("");
    if (!text.trim()) continue;
    lines.push(text);
    const jc = attr(kid(kid(p, "w:pPr"), "w:jc"), "w:val");
    align ??= jcToAlign(jc);
    font ??= firstFontIn(p);
    sizePt ??= firstSizeIn(p);
  }
  return {
    text: lines.length ? lines.join("\n") : null,
    align,
    font,
    size: sizePt,
    pn,
    image,
  };
}

/** 段内第一个含 rFonts 的 rPr → 字体 */
function firstFontIn(p) {
  const rPr = deepAll(p, "w:rPr").find((r) => kid(r, "w:rFonts"));
  return rPr ? fontFromRPr(rPr) : null;
}
/** 段内第一个含 sz 的 rPr → 字号 */
function firstSizeIn(p) {
  const rPr = deepAll(p, "w:rPr").find((r) => kid(r, "w:sz"));
  return rPr ? halfPtFromRPr(rPr) : null;
}
/** jc → MDDTT 对齐值；both/distribute/缺省视为默认（左） */
function jcToAlign(jc) {
  return jc === "center" || jc === "right" || jc === "left" ? jc : null;
}

// ================= 保存与摘要 =================

/** 预设名校验（不合法时抛错） */
export function assertPresetName(name) {
  if (!validPresetName(name)) {
    throw new Error(`预设名「${name}」不合法（仅限中英文、数字、下划线、连字符，且不以连字符开头）`);
  }
}

/**
 * 保存提取结果（extractPresetOptionsFromBuffer 的返回值）为自定义预设。
 * 页眉/页脚图片落盘到预设目录旁并在 options 中记录路径。返回 { file, options, notes }
 * dir 为预设目录（默认 ~/.mddtt/presets/，插件传入自己的存储目录）
 * source 仅用于在预设文件中记录来源文档
 */
export async function savePresetFromExtraction(
  extraction,
  name,
  { overwrite = false, dir = presetsDir(), source = null } = {}
) {
  assertPresetName(name);
  const file = customPresetPath(name, dir);
  if (fs.existsSync(file) && !overwrite) {
    throw new Error(`预设已存在「${file}」，如需覆盖请加 --overwrite`);
  }
  const { options, notes, images } = extraction;
  fs.mkdirSync(dir, { recursive: true });
  // 图片落盘：<预设目录>/<名>.header.png 等，options 记录绝对路径与显示高度
  for (const kind of ["header", "footer"]) {
    const img = images?.[kind];
    if (!img) continue;
    const imgPath = path.join(dir, `${name}.${kind}.${img.ext}`);
    fs.writeFileSync(imgPath, img.data);
    options[kind] = {
      ...options[kind],
      image: { path: imgPath, ...(img.heightCm != null ? { heightCm: img.heightCm } : {}) },
    };
  }
  const data = {
    name,
    source: source ?? null,
    extractedAt: new Date().toISOString(),
    notes,
    options,
  };
  fs.writeFileSync(file, JSON.stringify(data, null, 2) + "\n", "utf-8");
  return { file, options, notes };
}

/** 提取并保存自定义预设（路径版）。返回 { file, options, notes } */
export async function extractAndSavePreset(docxPath, name, opts = {}) {
  assertPresetName(name); // 先校验预设名，保持命令行原有的错误优先级
  const extraction = await extractPresetOptions(docxPath);
  return savePresetFromExtraction(extraction, name, { ...opts, source: path.resolve(docxPath) });
}

/** 提取结果的人读摘要（逐行） */
export function presetSummaryLines(options) {
  const L = [];
  const cn = { left: "左", center: "居中", right: "右", justify: "两端" };
  if (options.page) {
    const p = options.page;
    const m = p.margin
      ? `，边距 上${p.margin.top}/右${p.margin.right}/下${p.margin.bottom}/左${p.margin.left}cm`
      : "";
    const hf = p.margin
      ? `${p.margin.header != null ? `，页眉距顶端${p.margin.header}cm` : ""}${p.margin.footer != null ? `，页脚距底端${p.margin.footer}cm` : ""}`
      : "";
    const v = p.vAlign ? `，内容${cn[p.vAlign] ?? p.vAlign}对齐` : "";
    L.push(`页面    ${p.size} ${p.orientation === "landscape" ? "横向" : "纵向"}${m}${hf}${v}`);
  }
  if (options.fonts?.body || options.sizes?.body || options.paragraph) {
    const parts = [];
    if (options.fonts?.body) parts.push(`${options.fonts.body.eastAsia} / ${options.fonts.body.ascii}`);
    if (options.sizes?.body) parts.push(`${options.sizes.body}pt`);
    const pa = options.paragraph ?? {};
    if (pa.firstLineChars != null) parts.push(`首行缩进${pa.firstLineChars}字符`);
    if (pa.line != null) {
      const rule = pa.lineRule ?? "auto";
      const unit = rule === "auto" ? "倍" : "pt";
      const ruleLabel = { auto: "", exact: "（固定）", atLeast: "（最小值）" }[rule] ?? "";
      parts.push(`行距${pa.line}${unit}${ruleLabel}`);
    }
    if (pa.afterLines != null) parts.push(`段后${pa.afterLines}行`);
    else if (pa.afterPt != null) parts.push(`段后${pa.afterPt}pt`);
    if (pa.align) parts.push(`${cn[pa.align] ?? pa.align}对齐`);
    if (parts.length) L.push(`正文    ${parts.join("，")}`);
  }
  if (options.fonts?.heading || options.sizes?.heading || options.heading) {
    const parts = [];
    if (options.fonts?.heading) parts.push(`${options.fonts.heading.eastAsia} / ${options.fonts.heading.ascii}`);
    if (options.sizes?.heading) parts.push(`H1-H6 字号 ${options.sizes.heading.map((s) => s ?? "?").join("/")}pt`);
    if (options.heading?.align) parts.push(`对齐 ${options.heading.align.map((a) => cn[a] ?? "左").join("/")}`);
    if (options.heading?.bold) parts.push(`加粗 ${options.heading.bold.map((b) => (b ? "√" : "×")).join("")}`);
    if (parts.length) L.push(`标题    ${parts.join("，")}`);
  }
  if (options.header?.text || options.header?.image) {
    const extra = [
      `${cn[options.header.align] ?? "居中"}对齐`,
      options.fonts?.header ? options.fonts.header.eastAsia : null,
      options.sizes?.header ? `${options.sizes.header}pt` : null,
      options.header?.image ? `图片已保存（高 ${options.header.image.heightCm ?? "?"}cm）` : null,
    ].filter(Boolean);
    const textInfo = options.header.text ? `${options.header.text.split("\n").length}行文字，` : "";
    L.push(`页眉    ${textInfo}${extra.join(" ")}`);
  }
  if (options.footer?.text) {
    L.push(`页脚    ${options.footer.text.split("\n").length}行文字，${cn[options.footer.align] ?? "居中"}对齐`);
  }
  if (options.pageNumber) {
    const pn = options.pageNumber;
    L.push(
      `页码    ${pn.pos === "bottom" ? "页脚" : "页眉"}${pn.align ? `${cn[pn.align] ?? ""}对齐` : ""}，格式「${pn.format}」${pn.start != null ? `，起始页码 ${pn.start}` : ""}`
    );
  }
  return L;
}
