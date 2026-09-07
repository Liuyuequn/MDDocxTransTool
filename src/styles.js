// 样式与编号：docx 样式表、字体对象、间距换算、列表编号配置
// 被 converter（组装 Document）、blocks（代码块字体）、header-footer（页眉页脚字体）共用

import {
  AlignmentType,
  LevelFormat,
  LineRuleType,
} from "docx";
import { ptToHalfPoint, ptToTwip } from "./options.js";

const BULLET_CHARS = ["\u2022", "\u25E6", "\u2022", "\u25E6", "\u2022", "\u25E6"];

/** 对齐方式字符串 → docx 枚举 */
export const ALIGN_MAP = {
  left: AlignmentType.LEFT,
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
  justify: AlignmentType.JUSTIFIED,
};

/** docx 字体对象：中文 eastAsia / 西文 ascii */
export function fontObj(f) {
  return { ascii: f.ascii, hAnsi: f.ascii, eastAsia: f.eastAsia };
}

/** 标题字体：未单独指定时跟随正文 */
export function headingFont(opts) {
  return fontObj(opts.fonts.heading || opts.fonts.body);
}

/**
 * 解析间距：lines 字段优先（按 字号×行高 换算 twip），否则用 pt；
 * lineCfg 为行距配置 { line, lineRule }：
 *   - lineRule=auto（默认）时 line 为倍数，w:line = line×240
 *   - lineRule=exact/atLeast 时 line 为固定行高（pt），w:line = pt×20
 */
export function resolveSpacing(spacing, sizePt, lineCfg) {
  const out = {};
  if (spacing.beforePt != null) out.before = ptToTwip(spacing.beforePt);
  if (spacing.beforeLines != null) out.before = linesToTwip(spacing.beforeLines, sizePt, lineCfg);
  if (spacing.afterPt != null) out.after = ptToTwip(spacing.afterPt);
  if (spacing.afterLines != null) out.after = linesToTwip(spacing.afterLines, sizePt, lineCfg);
  if (lineCfg && lineCfg.line != null) {
    const rule = lineCfg.lineRule || "auto";
    if (rule === "auto") {
      out.line = Math.round(lineCfg.line * 240);
      out.lineRule = LineRuleType.AUTO;
    } else if (rule === "exact" || rule === "atLeast") {
      out.line = ptToTwip(lineCfg.line);
      out.lineRule = rule === "exact" ? LineRuleType.EXACT : LineRuleType.AT_LEAST;
    }
  }
  return out;
}

/** 单行行高（pt）：auto 时为 字号×倍数；exact/atLeast 时为固定值 */
function lineHeightPt(lineCfg, sizePt) {
  if (!lineCfg || lineCfg.line == null) return null;
  const rule = lineCfg.lineRule || "auto";
  if (rule === "auto") return sizePt * lineCfg.line;
  return lineCfg.line;
}

/** "行"单位 → twip：1 行 = 单行行高（pt）× 20 */
export function linesToTwip(lines, sizePt, lineCfg) {
  const h = lineHeightPt(lineCfg, sizePt);
  if (h == null) return Math.round(lines * sizePt * 1.28 * 20);
  return Math.round(lines * h * 20);
}

/**
 * 正文段落间距。
 * 除写入 docDefaults 外，普通正文段落也会直接使用这组属性，避免 Word
 * 在补全 Normal/List Paragraph 样式时把倍数行距显示或覆盖为单倍行距。
 */
export function bodyParagraphSpacing(opts) {
  return resolveSpacing(
    { beforePt: 0, afterPt: opts.paragraph.afterPt, afterLines: opts.paragraph.afterLines },
    opts.sizes.body,
    opts.paragraph
  );
}

/** docx 样式表：正文默认 + 六级标题；超链接使用 docx 内置的 Hyperlink 字符样式 */
export function buildStyles(opts) {
  const headingDefaults = (level) => {
    const i = level - 1;
    return {
      run: {
        font: headingFont(opts),
        size: ptToHalfPoint(opts.sizes.heading[i]),
        bold: opts.heading.bold[i] || undefined,
        color: "000000",
      },
      paragraph: {
        alignment: opts.heading.align[i] ? ALIGN_MAP[opts.heading.align[i]] : AlignmentType.LEFT,
        spacing: resolveSpacing(opts.heading.spacing, opts.sizes.heading[i], opts.paragraph),
      },
    };
  };
  return {
    default: {
      document: {
        run: { font: fontObj(opts.fonts.body), size: ptToHalfPoint(opts.sizes.body) },
        paragraph: { spacing: bodyParagraphSpacing(opts) },
      },
      heading1: headingDefaults(1),
      heading2: headingDefaults(2),
      heading3: headingDefaults(3),
      heading4: headingDefaults(4),
      heading5: headingDefaults(5),
      heading6: headingDefaults(6),
    },
  };
}

export function buildBulletNumbering() {
  return {
    reference: "md-bullet",
    levels: Array.from({ length: 6 }, (_, l) => ({
      level: l,
      format: LevelFormat.BULLET,
      text: BULLET_CHARS[l],
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720 * (l + 1), hanging: 360 } } },
    })),
  };
}

export function buildOrderedNumbering() {
  return {
    reference: "md-ordered",
    levels: Array.from({ length: 6 }, (_, l) => ({
      level: l,
      format: LevelFormat.DECIMAL,
      text: `%${l + 1}.`,
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720 * (l + 1), hanging: 360 } } },
    })),
  };
}
