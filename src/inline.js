// 行内元素解析：加粗/斜体/删除线/行内代码/超链接/图片/换行
// 以及图片文件加载（正文图片与页眉 logo 共用）

import fs from "node:fs";
import path from "node:path";
import { imageSize } from "image-size";
import {
  ExternalHyperlink,
  ImageRun,
  ShadingType,
  TextRun,
} from "docx";
import { ptToHalfPoint } from "./options.js";
import { fontObj } from "./styles.js";
import { contentWidthTwip } from "./page.js";

const IMAGE_TYPES = { ".png": "png", ".jpg": "jpg", ".jpeg": "jpg", ".gif": "gif", ".bmp": "bmp" };

/**
 * 换行后补足首行缩进。
 *
 * Word 的 firstLineChars 只作用于段落真正意义上的第一行；由手动换行符（w:br）
 * 产生的后续各行不会缩进。中文排版要求一段之内每一行都缩进两字符，因此在换行符后
 * 补两个全角空格（U+3000）——一个全角空格恰好等于一个汉字宽度，实测可使后续行
 * 与首行完全对齐（零宽空格、窄空格等写法均无法精确对齐）。
 */
const LINE_INDENT = "\u3000\u3000";

/** 正文图片宽度上限（像素）：页面内容区宽度按 96 DPI 换算（1px = 15twip），随页面设置动态变化 */
export function maxImageWidthPx(opts) {
  return Math.floor(contentWidthTwip(opts) / 15);
}

/** 读取 markdown-it token 的属性 */
export function getAttr(token, name) {
  const attr = (token.attrs || []).find(([k]) => k === name);
  return attr ? attr[1] : null;
}

/** 加载并校验图片文件：返回 { data, dim, type } */
export function loadImage(absPath) {
  if (!fs.existsSync(absPath)) throw new Error(`找不到图片文件: ${absPath}`);
  const data = fs.readFileSync(absPath);
  const dim = imageSize(data);
  const type = IMAGE_TYPES[path.extname(absPath).toLowerCase()];
  if (!type) throw new Error(`不支持的图片格式: ${path.extname(absPath)}`);
  return { data, dim, type };
}

/** 解析行内 token 序列为 docx run 序列 */
export function parseInline(children, ctx) {
  const runs = [];
  if (!children) return runs;
  let bold = false;
  let italic = false;
  let strike = false;
  let link = null;
  for (const t of children) {
    switch (t.type) {
      case "text":
      case "text_special":
        runs.push(makeRun(ctx.uppercase ? t.content.toUpperCase() : t.content, { bold, italic, strike, link, ctx }));
        break;
      case "strong_open":
        bold = true;
        break;
      case "strong_close":
        bold = false;
        break;
      case "em_open":
        italic = true;
        break;
      case "em_close":
        italic = false;
        break;
      case "s_open":
        strike = true;
        break;
      case "s_close":
        strike = false;
        break;
      case "code_inline":
        runs.push(
          new TextRun({
            text: t.content,
            font: fontObj(ctx.opts.fonts.code),
            size: ptToHalfPoint(ctx.opts.sizes.code),
            shading: { type: ShadingType.CLEAR, fill: "EFEFEF" },
          })
        );
        break;
      case "link_open":
        link = getAttr(t, "href");
        break;
      case "link_close":
        link = null;
        break;
      case "image":
        try {
          const imgRun = makeImageRun(t, ctx);
          // 图片被超链接包裹（[![alt](src)](url)）时，图片本身可点击跳转
          runs.push(link ? new ExternalHyperlink({ link, children: [imgRun] }) : imgRun);
        } catch {
          runs.push(makeRun(`[图片缺失: ${t.content || getAttr(t, "src") || "未知来源"}]`, { ctx }));
        }
        break;
      case "hardbreak":
        // 换行符后补首行缩进，使段内后续各行与首行左端对齐（中文排版惯例）
        runs.push(new TextRun({ break: 1 }));
        runs.push(new TextRun({ text: LINE_INDENT }));
        break;
      case "softbreak":
        runs.push(new TextRun({ text: " " }));
        break;
      default:
        break;
    }
  }
  return runs;
}

function makeRun(text, { bold, italic, strike, link, ctx }) {
  const italicOn = italic && ctx.opts.styles.italic;
  const strikeOn = strike && ctx.opts.styles.strike;
  const runProps = {
    text,
    bold: bold || ctx.forceBold || undefined,
    italics: italicOn || undefined,
    strike: strikeOn || undefined,
  };
  if (link) {
    return new ExternalHyperlink({
      link,
      children: [new TextRun({ ...runProps, style: "Hyperlink" })],
    });
  }
  return new TextRun(runProps);
}

/** 正文图片：超出内容区宽度时等比缩小 */
function makeImageRun(token, ctx) {
  const src = getAttr(token, "src");
  if (!src) throw new Error("图片缺少 src 属性");
  const absPath = path.resolve(ctx.basePath, decodeURIComponent(src.replace(/^file:\/\//, "")));
  const { data, dim, type } = loadImage(absPath);
  const scale = Math.min(1, maxImageWidthPx(ctx.opts) / dim.width);
  const drawingId = String(ctx.drawingIds.next++);
  return new ImageRun({
    type,
    data,
    altText: {
      id: drawingId,
      name: `MDDTT Image ${drawingId}`,
      description: token.content || undefined,
    },
    transformation: {
      width: Math.round(dim.width * scale),
      height: Math.round(dim.height * scale),
    },
  });
}
