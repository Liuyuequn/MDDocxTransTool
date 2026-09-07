// 块级元素解析：标题/段落/列表/引用/代码块/表格/分隔线
// 消费 markdown-it 的 token 流，产出 docx 块级对象

import {
  AlignmentType,
  BorderStyle,
  HeadingLevel,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  VerticalAlignTable,
  WidthType,
} from "docx";
import { ptToHalfPoint } from "./options.js";
import { ALIGN_MAP, bodyParagraphSpacing, fontObj } from "./styles.js";
import { parseInline, getAttr } from "./inline.js";

const HEADINGS = [
  HeadingLevel.HEADING_1,
  HeadingLevel.HEADING_2,
  HeadingLevel.HEADING_3,
  HeadingLevel.HEADING_4,
  HeadingLevel.HEADING_5,
  HeadingLevel.HEADING_6,
];

/** 在 [start, end) 范围内查找与 openType 配对的 close token 下标 */
export function findClose(tokens, start, end, openType, closeType) {
  let depth = 0;
  for (let j = start; j < end; j++) {
    if (tokens[j].type === openType) depth++;
    else if (tokens[j].type === closeType && --depth === 0) return j;
  }
  return -1;
}

export function parseBlocks(tokens, start, end, ctx, state) {
  const out = [];
  let i = start;
  while (i < end) {
    const t = tokens[i];
    switch (t.type) {
      case "heading_open": {
        const level = Number(t.tag.slice(1));
        const hCtx = { ...ctx, uppercase: ctx.opts.styles.headingUppercase };
        out.push(
          new Paragraph({
            heading: HEADINGS[level - 1],
            children: parseInline(tokens[i + 1].children, hCtx),
          })
        );
        i += 3;
        break;
      }
      case "paragraph_open": {
        const inlineTok = tokens[i + 1];
        const imgInfo = imageParagraphInfo(inlineTok.children);
        if (imgInfo) {
          out.push(
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 120, after: imgInfo.title ? 60 : 120 },
              children: parseInline(inlineTok.children, ctx),
            })
          );
          if (imgInfo.title) {
            // 图片标题（![alt](src "标题")）在图片下方居中输出
            out.push(
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 120 },
                children: [new TextRun({ text: imgInfo.title, color: "666666", size: ptToHalfPoint(ctx.opts.sizes.body) })],
              })
            );
          }
        } else {
          out.push(makeParagraph(parseInline(inlineTok.children, ctx), ctx));
        }
        i += 3;
        break;
      }
      case "bullet_list_open":
      case "ordered_list_open": {
        const closeType = t.type.replace("_open", "_close");
        const close = findClose(tokens, i, end, t.type, closeType);
        if (close === -1) { i++; break; }
        const isOrdered = t.type === "ordered_list_open";
        const itemCtx = {
          ...ctx,
          listLevel: ctx.listLevel + 1,
          listType: isOrdered ? "md-ordered" : "md-bullet",
          instance: isOrdered ? ++state.orderedInstance : 0,
        };
        out.push(...parseBlocks(tokens, i + 1, close, itemCtx, state));
        i = close + 1;
        break;
      }
      case "list_item_open":
      case "list_item_close":
      case "space":
        i++;
        break;
      case "fence":
      case "code_block": {
        out.push(...makeCodeBlock(t.content, ctx));
        i++;
        break;
      }
      case "blockquote_open": {
        const close = findClose(tokens, i, end, "blockquote_open", "blockquote_close");
        if (close === -1) { i++; break; }
        out.push(...parseBlocks(tokens, i + 1, close, { ...ctx, quote: true }, state));
        i = close + 1;
        break;
      }
      case "hr": {
        out.push(
          new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "AAAAAA", space: 1 } },
            spacing: { before: 240, after: 240 },
          })
        );
        i++;
        break;
      }
      case "table_open": {
        const close = findClose(tokens, i, end, "table_open", "table_close");
        if (close === -1) { i++; break; }
        out.push(parseTable(tokens, i + 1, close, ctx));
        out.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
        i = close + 1;
        break;
      }
      default:
        i++;
        break;
    }
  }
  return out;
}

/** 普通段落构造：应用正文对齐、首行缩进（仅普通段落）、列表编号、引用样式 */
function makeParagraph(runs, ctx, extra = {}) {
  const { opts } = ctx;
  // 行距同时写入段落直接格式。只依赖 docDefaults 时，Word 可能补全内置
  // Normal/List Paragraph 样式并在界面中把有效行距显示为 1 倍。
  const props = { children: runs, spacing: bodyParagraphSpacing(opts) };
  const isNormal = ctx.listLevel < 0 && !ctx.quote;
  if (isNormal && opts.paragraph.align) props.alignment = ALIGN_MAP[opts.paragraph.align];
  if (isNormal && opts.paragraph.firstLineChars > 0) {
    props.indent = { firstLineChars: Math.round(opts.paragraph.firstLineChars * 100) };
  }
  if (ctx.listLevel >= 0) {
    props.numbering = {
      reference: ctx.listType,
      level: Math.min(ctx.listLevel, 5),
      instance: ctx.instance,
    };
  }
  if (ctx.quote) {
    props.indent = { left: 480 };
    props.border = { left: { style: BorderStyle.SINGLE, size: 18, color: "CCCCCC", space: 10 } };
  }
  return new Paragraph({ ...props, ...extra });
}

/** 判断段落是否只包含图片（图片独立成段并居中，不参与首行缩进），返回 { title } 或 null */
function imageParagraphInfo(children) {
  const meaningful = (children || []).filter((tok) => tok.type !== "text" || tok.content.trim() !== "");
  let imageTok = null;
  if (meaningful.length === 1 && meaningful[0].type === "image") {
    imageTok = meaningful[0];
  } else if (
    meaningful.length === 3 &&
    meaningful[0].type === "link_open" &&
    meaningful[1].type === "image" &&
    meaningful[2].type === "link_close"
  ) {
    // 链接包裹图片：[![alt](src)](url)
    imageTok = meaningful[1];
  }
  if (!imageTok) return null;
  return { title: getAttr(imageTok, "title") || null };
}

function makeCodeBlock(content, ctx) {
  const { opts } = ctx;
  const lines = content.replace(/\n$/, "").split("\n");
  return lines.map((line, idx) =>
    new Paragraph({
      children: [
        new TextRun({
          text: line.length ? line : " ",
          font: fontObj(opts.fonts.code),
          size: ptToHalfPoint(opts.sizes.code),
        }),
      ],
      shading: { type: ShadingType.CLEAR, fill: "F2F2F2" },
      spacing: { after: idx === lines.length - 1 ? 200 : 0 },
      indent: { left: 120, right: 120 },
    })
  );
}

function parseTable(tokens, start, end, ctx) {
  const rows = [];
  let isHeader = true;
  let colCount = 0;
  let j = start;
  while (j < end) {
    if (tokens[j].type !== "tr_open") { j++; continue; }
    const trClose = findClose(tokens, j, end, "tr_open", "tr_close");
    if (colCount === 0) {
      for (let k = j + 1; k < trClose; k++) {
        if (tokens[k].type === "th_open" || tokens[k].type === "td_open") colCount++;
      }
    }
    const cells = [];
    let k = j + 1;
    while (k < trClose) {
      const type = tokens[k].type;
      if (type === "th_open" || type === "td_open") {
        const closeType = type.replace("_open", "_close");
        const cellEnd = findClose(tokens, k, trClose, type, closeType);
        const inlineTok = tokens.slice(k + 1, cellEnd).find((x) => x.type === "inline");
        const cellCtx = isHeader ? { ...ctx, forceBold: true } : ctx;
        const runs = parseInline(inlineTok ? inlineTok.children : [], cellCtx);
        const style = getAttr(tokens[k], "style");
        // 水平对齐：表头居中；数据行按 markdown 语法标注（无标注则默认左对齐）
        let alignment;
        if (isHeader) {
          alignment = AlignmentType.CENTER;
        } else if (style) {
          if (/text-align:\s*center/.test(style)) alignment = AlignmentType.CENTER;
          else if (/text-align:\s*(right|end)/.test(style)) alignment = AlignmentType.RIGHT;
        }
        cells.push(
          new TableCell({
            children: [new Paragraph({ children: runs, alignment })],
            verticalAlign: VerticalAlignTable.CENTER,
            width: colCount ? { size: Math.floor(1000 / colCount) / 10, type: WidthType.PERCENTAGE } : undefined,
            shading: isHeader ? { type: ShadingType.CLEAR, fill: "EFEFEF" } : undefined,
          })
        );
        k = cellEnd + 1;
      } else {
        k++;
      }
    }
    // false 会被 docx 序列化为不符合当前 Schema 的 w:tblHeader w:val="false"；
    // 非表头行应直接省略该可选元素。
    rows.push(new TableRow({ children: cells, tableHeader: isHeader || undefined }));
    if (isHeader) isHeader = false;
    j = trClose + 1;
  }
  return new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } });
}
