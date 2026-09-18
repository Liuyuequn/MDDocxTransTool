// md → docx 转换编排层：解析 markdown → 组装 docx Document → 写出文件
// 具体职责分散在：page.js（页面属性）、styles.js（样式编号）、
// blocks.js（块级解析）、inline.js（行内解析）、header-footer.js（页眉页脚）

import fs from "node:fs";
import path from "node:path";
import JSZip from "jszip";
import { Document, Packer, Paragraph } from "docx";
import { cmToTwip } from "./options.js";
import { buildPageSize, toSectionVAlign, buildPageNumbers } from "./page.js";
import { buildStyles, buildBulletNumbering, buildOrderedNumbering } from "./styles.js";
import { parseBlocks } from "./blocks.js";
import { buildHeaderFooter } from "./header-footer.js";
import { createMarkdownParser } from "./markdown.js";

// 解析器按 breaks 配置缓存：一次转换只创建一次（同一转换内的多次 buildDocument 复用）
let cachedParser = null;

function parserFor(breaks) {
  if (!cachedParser || cachedParser.breaks !== breaks) {
    cachedParser = { breaks, md: createMarkdownParser(breaks) };
  }
  return cachedParser.md;
}

/**
 * 将 Markdown 文本转换为 docx 二进制内容（opts 为合并完成的完整配置）。
 * basePath 用于解析相对图片路径，通常是 md 文件所在目录。
 * 不触碰文件系统之外的输出，便于宿主（如 VS Code 插件）自行决定读写方式。
 */
export async function convertMarkdownText(mdText, opts, basePath) {
  const doc = buildDocument(mdText, opts, basePath);
  const buffer = await Packer.toBuffer(doc);
  return fixSettingsOrder(buffer);
}

/** 将 markdown 文件转换为 docx 文件（opts 为合并完成的完整配置） */
export async function convertMarkdownFile(inputPath, outputPath, opts) {
  const mdText = fs.readFileSync(inputPath, "utf-8");
  const buffer = await convertMarkdownText(mdText, opts, path.dirname(path.resolve(inputPath)));
  fs.writeFileSync(outputPath, buffer);
}

/**
 * 修正 settings.xml 中 w:compat 的子元素顺序。
 *
 * docx 9.x 的 Compatibility 生成器把 w:compatSetting 固定写在 w:compat 的第一个位置，
 * 但 OOXML 的 CT_Compat 要求它排在各个兼容性开关（如 w:doNotExpandShiftReturn）之后，
 * 顺序不合会让文档通不过 Open XML Schema 校验。这里把它移到末尾。
 */
async function fixSettingsOrder(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const file = zip.file("word/settings.xml");
  if (!file) return buffer;
  const xml = await file.async("string");
  const compat = /<w:compat>([\s\S]*?)<\/w:compat>/.exec(xml);
  if (!compat) return buffer;
  const settings = [...compat[1].matchAll(/<w:compatSetting\b[^>]*\/>/g)].map((m) => m[0]);
  if (!settings.length || compat[1].trim().endsWith(settings[settings.length - 1])) return buffer;
  const rest = compat[1].replace(/<w:compatSetting\b[^>]*\/>/g, "");
  zip.file("word/settings.xml", xml.replace(compat[0], `<w:compat>${rest}${settings.join("")}</w:compat>`));
  return zip.generateAsync({ type: "nodebuffer" });
}

function buildDocument(mdText, opts, basePath) {
  const tokens = parserFor(opts.markdown.breaks).parse(mdText, {});
  const ctx = makeCtx(opts, basePath);
  const state = { orderedInstance: 0 };
  const blocks = parseBlocks(tokens, 0, tokens.length, ctx, state);

  const headerFooter = buildHeaderFooter(opts, basePath);

  return new Document({
    styles: buildStyles(opts),
    numbering: { config: [buildBulletNumbering(), buildOrderedNumbering()] },
    // doNotExpandShiftReturn：不把手动换行符（w:br，即 Markdown 的换行）结尾的行
    // 拉伸到整行宽。两端对齐（--align justify）时，Word 默认会把段落内除末行以外的
    // 每一行都扩展到左右边距，表现为字间距被强行撑开；本工具用换行符分行，因此必须
    // 关闭该行为，换行后的各行保持正常字间距与对齐。
    compatibility: { doNotExpandShiftReturn: true },
    sections: [
      {
        properties: {
          page: {
            size: buildPageSize(opts),
            margin: {
              top: cmToTwip(opts.page.margin.top),
              right: cmToTwip(opts.page.margin.right),
              bottom: cmToTwip(opts.page.margin.bottom),
              left: cmToTwip(opts.page.margin.left),
              header: opts.page.margin.header != null
                ? cmToTwip(opts.page.margin.header)
                : undefined,
              footer: opts.page.margin.footer != null
                ? cmToTwip(opts.page.margin.footer)
                : undefined,
            },
            pageNumbers: buildPageNumbers(opts),
          },
          verticalAlign: toSectionVAlign(opts.page.vAlign),
          titlePage: headerFooter.titlePage || undefined,
        },
        headers: headerFooter.headers,
        footers: headerFooter.footers,
        children: blocks.length ? blocks : [new Paragraph({ children: [] })],
      },
    ],
  });
}

/** 解析上下文：块级/行内解析共享的配置与状态 */
function makeCtx(opts, basePath) {
  return {
    basePath,
    opts,
    listLevel: -1,
    listType: null,
    instance: 0,
    quote: false,
    forceBold: false,
    // docx 9.x 未显式指定 wp:docPr id 时可能为多个图片重复生成 id=1。
    // 所有通过浅拷贝派生的解析上下文共享此计数器。
    drawingIds: { next: 1 },
  };
}
