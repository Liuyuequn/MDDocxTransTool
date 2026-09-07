// md → docx 转换编排层：解析 markdown → 组装 docx Document → 写出文件
// 具体职责分散在：page.js（页面属性）、styles.js（样式编号）、
// blocks.js（块级解析）、inline.js（行内解析）、header-footer.js（页眉页脚）

import fs from "node:fs";
import path from "node:path";
import MarkdownIt from "markdown-it";
import { Document, Packer, Paragraph } from "docx";
import { cmToTwip } from "./options.js";
import { buildPageSize, toSectionVAlign, buildPageNumbers } from "./page.js";
import { buildStyles, buildBulletNumbering, buildOrderedNumbering } from "./styles.js";
import { parseBlocks } from "./blocks.js";
import { buildHeaderFooter } from "./header-footer.js";

const md = new MarkdownIt({ html: false, linkify: true, typographer: true });

/** 将 markdown 文件转换为 docx 文件（opts 为合并完成的完整配置） */
export async function convertMarkdownFile(inputPath, outputPath, opts) {
  const mdText = fs.readFileSync(inputPath, "utf-8");
  const doc = buildDocument(mdText, opts, path.dirname(path.resolve(inputPath)));
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
}

function buildDocument(mdText, opts, basePath) {
  const tokens = md.parse(mdText, {});
  const ctx = makeCtx(opts, basePath);
  const state = { orderedInstance: 0 };
  const blocks = parseBlocks(tokens, 0, tokens.length, ctx, state);

  const headerFooter = buildHeaderFooter(opts, basePath);

  return new Document({
    styles: buildStyles(opts),
    numbering: { config: [buildBulletNumbering(), buildOrderedNumbering()] },
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
