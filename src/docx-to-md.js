// docx → Markdown 编排层：OOXML 预处理 → mammoth HTML → Markdown 渲染

import fs from "node:fs";
import path from "node:path";
import mammoth from "mammoth";
import { createImageConverter } from "./docx-import/image-handler.js";
import { prepareDocxBufferForMarkdown } from "./docx-import/package-reader.js";
import { renderMarkdown } from "./docx-import/renderers/markdown.js";

/** 图片输出目录名（位于输出 md 文件同目录下） */
export const PICTURE_DIR = "MDPictures";

const STYLE_MAP = [
  ...Array.from({ length: 6 }, (_, index) =>
    `p[style-name='Heading ${index + 1}'] => h${index + 1}:fresh`
  ),
  ...Array.from({ length: 6 }, (_, index) =>
    `p[style-name='标题 ${index + 1}'] => h${index + 1}:fresh`
  ),
  "p[style-name='Quote'] => blockquote",
];

/**
 * 将 docx 二进制内容转换为 Markdown 字符串。
 * outputDir 为图片输出目录的父目录（图片写入 outputDir/MDPictures/）。
 * 修订默认接受并附原文注释；批注附于范围之后；文本框和分节内容线性排列。
 */
export async function convertDocxBufferToMarkdown(input, outputDir) {
  const { buffer, annotations } = await prepareDocxBufferForMarkdown(input);
  const { value: rawHtml } = await mammoth.convertToHtml(
    { buffer },
    {
      styleMap: STYLE_MAP,
      convertImage: createImageConverter(outputDir, PICTURE_DIR),
    }
  );
  return renderMarkdown(rawHtml, annotations);
}

/** 将 docx 转换为 Markdown 字符串（convertDocxBufferToMarkdown 的路径版） */
export async function convertDocxToMarkdown(docxPath, outputDir) {
  if (!fs.existsSync(docxPath)) throw new Error(`找不到文件 ${docxPath}`);
  return convertDocxBufferToMarkdown(fs.readFileSync(docxPath), outputDir);
}

/** 将 docx 转换为 Markdown 并写入文件。 */
export async function convertDocxFile(docxPath, outputPath) {
  const out = outputPath || docxPath.replace(/\.docx$/i, ".md");
  const md = await convertDocxToMarkdown(docxPath, path.dirname(path.resolve(out)));
  fs.writeFileSync(out, md, "utf-8");
  return out;
}
