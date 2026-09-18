import fs from "node:fs";
import JSZip from "jszip";
import { js2xml, xml2js } from "xml-js";
import { createAnnotationStore } from "./document-model.js";
import { attachCommentAnnotations, readComments } from "./extractors/comments.js";
import { downgradeFloatingDrawings } from "./extractors/drawings.js";
import { acceptRevisions } from "./extractors/revisions.js";
import { linearizeSections } from "./extractors/sections.js";
import { flattenTextBoxes } from "./extractors/textboxes.js";

/**
 * 在交给 mammoth 之前规范化 docx：接受修订、插入批注说明、展平文本框、
 * 降级浮动 DrawingML，并将多栏/多节内容按源顺序线性排列。
 * 接受 Buffer / Uint8Array。
 */
export async function prepareDocxBufferForMarkdown(input) {
  const zip = await JSZip.loadAsync(input);
  const documentFile = zip.file("word/document.xml");
  if (!documentFile) throw new Error("不是有效的 docx 文件（缺少 word/document.xml）");

  const documentXml = await documentFile.async("string");
  const document = xml2js(documentXml, { compact: false });
  const body = findFirstElement(document, "w:body");
  if (!body) throw new Error("不是有效的 docx 文件（document.xml 缺少 w:body）");
  const annotations = createAnnotationStore();

  acceptRevisions(body, annotations);

  const commentsFile = zip.file("word/comments.xml");
  const commentsDocument = commentsFile
    ? xml2js(await commentsFile.async("string"), { compact: false })
    : null;
  attachCommentAnnotations(body, readComments(commentsDocument), annotations);

  flattenTextBoxes(body);
  downgradeFloatingDrawings(body);
  linearizeSections(body);

  zip.file("word/document.xml", js2xml(document, { compact: false, spaces: 0 }));
  const buffer = await zip.generateAsync({ type: "nodebuffer" });
  return { buffer, annotations };
}

/** 从 docx 文件读取并规范化（prepareDocxBufferForMarkdown 的路径版） */
export async function prepareDocxForMarkdown(docxPath) {
  return prepareDocxBufferForMarkdown(fs.readFileSync(docxPath));
}

function findFirstElement(node, name) {
  if (node?.type === "element" && node.name === name) return node;
  for (const child of node?.elements || []) {
    const found = findFirstElement(child, name);
    if (found) return found;
  }
  return null;
}
