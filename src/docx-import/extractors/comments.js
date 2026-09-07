import {
  annotationMarkerElement,
  commentAnnotation,
  nodeText,
  normalizedText,
} from "../document-model.js";

/** 从 comments.xml 建立 comment id → 正文映射。 */
export function readComments(commentsDocument) {
  const comments = new Map();
  walk(commentsDocument, (element) => {
    if (element.name !== "w:comment") return;
    const id = element.attributes?.["w:id"];
    if (id != null) comments.set(String(id), normalizedText(nodeText(element)));
  });
  return comments;
}

/**
 * 提取批注范围，在范围结束位置插入 Markdown 注释占位符，并移除 Word 的批注标记。
 * 跨段和嵌套范围通过 active map 按文档顺序累计可见文本。
 */
export function attachCommentAnnotations(root, comments, annotations) {
  const active = new Map();
  const emitted = new Set();
  processElement(root, comments, annotations, active, emitted);
}

function processElement(parent, comments, annotations, active, emitted) {
  if (!Array.isArray(parent?.elements)) return;
  const output = [];

  for (const child of parent.elements) {
    if (child?.type !== "element") {
      output.push(child);
      continue;
    }

    if (child.name === "w:commentRangeStart") {
      const id = String(child.attributes?.["w:id"] ?? "");
      if (id) active.set(id, "");
      continue;
    }

    if (child.name === "w:commentRangeEnd") {
      const id = String(child.attributes?.["w:id"] ?? "");
      if (id && !emitted.has(id)) {
        const marker = annotations.add(commentAnnotation(active.get(id) ?? "", comments.get(id) ?? ""));
        output.push(annotationMarkerElement(parent.name, marker));
        emitted.add(id);
      }
      active.delete(id);
      continue;
    }

    if (child.name === "w:r") {
      const referenceIds = collectReferenceIds(child);
      processElement(child, comments, annotations, active, emitted);
      removeCommentReferences(child);
      if (hasVisibleRunContent(child)) output.push(child);
      for (const id of referenceIds) {
        if (emitted.has(id)) continue;
        const marker = annotations.add(commentAnnotation("", comments.get(id) ?? ""));
        output.push(annotationMarkerElement(parent.name, marker));
        emitted.add(id);
      }
      continue;
    }

    if (child.name === "w:t" || child.name === "w:delText") {
      const text = nodeText(child);
      if (!annotations.isMarker(text)) appendToActive(active, text);
      output.push(child);
      continue;
    }

    if (child.name === "w:tab") appendToActive(active, "\t");
    if (child.name === "w:br" || child.name === "w:cr") appendToActive(active, "\n");
    processElement(child, comments, annotations, active, emitted);
    output.push(child);
  }

  parent.elements = output;
  if (parent.name === "w:p" && active.size) appendToActive(active, "\n");
}

function appendToActive(active, text) {
  for (const [id, value] of active) active.set(id, value + text);
}

function collectReferenceIds(node, out = []) {
  if (node?.type === "element" && node.name === "w:commentReference") {
    const id = node.attributes?.["w:id"];
    if (id != null) out.push(String(id));
  }
  for (const child of node?.elements || []) collectReferenceIds(child, out);
  return out;
}

function removeCommentReferences(node) {
  if (!Array.isArray(node?.elements)) return;
  node.elements = node.elements.filter((child) => child?.name !== "w:commentReference");
  for (const child of node.elements) removeCommentReferences(child);
}

function hasVisibleRunContent(run) {
  return (run.elements || []).some((child) =>
    child.type === "element" && !["w:rPr", "w:commentReference"].includes(child.name)
  );
}

function walk(node, visit) {
  if (node?.type === "element") visit(node);
  for (const child of node?.elements || []) walk(child, visit);
}
