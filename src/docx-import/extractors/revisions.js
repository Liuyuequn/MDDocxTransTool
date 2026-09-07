import {
  annotationMarkerElement,
  nodeText,
  revisionComment,
} from "../document-model.js";

const ACCEPTED_REVISIONS = new Set(["w:ins", "w:moveTo"]);
const REJECTED_REVISIONS = new Set(["w:del", "w:moveFrom"]);
const REVISION_ELEMENTS = new Set([...ACCEPTED_REVISIONS, ...REJECTED_REVISIONS]);

/**
 * 接受正文内容修订：插入/移入内容保留，删除/移出内容移除。
 * 相邻修订元素视为一次修订，并在修订后内容后插入原文说明占位符。
 */
export function acceptRevisions(root, annotations) {
  transformChildren(root, annotations);
}

function transformChildren(parent, annotations) {
  if (!Array.isArray(parent?.elements)) return;
  const output = [];

  for (let i = 0; i < parent.elements.length;) {
    const child = parent.elements[i];
    if (child?.type !== "element" || !REVISION_ELEMENTS.has(child.name)) {
      transformChildren(child, annotations);
      output.push(child);
      i++;
      continue;
    }

    let beforeText = "";
    let acceptedText = "";
    const accepted = [];
    while (i < parent.elements.length) {
      const revision = parent.elements[i];
      if (revision?.type !== "element" || !REVISION_ELEMENTS.has(revision.name)) break;
      if (REJECTED_REVISIONS.has(revision.name)) {
        beforeText += nodeText(revision);
      } else {
        acceptedText += nodeText(revision);
        for (const nested of revision.elements || []) {
          transformChildren(nested, annotations);
          accepted.push(nested);
        }
      }
      i++;
    }

    output.push(...accepted);
    if (beforeText || acceptedText) {
      const marker = annotations.add(revisionComment(beforeText));
      output.push(annotationMarkerElement(parent.name, marker));
    }
  }

  parent.elements = output;
}
