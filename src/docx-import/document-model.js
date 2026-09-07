import crypto from "node:crypto";

/** 创建一次转换独享的注释占位符仓库，避免与正文中的普通文本冲突。 */
export function createAnnotationStore() {
  const prefix = `MDTTANNOTATION${crypto.randomBytes(8).toString("hex").toUpperCase()}`;
  const entries = [];

  return {
    add(comment) {
      const marker = `${prefix}${entries.length}END`;
      entries.push({ marker, comment });
      return marker;
    },
    isMarker(text) {
      return String(text).startsWith(prefix);
    },
    restore(markdown) {
      let result = markdown;
      for (const { marker, comment } of entries) {
        result = result.replaceAll(marker, comment);
      }
      return result;
    },
  };
}

/** xml-js 非 compact 模式下的元素构造辅助。 */
export function xmlElement(name, elements = [], attributes = undefined) {
  return {
    type: "element",
    name,
    ...(attributes ? { attributes } : {}),
    ...(elements.length ? { elements } : {}),
  };
}

/** 根据父节点允许的内容类型创建占位符；段落中放 run，块容器中放 paragraph。 */
export function annotationMarkerElement(parentName, marker) {
  const run = xmlElement("w:r", [
    xmlElement("w:t", [{ type: "text", text: marker }], { "xml:space": "preserve" }),
  ]);
  return parentName === "w:p" || parentName === "w:hyperlink"
    ? run
    : xmlElement("w:p", [run]);
}

/** 读取节点中的可见文本；用于修订原文、批注范围和批注正文。 */
export function nodeText(node) {
  if (!node) return "";
  if (node.type === "text") return node.text ?? "";
  if (node.type !== "element") return "";
  if (node.name === "w:tab") return "\t";
  if (node.name === "w:br" || node.name === "w:cr") return "\n";
  return (node.elements || []).map(nodeText).join("");
}

export function normalizedText(text) {
  return String(text ?? "").replace(/\s+/g, " ").trim();
}

/** HTML 注释不能包含连续双连字符。 */
export function markdownComment(text) {
  return `<!-- ${normalizedText(text).replace(/--/g, "——")} -->`;
}

export function revisionComment(beforeText) {
  const before = normalizedText(beforeText) || "（无）";
  return markdownComment(`此处系修订；修订前原文：${before}`);
}

/** 超过 20 个 Unicode 字符时只保留首尾各 10 个字符。 */
export function summarizeCommentRange(text) {
  const chars = Array.from(normalizedText(text));
  if (chars.length <= 20) return chars.join("");
  return `${chars.slice(0, 10).join("")}……${chars.slice(-10).join("")}`;
}

export function commentAnnotation(rangeText, commentText) {
  const range = summarizeCommentRange(rangeText) || "（无选定范围）";
  const comment = normalizedText(commentText) || "（无内容）";
  return markdownComment(`批注范围：“${range}”；此处有批注：${comment}`);
}
