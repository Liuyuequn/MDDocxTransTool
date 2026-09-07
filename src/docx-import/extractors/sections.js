/**
 * 移除分节属性和显式分栏符，仅保留源文档中的内容顺序。
 * 页面、分栏、各节页眉页脚等复杂版式按约定不映射到 Markdown。
 */
export function linearizeSections(root) {
  prune(root);
}

function prune(node) {
  if (!Array.isArray(node?.elements)) return;
  node.elements = node.elements.filter((child) => {
    if (child?.type !== "element") return true;
    if (child.name === "w:sectPr") return false;
    if (child.name === "w:br" && child.attributes?.["w:type"] === "column") return false;
    return true;
  });
  for (const child of node.elements) prune(child);
}
