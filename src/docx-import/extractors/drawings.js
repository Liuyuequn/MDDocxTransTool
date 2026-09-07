const INLINE_CHILDREN = new Set([
  "wp:extent",
  "wp:effectExtent",
  "wp:docPr",
  "wp:cNvGraphicFramePr",
  "a:graphic",
]);

/** 将浮动 DrawingML 对象改为普通行内对象；位置、环绕和层级信息按约定丢弃。 */
export function downgradeFloatingDrawings(root) {
  walk(root, (element) => {
    if (element.name !== "wp:anchor") return;
    element.name = "wp:inline";
    element.attributes = pickDistanceAttributes(element.attributes);
    element.elements = (element.elements || []).filter((child) =>
      child.type !== "element" || INLINE_CHILDREN.has(child.name)
    );
  });
}

function pickDistanceAttributes(attributes = {}) {
  const result = {};
  for (const key of ["distT", "distB", "distL", "distR"]) {
    if (attributes[key] != null) result[key] = attributes[key];
  }
  return result;
}

function walk(node, visit) {
  if (node?.type === "element") visit(node);
  for (const child of node?.elements || []) walk(child, visit);
}
