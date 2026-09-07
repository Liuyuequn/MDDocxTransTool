/**
 * 将文本框中的段落/表格移动到其锚定块之后，并移除原文本框图形。
 * 这样 mammoth 会按普通文档流线性处理文本框内容，避免内容被形状容器吞掉。
 */
export function flattenTextBoxes(root) {
  flattenContainer(root);
}

function flattenContainer(container) {
  if (!Array.isArray(container?.elements)) return;
  const output = [];

  for (const child of container.elements) {
    if (child?.type !== "element") {
      output.push(child);
      continue;
    }

    const textBoxes = descendants(child, "w:txbxContent");
    if (!textBoxes.length) {
      flattenContainer(child);
      output.push(child);
      continue;
    }

    const extractedBlocks = textBoxes.flatMap((box) =>
      (box.elements || []).filter((element) =>
        element.type === "element" && (element.name === "w:p" || element.name === "w:tbl")
      )
    );

    removeTextBoxHosts(child);
    flattenContainer(child);
    if (!isEmptyParagraph(child)) output.push(child);

    for (const block of extractedBlocks) {
      flattenContainer(block);
      output.push(block);
    }
  }

  container.elements = output;
}

function descendants(node, name, out = []) {
  for (const child of node?.elements || []) {
    if (child?.type !== "element") continue;
    if (child.name === name) out.push(child);
    descendants(child, name, out);
  }
  return out;
}

function containsTextBox(node) {
  return node?.name === "w:txbxContent" || descendants(node, "w:txbxContent").length > 0;
}

function removeTextBoxHosts(node) {
  if (!Array.isArray(node?.elements)) return;
  node.elements = node.elements.filter((child) =>
    !(child?.type === "element" && ["w:drawing", "w:pict"].includes(child.name) && containsTextBox(child))
  );
  for (const child of node.elements) removeTextBoxHosts(child);
}

function isEmptyParagraph(element) {
  if (element.name !== "w:p") return false;
  return !(element.elements || []).some((child) =>
    child.type === "element" && !["w:pPr", "w:bookmarkStart", "w:bookmarkEnd"].includes(child.name)
  );
}
