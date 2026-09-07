import crypto from "node:crypto";
import { parse as parseHtml } from "node-html-parser";
import TurndownService from "turndown";
import { gfm } from "turndown-plugin-gfm";

const MERGED_TABLE_NOTE = "<!-- 该表格含合并单元格，Markdown 表格无法表达，以 HTML 形式保留 -->";
const MERGE_ATTR_RE = /\b(?:colspan|rowspan)\s*=/i;

/** HTML → Markdown，并回填 OOXML 预处理阶段生成的修订/批注注释。 */
export function renderMarkdown(rawHtml, annotations) {
  const placeholderPrefix = `MDTTHTMLTABLE-${crypto.randomBytes(8).toString("hex")}`;
  const placeholder = (index) => `${placeholderPrefix}-${index}`;
  const { html, htmlTables } = preprocessHtml(rawHtml, placeholder);

  const turndown = createTurndownService();
  let markdown = turndown.turndown(html);

  htmlTables.forEach((tableHtml, index) => {
    markdown = markdown.replace(
      placeholder(index),
      () => `${MERGED_TABLE_NOTE}\n\n${tableHtml}`
    );
  });

  markdown = annotations.restore(markdown);
  return postProcess(markdown);
}

/** 修复 mammoth 表格结构；合并单元格以 HTML 回填，普通表格转 GFM。 */
function preprocessHtml(rawHtml, placeholder) {
  const root = parseHtml(rawHtml);
  const htmlTables = [];
  for (const table of root.querySelectorAll("table")) {
    const rows = directRows(table);
    if (rows.length === 0) continue;
    const hasMerge = rows.some((row) =>
      directCells(row).some((cell) => MERGE_ATTR_RE.test(cell.rawAttrs))
    );
    if (hasMerge) {
      htmlTables.push(buildCleanTable(rows, { keepAttrs: true }));
      table.insertAdjacentHTML("beforebegin", `<p>${placeholder(htmlTables.length - 1)}</p>`);
      table.remove();
    } else {
      table.insertAdjacentHTML("beforebegin", buildCleanTable(rows, { keepAttrs: false }));
      table.remove();
    }
  }

  for (const item of root.querySelectorAll("li")) {
    const meaningful = item.childNodes.filter((node) => node.nodeType === 1 || node.text.trim());
    const onlyParagraph =
      meaningful.length === 1 && meaningful[0].nodeType === 1 && meaningful[0].tagName === "P"
        ? meaningful[0]
        : null;
    if (onlyParagraph) onlyParagraph.replaceWith(...onlyParagraph.childNodes);
  }
  return { html: root.toString(), htmlTables };
}

function directRows(table) {
  const rows = [];
  for (const child of table.childNodes) {
    if (child.nodeType !== 1) continue;
    if (child.tagName === "TR") rows.push(child);
    else if (["THEAD", "TBODY", "TFOOT"].includes(child.tagName)) {
      for (const nested of child.childNodes) {
        if (nested.nodeType === 1 && nested.tagName === "TR") rows.push(nested);
      }
    }
  }
  return rows;
}

function directCells(row) {
  return row.childNodes.filter((node) =>
    node.nodeType === 1 && (node.tagName === "TD" || node.tagName === "TH")
  );
}

function buildCleanTable(rows, { keepAttrs }) {
  const tableRows = rows.map((row, index) => {
    const isHeader = index === 0;
    const cells = directCells(row).map((cell) => {
      const tag = keepAttrs ? cell.tagName.toLowerCase() : isHeader ? "th" : "td";
      const attrs = keepAttrs && cell.rawAttrs ? ` ${cell.rawAttrs}` : "";
      const content = cell.innerHTML.replace(/<\/?p>/g, "").trim();
      return `<${tag}${attrs}>${content}</${tag}>`;
    });
    return `<tr>${cells.join("")}</tr>`;
  });
  const tableHead = `<thead>${tableRows[0]}</thead>`;
  const tableBody = tableRows.length > 1 ? `<tbody>${tableRows.slice(1).join("")}</tbody>` : "";
  return `<table>${tableHead}${tableBody}</table>`;
}

function createTurndownService() {
  const turndown = new TurndownService({
    headingStyle: "atx",
    codeBlockStyle: "fenced",
    bulletListMarker: "-",
    emDelimiter: "*",
    strongDelimiter: "**",
  });
  turndown.addRule("emptyParagraph", {
    filter: (node) =>
      node.nodeName === "P" && !node.textContent.trim() && !node.querySelector("img"),
    replacement: () => "",
  });
  turndown.use(gfm);
  return turndown;
}

function postProcess(markdown) {
  return markdown
    .replace(/\n{3,}/g, "\n\n")
    .replace(/(?<! ) {3,}$/gm, "")
    .replace(/^\n+/, "")
    .replace(/\n*$/, "\n");
}
