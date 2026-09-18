// Markdown 解析器配置
//
// markdown-it 15 的 newline 规则在生成 token 时**忽略** breaks 选项：
//   dist/markdown-it.mjs -> function newline(state, silent) { ... state.push("softbreak", "br", 0); }
// 该选项只在 HTML 渲染器里生效（default_rules.softbreak 返回 <br> 或 "\n"）。
// 本工具的转换链路是「token → docx」，softbreak 原先被固定渲染为空格，
// 因此中文文档里每行一个换行符的写法会被整段并成一行，换行丢失。
//
// 这里用自定义 newline 规则修正：breaks 开启时，单个换行符同样产生 hardbreak。

import MarkdownIt from "markdown-it";

const NEWLINE = 10;
const SPACE = 32;
const isSpace = (code) => code === SPACE || (code >= 9 && code <= 13);

/** breaks 开启时的换行规则：单个换行符产出 hardbreak（替代 markdown-it 的 softbreak） */
function newlineWithBreaks(state, silent) {
  let pos = state.pos;
  if (state.src.charCodeAt(pos) !== NEWLINE) return false;

  if (!silent) {
    const pmax = state.pending.length - 1;
    // 行尾两个及以上空格：按 CommonMark 规则去掉行尾空格后换行
    if (pmax >= 0 && state.pending.charCodeAt(pmax) === SPACE) {
      let ws = pmax;
      while (ws >= 1 && state.pending.charCodeAt(ws - 1) === SPACE) ws--;
      state.pending = state.pending.slice(0, ws);
    }
    state.push("hardbreak", "br", 0);
  }

  pos++;
  const max = state.posMax;
  while (pos < max && isSpace(state.src.charCodeAt(pos))) pos++;
  state.pos = pos;
  return true;
}

/** 按 breaks 配置创建解析器；breaks 关闭时使用 markdown-it 默认规则 */
export function createMarkdownParser(breaks) {
  const md = new MarkdownIt({ html: false, linkify: true, typographer: true });
  if (breaks === false) return md;
  try {
    md.inline.ruler.at("newline", newlineWithBreaks);
  } catch {
    // 规则名变化时退回 markdown-it 默认行为（仅软换行按空格处理）
  }
  return md;
}
