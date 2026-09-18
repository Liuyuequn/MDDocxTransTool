// 单元测试：纯函数层（单位换算、字号解析、深合并、参数解析、docx 导入注释）
// 运行：node test/unit-test.mjs（npm test 会先跑本文件再跑端到端校验）

import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import MarkdownIt from "markdown-it";
import { Document, Packer } from "docx";
import JSZip from "jszip";
import { parseBlocks } from "../src/blocks.js";
import { presets } from "../src/presets.js";
import { createMarkdownParser } from "../src/markdown.js";
import { convertMarkdownFile } from "../src/converter.js";
import {
  cmToTwip,
  ptToTwip,
  ptToHalfPoint,
  parseFontSize,
  mergeOptions,
  PAGE_SIZES,
  defaultOptions,
} from "../src/options.js";
import { argsHelpText, parseArgs, parseMargin, parseHeadingSize } from "../src/args.js";
import { validPresetName } from "../src/preset-extract.js";
import { maxImageWidthPx } from "../src/inline.js";
import { linesToTwip } from "../src/styles.js";
import {
  commentAnnotation,
  revisionComment,
  summarizeCommentRange,
} from "../src/docx-import/document-model.js";

async function markdownParagraphs(source, opts = defaultOptions, md = new MarkdownIt()) {
  const tokens = md.parse(source, {});
  const ctx = { opts, basePath: process.cwd(), listLevel: -1, quote: false };
  const children = parseBlocks(tokens, 0, tokens.length, ctx, { orderedInstance: 0 });
  const zip = await JSZip.loadAsync(await Packer.toBuffer(new Document({ sections: [{ children }] })));
  const xml = await zip.file("word/document.xml").async("string");
  return xml.match(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g) || [];
}

/** 走真实解析器（含 breaks 处理），用于校验换行识别 */
async function paragraphsFromSource(source, opts = defaultOptions) {
  return markdownParagraphs(source, opts, createMarkdownParser(opts.markdown.breaks));
}

/** 端到端：按配置写出真实 docx 再读回 document.xml（覆盖 converter 的 breaks 接线） */
async function documentXmlFor(source, opts = defaultOptions) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "mddtt-unit-"));
  const mdPath = path.join(dir, "in.md");
  const docxPath = path.join(dir, "out.docx");
  try {
    fs.writeFileSync(mdPath, source, "utf-8");
    await convertMarkdownFile(mdPath, docxPath, opts);
    const zip = await JSZip.loadAsync(fs.readFileSync(docxPath));
    return await zip.file("word/document.xml").async("string");
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

for (const newline of ["\n", "\r\n"]) {
  test(`Markdown 换行 ${JSON.stringify(newline)}：两空格保留段内换行，空行分段`, async () => {
    const paragraphs = await markdownParagraphs(`第一句。  ${newline}**第二句。**${newline}${newline}第三句。`);
    assert.equal(paragraphs.length, 2);
    assert.equal((paragraphs[0].match(/<w:br\/>/g) || []).length, 1);
    assert.match(paragraphs[0], /第一句。[\s\S]*<w:br\/>[\s\S]*第二句。/);
    assert.match(paragraphs[0], /<w:b\/>/);
    assert.match(paragraphs[0], /w:line="307" w:lineRule="auto"/);
    // 1.5 行段后距按四号正文的行高换算：14 × 1.28 × 1.5 × 20 = 538 twip。
    for (const paragraph of paragraphs) assert.match(paragraph, /w:after="538"/);
    assert.match(paragraphs[1], /第三句。/);
  });
}

test("Markdown 单个普通换行在 --no-breaks 下按软换行处理", async () => {
  const noBreaks = mergeOptions(defaultOptions, { markdown: { breaks: false } });
  const paragraphs = await paragraphsFromSource("第一句。\n第二句。", noBreaks);
  assert.equal(paragraphs.length, 1);
  assert.doesNotMatch(paragraphs[0], /<w:br/);
  assert.match(paragraphs[0], /第一句。[\s\S]*第二句。/);
});

// ============ markdown.js：换行识别（breaks） ============

test("createMarkdownParser: 默认 breaks 时单个换行符产出 hardbreak", () => {
  const md = createMarkdownParser(true);
  const tokens = md.parse("甲行\n乙行", {});
  const inline = tokens.find((t) => t.type === "inline");
  assert.deepEqual(inline.children.map((t) => t.type), ["text", "hardbreak", "text"]);
});

test("createMarkdownParser: breaks 关闭时单个换行符仍为 softbreak", () => {
  const md = createMarkdownParser(false);
  const tokens = md.parse("甲行\n乙行", {});
  const inline = tokens.find((t) => t.type === "inline");
  assert.deepEqual(inline.children.map((t) => t.type), ["text", "softbreak", "text"]);
});

test("createMarkdownParser: breaks 开启时双空格换行与行尾反斜杠同为 hardbreak", () => {
  const md = createMarkdownParser(true);
  for (const src of ["甲行  \n乙行", "甲行\\\n乙行"]) {
    const inline = md.parse(src, {}).find((t) => t.type === "inline");
    assert.deepEqual(inline.children.map((t) => t.type), ["text", "hardbreak", "text"]);
  }
});

test("createMarkdownParser: breaks 开启时空行仍分段", () => {
  const md = createMarkdownParser(true);
  const tokens = md.parse("甲段\n\n乙段", {});
  assert.equal(tokens.filter((t) => t.type === "inline").length, 2);
});

test("createMarkdownParser: breaks 不改变代码块内容", () => {
  const md = createMarkdownParser(true);
  const tokens = md.parse("```\n甲行\n乙行\n```", {});
  const fence = tokens.find((t) => t.type === "fence");
  assert.equal(fence.content, "甲行\n乙行\n");
});

test("normalizeBlankLines: 换行识别不改变代码块内容", () => {
  const md = createMarkdownParser(true);
  const tokens = md.parse("```\ncode  \n\ncode2\n```", {});
  const fence = tokens.find((t) => t.type === "fence");
  assert.equal(fence.content, "code  \n\ncode2\n");
});

test("defaultOptions: 默认启用换行识别（breaks）", () => {
  assert.equal(defaultOptions.markdown.breaks, true);
});

test("convertMarkdownFile: 默认配置下单个换行渲染为段内换行（w:br）", async () => {
  const xml = await documentXmlFor("甲行\n乙行\n\n新段落\n");
  const paragraphs = xml.match(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g) || [];
  assert.equal(paragraphs.length, 2);
  assert.equal((paragraphs[0].match(/<w:br\/>/g) || []).length, 1);
  assert.match(paragraphs[0], /甲行[\s\S]*<w:br\/>[\s\S]*乙行/);
  assert.doesNotMatch(paragraphs[1], /<w:br\/>/);
});

test("convertMarkdownFile: --no-breaks 配置下单换行合并为同行", async () => {
  const noBreaks = mergeOptions(defaultOptions, { markdown: { breaks: false } });
  const xml = await documentXmlFor("甲行\n乙行\n", noBreaks);
  assert.doesNotMatch(xml, /<w:br\/>/);
  assert.match(xml, /甲行[\s\S]*乙行/);
});

// ============ converter.js：settings.xml 兼容性设置 ============

/** 读取生成文档的 settings.xml */
async function settingsXmlFor(source, opts = defaultOptions) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "mddtt-set-"));
  const mdPath = path.join(dir, "in.md");
  const docxPath = path.join(dir, "out.docx");
  try {
    fs.writeFileSync(mdPath, source, "utf-8");
    await convertMarkdownFile(mdPath, docxPath, opts);
    const zip = await JSZip.loadAsync(fs.readFileSync(docxPath));
    return await zip.file("word/settings.xml").async("string");
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

test("convertMarkdownFile: 换行后补两字符缩进（全角空格×2）", async () => {
  const xml = await documentXmlFor("甲行\n乙行\n");
  // 换行符后紧接两个全角空格（U+3000），使段内后续各行与首行左端对齐
  assert.match(xml, /<w:br\/><\/w:r><w:r><w:t xml:space="preserve">\u3000\u3000<\/w:t><\/w:r><w:r><w:t[^>]*>乙行/);
  assert.equal((xml.match(/\u3000\u3000/g) || []).length, 1);
});

test("convertMarkdownFile: 换行缩进不重复叠加（续行原有前导空格被忽略）", async () => {
  const xml = await documentXmlFor("甲行  \n   乙行\n");
  assert.equal((xml.match(/\u3000/g) || []).length, 2);
});

test("convertMarkdownFile: --no-breaks 时单换行不产生缩进", async () => {
  const noBreaks = mergeOptions(defaultOptions, { markdown: { breaks: false } });
  const xml = await documentXmlFor("甲行\n乙行\n", noBreaks);
  assert.doesNotMatch(xml, /\u3000/);
});

test("convertMarkdownFile: 换行后的行内代码同样获得缩进", async () => {
  const xml = await documentXmlFor("甲行\n`code` 后文\n");
  assert.match(xml, /<w:br\/><\/w:r><w:r><w:t xml:space="preserve">\u3000\u3000<\/w:t><\/w:r>/);
});

test("settings.xml: 写入 doNotExpandShiftReturn（不拉伸换行符结尾的行）", async () => {
  const settings = await settingsXmlFor("甲行\n乙行\n");
  assert.match(settings, /<w:doNotExpandShiftReturn\/>/);
});

test("settings.xml: compatSetting 排在 doNotExpandShiftReturn 之后（符合 CT_Compat 顺序）", async () => {
  const settings = await settingsXmlFor("甲行\n乙行\n");
  const compat = /<w:compat>[\s\S]*?<\/w:compat>/.exec(settings);
  assert.ok(compat, "settings.xml 应包含 w:compat");
  const flagsAt = compat[0].indexOf("<w:doNotExpandShiftReturn/>");
  const settingAt = compat[0].indexOf("<w:compatSetting");
  assert.ok(flagsAt >= 0 && settingAt >= 0, "两个元素都应存在");
  assert.ok(flagsAt < settingAt, "compatSetting 必须位于兼容性开关之后");
});

test("sundy 混合换行：段内 1.28 倍，段后 1.5 行，不插入空段落", async () => {
  const paragraphs = await markdownParagraphs("甲。  \n乙。\n\n丙。  \n丁。", mergeOptions(defaultOptions, presets.sundy));
  assert.equal(paragraphs.length, 2);
  for (const paragraph of paragraphs) {
    assert.match(paragraph, /w:after="538"/);
    assert.match(paragraph, /w:line="307" w:lineRule="auto"/);
    assert.equal((paragraph.match(/<w:br\/>/g) || []).length, 1);
  }
});

test("显式行距和段后距参数仍可覆盖默认换行排版", async () => {
  const { patch } = parseArgs(["--line-height", "2", "--para-spacing", "8"]);
  const paragraphs = await markdownParagraphs("甲。  \n乙。\n\n丙。", mergeOptions(defaultOptions, patch));
  for (const paragraph of paragraphs) {
    assert.match(paragraph, /w:after="160"/);
    assert.match(paragraph, /w:line="480" w:lineRule="auto"/);
  }
});

// ============ preset-extract.js：预设名校验 ============

test("validPresetName: 中英文/数字/下划线/连字符合法", () => {
  assert.equal(validPresetName("firm"), true);
  assert.equal(validPresetName("律所模板1"), true);
  assert.equal(validPresetName("my-preset_2"), true);
});

test("validPresetName: 拒绝路径分隔符、空格与空值", () => {
  assert.equal(validPresetName("a/b"), false);
  assert.equal(validPresetName("a b"), false);
  assert.equal(validPresetName(".."), false);
  assert.equal(validPresetName("-x"), false); // 不以连字符开头
  assert.equal(validPresetName(""), false);
  assert.equal(validPresetName(null), false);
});

// ============ options.js：单位换算 ============

test("cmToTwip: 2.54cm = 1440 twip（1 英寸）", () => {
  assert.equal(cmToTwip(2.54), 1440);
});

test("cmToTwip: 0 = 0", () => {
  assert.equal(cmToTwip(0), 0);
});

test("cmToTwip: 3.18cm 四舍五入为 1803", () => {
  assert.equal(cmToTwip(3.18), 1803);
});

test("ptToTwip: 1pt = 20 twip", () => {
  assert.equal(ptToTwip(12), 240);
});

test("ptToHalfPoint: 四号 14pt = 28 半磅", () => {
  assert.equal(ptToHalfPoint(14), 28);
});

test("ptToHalfPoint: 五号 10.5pt = 21 半磅", () => {
  assert.equal(ptToHalfPoint(10.5), 21);
});

// ============ options.js：字号解析 ============

test("parseFontSize: 数字字符串解析为 pt", () => {
  assert.equal(parseFontSize("14"), 14);
});

test("parseFontSize: 纯数字直接返回", () => {
  assert.equal(parseFontSize(12), 12);
});

test("parseFontSize: 中文字号名映射（四号=14）", () => {
  assert.equal(parseFontSize("四号"), 14);
});

test("parseFontSize: 中文字号名映射（五号=10.5）", () => {
  assert.equal(parseFontSize("五号"), 10.5);
});

test("parseFontSize: 容忍前后空白", () => {
  assert.equal(parseFontSize("  小五 "), 9);
});

test("parseFontSize: 无法识别时抛错", () => {
  assert.throws(() => parseFontSize("胡说"), /无法识别的字号/);
});

// ============ options.js：深合并 ============

test("mergeOptions: 嵌套对象递归合并", () => {
  const base = { page: { size: "A4", margin: { top: 1, left: 1 } } };
  const patch = { page: { margin: { top: 2 } } };
  const merged = mergeOptions(base, patch);
  assert.deepEqual(merged, { page: { size: "A4", margin: { top: 2, left: 1 } } });
});

test("mergeOptions: 数组整体覆盖而非逐项合并", () => {
  const base = { sizes: { heading: [22, 16, 14, 12, 11, 11] } };
  const patch = { sizes: { heading: [22, 16, 14, 14, 14, 14] } };
  const merged = mergeOptions(base, patch);
  assert.deepEqual(merged.sizes.heading, [22, 16, 14, 14, 14, 14]);
});

test("mergeOptions: undefined 字段不覆盖已有值", () => {
  const base = { header: { text: "A" } };
  const patch = { header: { text: undefined, align: "left" } };
  const merged = mergeOptions(base, patch);
  assert.equal(merged.header.text, "A");
  assert.equal(merged.header.align, "left");
});

test("mergeOptions: 多个补丁按顺序叠加（后者胜）", () => {
  const base = { fonts: { body: { eastAsia: "等线" } } };
  const p1 = { fonts: { body: { eastAsia: "仿宋" } } };
  const p2 = { fonts: { body: { eastAsia: "黑体" } } };
  assert.equal(mergeOptions(base, p1, p2).fonts.body.eastAsia, "黑体");
});

test("mergeOptions: null 值直接覆盖（用于显式清空语义）", () => {
  const base = { header: { text: "A" } };
  const patch = { header: { text: null } };
  assert.equal(mergeOptions(base, patch).header.text, null);
});

// ============ args.js：parseMargin ============

test("parseMargin: 单值四边统一", () => {
  const patch = { page: {} };
  parseMargin("2.5", patch);
  assert.deepEqual(patch.page.margin, { top: 2.5, right: 2.5, bottom: 2.5, left: 2.5 });
});

test("parseMargin: 四值按 上,右,下,左 顺序", () => {
  const patch = { page: {} };
  parseMargin("2.54,3.18,2.54,3.18", patch);
  assert.deepEqual(patch.page.margin, { top: 2.54, right: 3.18, bottom: 2.54, left: 3.18 });
});

test("parseMargin: 容忍数值间空白", () => {
  const patch = { page: {} };
  parseMargin("1, 2, 3, 4", patch);
  assert.deepEqual(patch.page.margin, { top: 1, right: 2, bottom: 3, left: 4 });
});

test("parseMargin: 非数字抛错", () => {
  assert.throws(() => parseMargin("abc", { page: {} }), /页边距格式错误/);
});

test("parseMargin: 负数抛错", () => {
  assert.throws(() => parseMargin("-1", { page: {} }), /页边距格式错误/);
});

test("parseMargin: 三个值抛错（应为 1 或 4 个）", () => {
  assert.throws(() => parseMargin("1,2,3", { page: {} }), /1 个或 4 个/);
});

// ============ args.js：parseHeadingSize ============

test("parseHeadingSize: auto 不写入字号", () => {
  const patch = { heading: {}, sizes: {} };
  parseHeadingSize("auto", patch);
  assert.equal(patch.sizes.heading, undefined);
});

test("parseHeadingSize: 单值扩展为六级相同", () => {
  const patch = { heading: {}, sizes: {} };
  parseHeadingSize("22", patch);
  assert.deepEqual(patch.sizes.heading, [22, 22, 22, 22, 22, 22]);
});

test("parseHeadingSize: 六级逗号分隔原样生效", () => {
  const patch = { heading: {}, sizes: {} };
  parseHeadingSize("22,16,14,14,14,14", patch);
  assert.deepEqual(patch.sizes.heading, [22, 16, 14, 14, 14, 14]);
});

test("parseHeadingSize: 中文字号名可用于标题", () => {
  const patch = { heading: {}, sizes: {} };
  parseHeadingSize("二号", patch);
  assert.deepEqual(patch.sizes.heading, [22, 22, 22, 22, 22, 22]);
});

test("parseHeadingSize: 错误级数抛错", () => {
  assert.throws(() => parseHeadingSize("1,2,3", { heading: {}, sizes: {} }), /单一值或逗号分隔的六级值/);
});

// ============ args.js：parseArgs ============

test("argsHelpText: 参数列表以对齐表格输出", () => {
  const help = argsHelpText();
  assert.match(help, /^┌─+┬─+┐\n│参数\s+│说明\s+│\n├─+┼─+┤/u);
  assert.match(help, /│--page-size, -s <值>\s+│A4/u);
  assert.match(help, /│--save-preset <值>\s+│将 docx/u);
  assert.match(help, /└─+┴─+┘$/u);
});

test("parseArgs: 别名与全名等价（-f 与 --font）", () => {
  const a = parseArgs(["-f", "宋体"]);
  const b = parseArgs(["--font", "宋体"]);
  assert.deepEqual(a.patch.fonts.body, b.patch.fonts.body);
  assert.equal(a.patch.fonts.body.eastAsia, "宋体");
});

test("parseArgs: --name=value 等价于 --name value", () => {
  const a = parseArgs(["--font=宋体"]);
  const b = parseArgs(["--font", "宋体"]);
  assert.deepEqual(a.patch, b.patch);
});

test("parseArgs: flag 参数无需取值", () => {
  const r = parseArgs(["--no-bold"]);
  assert.deepEqual(r.patch.heading.bold, [false, false, false, false, false, false]);
});

test("parseArgs: 输出控制参数单独归位（output/overwrite/preset）", () => {
  const r = parseArgs(["-o", "out.docx", "--overwrite", "--preset", "sundy"]);
  assert.equal(r.output, "out.docx");
  assert.equal(r.overwrite, true);
  assert.equal(r.preset, "sundy");
});

test("parseArgs: 未知参数记入 errors", () => {
  const r = parseArgs(["--no-such-param"]);
  assert.equal(r.errors.length, 1);
  assert.match(r.errors[0], /未知参数/);
});

test("parseArgs: 参数缺少取值记入 errors", () => {
  const r = parseArgs(["--font"]);
  assert.equal(r.errors.length, 1);
  assert.match(r.errors[0], /缺少取值/);
});

test("parseArgs: 取值非法记入 errors（校验函数）", () => {
  const r = parseArgs(["--orientation", "diagonal"]);
  assert.equal(r.errors.length, 1);
  assert.match(r.errors[0], /portrait 或 landscape/);
});

test("parseArgs: 互斥页眉参数记入 errors", () => {
  const r = parseArgs(["--header", "A", "--header-left", "B"]);
  assert.ok(r.errors.some((e) => e.includes("互斥")));
});

test("parseArgs: --total-pages 在无 Y 模板上追加 /共Y页", () => {
  const r = parseArgs(["-p", "bottom", "--total-pages"]);
  assert.equal(r.patch.pageNumber.format, "X/共Y页");
});

test("parseArgs: --total-pages 不改动已含 Y 的模板", () => {
  const r = parseArgs(["--page-num-format", "第X页/共Y页", "--total-pages"]);
  assert.equal(r.patch.pageNumber.format, "第X页/共Y页");
});

test("parseArgs: 非选项 token 记为多余位置参数", () => {
  const r = parseArgs(["to"]);
  assert.equal(r.errors.length, 1);
  assert.match(r.errors[0], /多余的位置参数/);
});

test("parseArgs: --line-rule 与 --line-height 组合", () => {
  const r = parseArgs(["--line-rule", "exact", "--line-height", "20"]);
  assert.equal(r.patch.paragraph.lineRule, "exact");
  assert.equal(r.patch.paragraph.line, 20);
});

test("parseArgs: 非法 --line-rule 记入 errors", () => {
  const r = parseArgs(["--line-rule", "foo"]);
  assert.equal(r.errors.length, 1);
  assert.match(r.errors[0], /auto、exact 或 atLeast/);
});

// ============ styles.js：行距换算 ============

test("linesToTwip: auto 倍数行距按 字号×倍数 换算", () => {
  assert.equal(linesToTwip(1, 14, { line: 1.5, lineRule: "auto" }), 420); // 14×1.5×20
});

test("linesToTwip: exact 固定行距与字号无关", () => {
  assert.equal(linesToTwip(1, 14, { line: 20, lineRule: "exact" }), 400);
  assert.equal(linesToTwip(1, 22, { line: 20, lineRule: "exact" }), 400);
});

test("linesToTwip: 无 lineCfg 时退回默认 1.28 倍", () => {
  assert.equal(linesToTwip(1, 14, null), Math.round(14 * 1.28 * 20));
});

// ============ inline.js：正文图片宽度上限 ============

test("maxImageWidthPx: 默认 A4 = 内容区宽度(twip)÷15 取整", () => {
  const contentTwip = PAGE_SIZES.A4.width - 2 * cmToTwip(defaultOptions.page.margin.left);
  assert.equal(maxImageWidthPx(defaultOptions), Math.floor(contentTwip / 15));
});

test("maxImageWidthPx: 边距变宽时上限随之缩小", () => {
  const wide = mergeOptions(defaultOptions, { page: { margin: { left: 5, right: 5 } } });
  const contentTwip = PAGE_SIZES.A4.width - 2 * cmToTwip(5);
  assert.equal(maxImageWidthPx(wide), Math.floor(contentTwip / 15));
});

test("maxImageWidthPx: 窄页面上限更小（A5 横向口径）", () => {
  const a5 = mergeOptions(defaultOptions, { page: { size: "A5" } });
  const contentTwip = PAGE_SIZES.A5.width - 2 * cmToTwip(defaultOptions.page.margin.left);
  assert.equal(maxImageWidthPx(a5), Math.floor(contentTwip / 15));
});

// ============ docx-import：修订与批注注释 ============

test("revisionComment: 修订前原文写入 Markdown 注释", () => {
  assert.equal(revisionComment("旧条款"), "<!-- 此处系修订；修订前原文：旧条款 -->");
});

test("summarizeCommentRange: 超过 20 字符时保留首尾各 10 字符", () => {
  assert.equal(
    summarizeCommentRange("ABCDEFGHIJKLMNOPQRSTUVWXYZ"),
    "ABCDEFGHIJ……QRSTUVWXYZ"
  );
});

test("commentAnnotation: 标注范围并包含批注正文", () => {
  assert.equal(
    commentAnnotation("被批注内容", "请核对此条款"),
    "<!-- 批注范围：“被批注内容”；此处有批注：请核对此条款 -->"
  );
});
